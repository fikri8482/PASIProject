Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO


Public Class SummaryOutstandingDanDeliveryPasiToAffExp
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
            "if (cboSupplierCode.GetItemCount() > 1) { " & vbCrLf & _
            "   txtSupplierName.SetText('==ALL=='); " & vbCrLf & _
            "   cboSupplierCode.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "var PeriodTo = new Date(); " & vbCrLf & _
            "dtPOPeriodFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtPOPeriodTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtSupplierDelDateFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtSupplierDelDateTo.SetDate(PeriodTo); " & vbCrLf & _
            "chkSupplierDelDate.SetChecked(false); " & vbCrLf & _
            "dtSupplierDelDateFrom.SetEnabled(false); " & vbCrLf & _
            "dtSupplierDelDateTo.SetEnabled(false); " & vbCrLf & _
            " " & vbCrLf & _
            "txtPONo.SetText(''); " & vbCrLf & _
            " " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(chkPOPeriod, chkPOPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            'Dim ls_filter As String = ""
            ls_SQL = ""

            'Dim ls_End As String = ""
            'ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

            ''AFFILIATE CODE
            'If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
            '    ls_filter = ls_filter + _
            '                  "                      AND POM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            'End If

            ''SUPPLIER CODE
            'If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
            '    ls_filter = ls_filter + _
            '                  "                      AND POM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            'End If

            ''AFFILIATE PO PERIOD
            'If chkPOPeriod.Checked = True Then
            '    ls_filter = ls_filter + _
            '                  "                      AND CONVERT(date,POM.Period) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyy-MM-01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyy-MM-" & ls_End) & "' " & vbCrLf
            'End If

            ''SUPPLIER DELIVERY DATE
            'If chkSupplierDelDate.Checked = True Then
            '    ls_filter = ls_filter + _
            '                  "                      AND CONVERT(date,POM.ETDVendor1) BETWEEN '" & Format(dtSupplierDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtSupplierDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
            'End If

            ''PASI DELIVERY DATE
            'If txtPONo.Text <> "" Then
            '    ls_filter = ls_filter + _
            '                  "                      AND POD.PONo = '" & txtPONo.Text & "' " & vbCrLf
            'End If



            'ls_SQL = "   SELECT *   " & vbCrLf & _
            '          "   ,StatusPO = CASE WHEN Remaining > 0 THEN 'OVER'   " & vbCrLf & _
            '          "   WHEN Remaining < 0 THEN 'OPEN'   " & vbCrLf & _
            '          "   ELSE 'CLOSE' END   " & vbCrLf & _
            '          "   FROM (   " & vbCrLf & _
            '          "   SELECT     " & vbCrLf & _
            '          "    	POM.Period    " & vbCrLf & _
            '          "    	,POM.PONo    " & vbCrLf & _
            '          "    	,POM.AffiliateID    " & vbCrLf & _
            '          "    	,POM.SupplierID    " & vbCrLf & _
            '          "    	,POKanban = '' "

            'ls_SQL = ls_SQL + "    	,PASISendAffiliateDate  = Format(POM.EntryDate,'yyyy-MM-dd') " & vbCrLf & _
            '                  "    	,POD.PartNo    " & vbCrLf & _
            '                  "    	,MP.PartName    " & vbCrLf & _
            '                  "    	,QtyPO = ISNULL(POD.TotalPOQty,0) " & vbCrLf & _
            '                  "   	,MPM.QtyBox   " & vbCrLf & _
            '                  "   	,MPM.BoxPallet   " & vbCrLf & _
            '                  "   	,VolumePallet = (POD.TotalPOQty/MPM.QtyBox)/MPM.BoxPallet   " & vbCrLf & _
            '                  "    	,ETDSupp = POM.ETDVendor1    " & vbCrLf & _
            '                  "    	,ETAAff = POM.ETAFactory1 " & vbCrLf & _
            '                  "    	,PASIDeliveryDate = DSM.DeliveryDate " & vbCrLf & _
            '                  "    	,PASISuratJalanNo = COALESCE(DSM.SuratJalanNo,RFM.SuratJalanNo) "

            'ls_SQL = ls_SQL + "  	,PASIInvoiceNo = SD.ShippingInstructionNo " & vbCrLf & _
            '                  "  	,PASIDeliveryQty = SD.ShippingQty   " & vbCrLf & _
            '                  "   	,Remaining = (ISNULL(POD.TotalPOQty,0) -   " & vbCrLf & _
            '                  "       (  " & vbCrLf & _
            '                  "           Select ISNULL(SUM(ShippingQty), 0)   " & vbCrLf & _
            '                  "           FROM ShippingInstruction_Detail RRD   " & vbCrLf & _
            '                  "           LEFT JOIN ShippingInstruction_Master RRM ON RRM.AffiliateID = RRD.AffiliateID   " & vbCrLf & _
            '                  "                 AND RRM.ForwarderID = RRD.ForwarderID " & vbCrLf & _
            '                  "                 AND RRM.ShippingInstructionNo = RRD.ShippingInstructionNo " & vbCrLf & _
            '                  "           WHERE RRD.OrderNo = POD.OrderNo1 " & vbCrLf & _
            '                  "             AND RRD.SupplierID = POD.SupplierID  "

            'ls_SQL = ls_SQL + "             AND RRD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
            '                  "             AND RRD.PartNo = POD.PartNo  " & vbCrLf & _
            '                  "             --AND RRM.ShippingInstructionDate <= RFM.ReceiveDate  " & vbCrLf & _
            '                  "       )) * (-1)  " & vbCrLf & _
            '                  "   FROM PO_Master_Export POM  " & vbCrLf & _
            '                  "   LEFT JOIN PO_Detail_Export POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
            '                  "    						AND POM.PoNo = POD.PONo    " & vbCrLf & _
            '                  "    						AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "  						AND POM.OrderNo1 = POD.OrderNo1  " & vbCrLf & _
            '                  "   LEFT JOIN ReceiveForwarder_Detail RFD ON POD.SupplierID = RFD.SupplierID  " & vbCrLf & _
            '                  "  										AND POD.AffiliateID = RFD.AffiliateID  "

            'ls_SQL = ls_SQL + "  										AND POD.PONo = RFD.PONo  " & vbCrLf & _
            '                  "  										AND POD.OrderNo1 = RFD.OrderNo  " & vbCrLf & _
            '                  "  										AND POD.PartNo = RFD.PartNo  " & vbCrLf & _
            '                  "   LEFT JOIN ReceiveForwarder_Master RFM ON RFD.SuratJalanNo = RFM.SuratJalanNo  " & vbCrLf & _
            '                  "  										AND RFD.SupplierID = RFM.SupplierID  " & vbCrLf & _
            '                  "  										AND RFD.AffiliateID = RFM.AffiliateID  " & vbCrLf & _
            '                  "  										AND RFD.PONo = RFM.PONo  " & vbCrLf & _
            '                  "  										AND RFD.OrderNo = RFM.OrderNo  " & vbCrLf & _
            '                  "   LEFT JOIN DOSupplier_Master_Export DSM ON RFM.suratJalanNo = DSM.SuratJalanNo   " & vbCrLf & _
            '                  "                                          AND RFM.affiliateID = DSM.affiliateID  " & vbCrLf & _
            '                  "                                          AND RFM.SupplierID = DSM.SupplierID   "

            'ls_SQL = ls_SQL + "  										AND RFM.PONo = DSM.PONo  " & vbCrLf & _
            '                  "  										AND RFM.OrderNo = DSM.OrderNo  " & vbCrLf & _
            '                  "   LEFT JOIN DOSupplier_Detail_Export DSD ON DSM.SuratJalanNo = DSD.SuratjalanNo    " & vbCrLf & _
            '                  "                                          AND DSM.AffiliateID = DSD.AffiliateID    " & vbCrLf & _
            '                  "                                          AND DSM.SupplierID = DSD.SupplierID    " & vbCrLf & _
            '                  "                                          AND DSM.PONO = DSD.PONO    " & vbCrLf & _
            '                  "                                          AND RFD.PartNo = DSD.PartNo    " & vbCrLf & _
            '                  "                                          AND RFD.PONO = DSD.PONO  " & vbCrLf & _
            '                  "   LEFT JOIN (Select ShippingInstructionNo,SuratJalanNo,AffiliateID,SupplierID,OrderNo,PartNo,ShippingQty=SUM(ShippingQty) " & vbCrLf & _
            '                  " 			   From ShippingInstruction_Detail " & vbCrLf & _
            '                  " 			 Group By ShippingInstructionNo,SuratJalanNo,AffiliateID,SupplierID,OrderNo,PartNo "

            'ls_SQL = ls_SQL + " 			 ) SD  " & vbCrLf & _
            '                  " 			ON RFD.SuratJalanNo = SD.SuratJalanNo " & vbCrLf & _
            '                  " 			AND RFD.AffiliateID = SD.AffiliateID " & vbCrLf & _
            '                  " 			AND RFD.SupplierID = SD.SupplierID " & vbCrLf & _
            '                  " 			AND RFD.OrderNo = SD.OrderNo " & vbCrLf & _
            '                  " 			AND RFD.PartNo = SD.PartNo " & vbCrLf & _
            '                  "   LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "   LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo AND MPM.AffiliateID = POD.AffiliateID AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
            '                  "   WHERE POD.TotalPOQty > 0 " & vbCrLf



            'ls_SQL = ls_SQL + ls_filter & vbCrLf

            'ls_SQL = ls_SQL + ")A " & vbCrLf & _
            '                  " Order By Period,PONo,AffiliateID,SupplierID,PASISuratJalanNo " & vbCrLf

            ls_SQL = uf_SQL()

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

    Private Function uf_SQL_Cek() As String
        Dim ls_filter As String = ""
        ls_SQL = ""

        Dim ls_End As String = ""
        ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

        'AFFILIATE CODE
        If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
            ls_filter = ls_filter +
                          "                      AND POM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
        End If

        'SUPPLIER CODE
        If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
            ls_filter = ls_filter +
                          "                      AND POM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
        End If

        'AFFILIATE PO PERIOD
        If chkPOPeriod.Checked = True Then
            ls_filter = ls_filter +
                          "                      AND CONVERT(date,POM.Period) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyy-MM-01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyy-MM-" & ls_End) & "' " & vbCrLf
        End If

        'SUPPLIER DELIVERY DATE
        If chkSupplierDelDate.Checked = True Then
            ls_filter = ls_filter +
                          "                      AND CONVERT(date,POM.ETDVendor1) BETWEEN '" & Format(dtSupplierDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtSupplierDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
        End If

        'PASI DELIVERY DATE
        If txtPONo.Text <> "" Then
            ls_filter = ls_filter +
                          "                      AND POD.PONo = '" & txtPONo.Text & "' " & vbCrLf
        End If



        ls_SQL = "   SELECT *   " & vbCrLf &
                  "   ,StatusPO = CASE WHEN Remaining > 0 THEN 'OVER'   " & vbCrLf &
                  "   WHEN Remaining < 0 THEN 'OPEN'   " & vbCrLf &
                  "   ELSE 'CLOSE' END   " & vbCrLf &
                  "   FROM (   " & vbCrLf &
                  "   SELECT  Distinct   " & vbCrLf &
                  "    	POM.Period    " & vbCrLf &
                  "    	,PONo=TRIM(POM.PONo)" & vbCrLf &
                  "    	,OrderNo=TRIM(POM.OrderNo1)" & vbCrLf &
                  "    	,AffiliateID=TRIM(POM.AffiliateID)" & vbCrLf &
                  "    	,SupplierID=TRIM(POM.SupplierID)" & vbCrLf &
                  "    	--,POKanban = '' "

        ls_SQL = ls_SQL + "    	,PASISendAffiliateDate  = Format(POM.EntryDate,'yyyy-MM-dd') " & vbCrLf &
                          "    	,PartNo=POD.PartNo    " & vbCrLf &
                          "    	,PartName=MP.PartName" & vbCrLf &
                          "    	,QtyPO = ISNULL(POD.TotalPOQty,0) " & vbCrLf &
                          "   	,QtyBox = ISNULL(POD.POQtyBox,MPM.QtyBox)   " & vbCrLf &
                          "   	,MPM.BoxPallet   " & vbCrLf &
                          "   	,VolumePallet = case When MPM.BoxPallet = '0' Then '0' Else (POD.TotalPOQty/ISNULL(POD.POQtyBox,MPM.QtyBox))/MPM.BoxPallet End" & vbCrLf &
                          "    	,ETDSupp = POM.ETDVendor1    " & vbCrLf &
                          "    	,ETAAff = POM.ETAFactory1 " & vbCrLf &
                          "    	,PASIDeliveryDate = DSM.DeliveryDate " & vbCrLf &
                          "    	,PASISuratJalanNo = COALESCE(DSM.SuratJalanNo,RFM.SuratJalanNo) "

        ls_SQL = ls_SQL + "  	,PASIInvoiceNo = SD.ShippingInstructionNo " & vbCrLf &
                          "  	,PASIDeliveryQty = SD.ShippingQty   " & vbCrLf &
                          "   	,Remaining = (ISNULL(POD.TotalPOQty,0) -   " & vbCrLf &
                          "       (  " & vbCrLf &
                          "           Select ISNULL(SUM(ShippingQty), 0)   " & vbCrLf &
                          "           FROM ShippingInstruction_Detail RRD   " & vbCrLf &
                          "           LEFT JOIN ShippingInstruction_Master RRM ON RRM.AffiliateID = RRD.AffiliateID   " & vbCrLf &
                          "                 AND RRM.ForwarderID = RRD.ForwarderID " & vbCrLf &
                          "                 AND RRM.ShippingInstructionNo = RRD.ShippingInstructionNo " & vbCrLf &
                          "           WHERE RRD.OrderNo = RFM.OrderNo " & vbCrLf &
                          "             AND RRD.SupplierID = RFM.SupplierID  " & vbCrLf &
                          "             AND RRD.AffiliateID = RFM.AffiliateID  " & vbCrLf &
                          "             AND RRD.PartNo = RFD.PartNo  " & vbCrLf &
                          "             --AND RRM.ShippingInstructionDate <= RFM.ReceiveDate  " & vbCrLf &
                          "       )) * (-1)  " & vbCrLf &
                          "   FROM PO_Master_Export POM  " & vbCrLf &
                          "   LEFT JOIN PO_Detail_Export POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf &
                          "    						AND POM.PoNo = POD.PONo    " & vbCrLf &
                          "    						AND POM.SupplierID = POD.SupplierID    " & vbCrLf &
                          "  						AND POM.OrderNo1 = POD.OrderNo1  " & vbCrLf &
                          "   LEFT JOIN DOSupplier_Detail_Export DSD ON POM.AffiliateID = DSD.AffiliateID    " & vbCrLf &
                          "                                          AND POM.SupplierID = DSD.SupplierID    " & vbCrLf &
                          "                                          AND POM.PONO = DSD.PONO    " & vbCrLf &
                          "                                          AND POD.PartNo = DSD.PartNo    " & vbCrLf &
                          "                                          AND POD.PONO = DSD.PONO  " & vbCrLf &
                          "   LEFT JOIN DOSupplier_Master_Export DSM ON DSD.suratJalanNo = DSM.SuratJalanNo   " & vbCrLf &
                          "                                          AND DSD.affiliateID = DSM.affiliateID  " & vbCrLf &
                          "                                          AND DSD.SupplierID = DSM.SupplierID   " & vbCrLf &
                          "  										AND DSD.PONo = DSM.PONo  " & vbCrLf &
                          "  										AND DSD.OrderNo = DSM.OrderNo  " & vbCrLf &
                          "   LEFT JOIN ReceiveForwarder_Master RFM ON DSM.SuratJalanNo = RFM.SuratJalanNo  " & vbCrLf &
                          "  										AND DSM.SupplierID = RFM.SupplierID  " & vbCrLf &
                          "  										AND DSM.AffiliateID = RFM.AffiliateID  " & vbCrLf &
                          "  										AND DSM.PONo = RFM.PONo  " & vbCrLf &
                          "  										AND DSM.OrderNo = RFM.OrderNo  " & vbCrLf &
                          "  										AND POM.ForwarderID = RFM.ForwarderID  " & vbCrLf &
                          "   LEFT JOIN ReceiveForwarder_Detail RFD ON RFM.SupplierID = RFD.SupplierID  " & vbCrLf &
                          "  										AND RFM.AffiliateID = RFD.AffiliateID  " & vbCrLf &
                          "  										AND RFM.PONo = RFD.PONo  " & vbCrLf &
                          "  										AND RFM.OrderNo = RFD.OrderNo  " & vbCrLf &
                          "  										AND DSD.PartNo = RFD.PartNo  " & vbCrLf &
                          "  										AND RFM.SuratJalanNo = RFD.SuratJalanNo  " & vbCrLf &
                          "   LEFT JOIN (Select ShippingInstructionNo,SuratJalanNo,AffiliateID,SupplierID,OrderNo,PartNo,ShippingQty=SUM(ShippingQty) " & vbCrLf &
                          " 			   From ShippingInstruction_Detail " & vbCrLf &
                          " 			 Group By ShippingInstructionNo,SuratJalanNo,AffiliateID,SupplierID,OrderNo,PartNo " & vbCrLf &
                          " 			 ) SD  " & vbCrLf &
                          " 			ON RFD.SuratJalanNo = SD.SuratJalanNo " & vbCrLf &
                          " 			AND RFD.AffiliateID = SD.AffiliateID " & vbCrLf &
                          " 			AND RFD.SupplierID = SD.SupplierID " & vbCrLf &
                          " 			AND RFD.OrderNo = SD.OrderNo " & vbCrLf &
                          " 			AND RFD.PartNo = SD.PartNo " & vbCrLf &
                          "   LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf &
                          "   LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo AND MPM.AffiliateID = POD.AffiliateID AND MPM.SupplierID = POD.SupplierID " & vbCrLf &
                          "   WHERE POD.TotalPOQty > 0 " & vbCrLf



        ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + ")A " & vbCrLf &
                          " Order By Period,PONo,AffiliateID,SupplierID,PASISuratJalanNo " & vbCrLf
        uf_SQL_Cek = ls_SQL
    End Function

    Private Function uf_SQL() As String
        Dim ls_filter As String = ""
        ls_SQL = ""

        Dim ls_End As String = ""
        ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

        'AFFILIATE CODE
        If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
            ls_filter = ls_filter +
                          "                      AND POM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
        End If

        'SUPPLIER CODE
        If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
            ls_filter = ls_filter +
                          "                      AND POM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
        End If

        'AFFILIATE PO PERIOD
        If chkPOPeriod.Checked = True Then
            ls_filter = ls_filter +
                          "                      AND CONVERT(date,POM.Period) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyy-MM-01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyy-MM-" & ls_End) & "' " & vbCrLf
        End If

        'SUPPLIER DELIVERY DATE
        If chkSupplierDelDate.Checked = True Then
            ls_filter = ls_filter +
                          "                      AND CONVERT(date,POM.ETDVendor1) BETWEEN '" & Format(dtSupplierDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtSupplierDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
        End If

        'PASI DELIVERY DATE
        If txtPONo.Text <> "" Then
            ls_filter = ls_filter +
                          "                      AND POD.PONo = '" & txtPONo.Text & "' " & vbCrLf
        End If



        ls_SQL = "   SELECT *   " & vbCrLf &
                  "   ,StatusPO = CASE WHEN Remaining > 0 THEN 'OVER'   " & vbCrLf &
                  "   WHEN Remaining < 0 THEN 'OPEN'   " & vbCrLf &
                  "   ELSE 'CLOSE' END   " & vbCrLf &
                  "   FROM (   " & vbCrLf &
                  "   SELECT  Distinct   " & vbCrLf &
                  "    	POM.Period    " & vbCrLf &
                  "    	,PONo=TRIM(POM.PONo)" & vbCrLf &
                  "    	,OrderNo=TRIM(POM.OrderNo1)" & vbCrLf &
                  "    	,AffiliateID=POM.AffiliateID" & vbCrLf &
                  "    	,SupplierID=POM.SupplierID" & vbCrLf &
                  "    	--,POKanban = '' " & vbCrLf

        ls_SQL = ls_SQL + "    	,PASISendAffiliateDate  = Format(POM.EntryDate,'yyyy-MM-dd') " & vbCrLf &
                          "    	,PartNo=POD.PartNo    " & vbCrLf &
                          "    	,PartName=MP.PartName" & vbCrLf &
                          "    	,QtyPO = ISNULL(POD.TotalPOQty,0) " & vbCrLf &
                          "   	,QtyBox = ISNULL(POD.POQtyBox,MPM.QtyBox)   " & vbCrLf &
                          "   	,MPM.BoxPallet   " & vbCrLf &
                          "   	,VolumePallet = case When MPM.BoxPallet = '0' Then '0' Else (POD.TotalPOQty/ISNULL(POD.POQtyBox,MPM.QtyBox))/MPM.BoxPallet End" & vbCrLf &
                          "    	,ETDSupp = POM.ETDVendor1    " & vbCrLf &
                          "    	,ETAAff = POM.ETAFactory1 " & vbCrLf &
                          "    	,PASIDeliveryDate = DSM.DeliveryDate " & vbCrLf &
                          "    	,PASISuratJalanNo = COALESCE(DSM.SuratJalanNo,RFM.SuratJalanNo) "

        ls_SQL = ls_SQL + "  	,PASIInvoiceNo = SD.ShippingInstructionNo " & vbCrLf &
                          "  	,PASIDeliveryQty = SD.ShippingQty   " & vbCrLf &
                          "   	,Remaining = (ISNULL(POD.TotalPOQty,0) -   " & vbCrLf &
                          "       (  " & vbCrLf &
                          "           Select ISNULL(SUM(ShippingQty), 0)   " & vbCrLf &
                          "           FROM ShippingInstruction_Detail RRD   " & vbCrLf &
                          "           LEFT JOIN ShippingInstruction_Master RRM ON RRM.AffiliateID = RRD.AffiliateID   " & vbCrLf &
                          "                 AND RRM.ForwarderID = RRD.ForwarderID " & vbCrLf &
                          "                 AND RRM.ShippingInstructionNo = RRD.ShippingInstructionNo " & vbCrLf &
                          "           WHERE RRD.OrderNo = POD.OrderNo1 " & vbCrLf &
                          "             AND RRD.SupplierID = POD.SupplierID  "

        ls_SQL = ls_SQL + "             AND RRD.AffiliateID = POD.AffiliateID  " & vbCrLf &
                          "             AND RRD.PartNo = POD.PartNo  " & vbCrLf &
                          "             --AND RRM.ShippingInstructionDate <= RFM.ReceiveDate  " & vbCrLf &
                          "       )) * (-1)  " & vbCrLf &
                          "   FROM PO_Master_Export POM  " & vbCrLf &
                          "   LEFT JOIN PO_Detail_Export POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf &
                          "    						AND POM.PoNo = POD.PONo    " & vbCrLf &
                          "    						AND POM.SupplierID = POD.SupplierID    " & vbCrLf &
                          "  						AND POM.OrderNo1 = POD.OrderNo1  " & vbCrLf &
                          "   LEFT JOIN ReceiveForwarder_Detail RFD ON POD.SupplierID = RFD.SupplierID  " & vbCrLf &
                          "  										AND POD.AffiliateID = RFD.AffiliateID  "

        ls_SQL = ls_SQL + "  										AND POD.PONo = RFD.PONo  " & vbCrLf &
                          "  										AND POD.OrderNo1 = RFD.OrderNo  " & vbCrLf &
                          "  										AND POD.PartNo = RFD.PartNo  " & vbCrLf &
                          "   LEFT JOIN ReceiveForwarder_Master RFM ON RFD.SuratJalanNo = RFM.SuratJalanNo  " & vbCrLf &
                          "  										AND RFD.SupplierID = RFM.SupplierID  " & vbCrLf &
                          "  										AND RFD.AffiliateID = RFM.AffiliateID  " & vbCrLf &
                          "  										AND RFD.PONo = RFM.PONo  " & vbCrLf &
                          "  										AND RFD.OrderNo = RFM.OrderNo  " & vbCrLf &
                          "   LEFT JOIN DOSupplier_Master_Export DSM ON RFM.suratJalanNo = DSM.SuratJalanNo   " & vbCrLf &
                          "                                          AND RFM.affiliateID = DSM.affiliateID  " & vbCrLf &
                          "                                          AND RFM.SupplierID = DSM.SupplierID   "

        ls_SQL = ls_SQL + "  										AND RFM.PONo = DSM.PONo  " & vbCrLf &
                          "  										AND RFM.OrderNo = DSM.OrderNo  " & vbCrLf &
                          "   LEFT JOIN DOSupplier_Detail_Export DSD ON DSM.SuratJalanNo = DSD.SuratjalanNo    " & vbCrLf &
                          "                                          AND DSM.AffiliateID = DSD.AffiliateID    " & vbCrLf &
                          "                                          AND DSM.SupplierID = DSD.SupplierID    " & vbCrLf &
                          "                                          AND DSM.PONO = DSD.PONO    " & vbCrLf &
                          "                                          AND RFD.PartNo = DSD.PartNo    " & vbCrLf &
                          "                                          AND RFD.PONO = DSD.PONO  " & vbCrLf &
                          "   LEFT JOIN (Select ShippingInstructionNo,SuratJalanNo,AffiliateID,SupplierID,OrderNo,PartNo,ShippingQty=SUM(ShippingQty) " & vbCrLf &
                          " 			   From ShippingInstruction_Detail " & vbCrLf &
                          " 			 Group By ShippingInstructionNo,SuratJalanNo,AffiliateID,SupplierID,OrderNo,PartNo "

        ls_SQL = ls_SQL + " 			 ) SD  " & vbCrLf &
                          " 			ON RFD.SuratJalanNo = SD.SuratJalanNo " & vbCrLf &
                          " 			AND RFD.AffiliateID = SD.AffiliateID " & vbCrLf &
                          " 			AND RFD.SupplierID = SD.SupplierID " & vbCrLf &
                          " 			AND RFD.OrderNo = SD.OrderNo " & vbCrLf &
                          " 			AND RFD.PartNo = SD.PartNo " & vbCrLf &
                          "   LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf &
                          "   LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo AND MPM.AffiliateID = POD.AffiliateID AND MPM.SupplierID = POD.SupplierID " & vbCrLf &
                          "   WHERE POD.TotalPOQty > 0 " & vbCrLf



        ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + ")A " & vbCrLf &
                          " Order By Period,PONo,AffiliateID,SupplierID,PASISuratJalanNo " & vbCrLf
        uf_SQL = ls_SQL
    End Function

    Private Function GetSummaryOutStanding() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""

                'Dim ls_End As String = ""
                'ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

                ''AFFILIATE CODE
                'If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                '    ls_filter = ls_filter + _
                '                  "                      AND POM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
                'End If

                ''SUPPLIER CODE
                'If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
                '    ls_filter = ls_filter + _
                '                  "                      AND POM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
                'End If

                ''AFFILIATE PO PERIOD
                'If chkPOPeriod.Checked = True Then
                '    ls_filter = ls_filter + _
                '                  "                      AND CONVERT(date,POM.Period) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyy-MM-01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyy-MM-" & ls_End) & "' " & vbCrLf
                'End If

                ''SUPPLIER DELIVERY DATE
                'If chkSupplierDelDate.Checked = True Then
                '    ls_filter = ls_filter + _
                '                  "                      AND CONVERT(date,POM.ETDVendor1) BETWEEN '" & Format(dtSupplierDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtSupplierDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
                'End If

                ''PASI DELIVERY DATE
                'If txtPONo.Text <> "" Then
                '    ls_filter = ls_filter + _
                '                  "                      AND POD.PONo = '" & txtPONo.Text & "' " & vbCrLf
                'End If



                'ls_sql = "   SELECT *   " & vbCrLf & _
                '      "   ,StatusPO = CASE WHEN Remaining > 0 THEN 'OVER'   " & vbCrLf & _
                '      "   WHEN Remaining < 0 THEN 'OPEN'   " & vbCrLf & _
                '      "   ELSE 'CLOSE' END   " & vbCrLf & _
                '      "   FROM (   " & vbCrLf & _
                '      "   SELECT     " & vbCrLf & _
                '      "    	POM.Period    " & vbCrLf & _
                '      "    	,POM.PONo    " & vbCrLf & _
                '      "    	,POM.AffiliateID    " & vbCrLf & _
                '      "    	,POM.SupplierID    " & vbCrLf & _
                '      "    	,POKanban = '' "

                'ls_sql = ls_sql + "    	,PASISendAffiliateDate  = Format(POM.EntryDate,'yyyy-MM-dd') " & vbCrLf & _
                '                  "    	,POD.PartNo    " & vbCrLf & _
                '                  "    	,MP.PartName    " & vbCrLf & _
                '                  "    	,QtyPO = ISNULL(POD.TotalPOQty,0) " & vbCrLf & _
                '                  "   	,MPM.QtyBox   " & vbCrLf & _
                '                  "   	,MPM.BoxPallet   " & vbCrLf & _
                '                  "   	,VolumePallet = (POD.TotalPOQty/MPM.QtyBox)/MPM.BoxPallet   " & vbCrLf & _
                '                  "    	,ETDSupp = POM.ETDVendor1    " & vbCrLf & _
                '                  "    	,ETAAff = POM.ETAFactory1 " & vbCrLf & _
                '                  "    	,PASIDeliveryDate = DSM.DeliveryDate " & vbCrLf & _
                '                  "    	,PASISuratJalanNo = COALESCE(DSM.SuratJalanNo,RFM.SuratJalanNo) "

                'ls_sql = ls_sql + "  	,PASIInvoiceNo = SD.ShippingInstructionNo " & vbCrLf & _
                '                  "  	,PASIDeliveryQty = SD.ShippingQty   " & vbCrLf & _
                '                  "   	,Remaining = (ISNULL(POD.TotalPOQty,0) -   " & vbCrLf & _
                '                  "       (  " & vbCrLf & _
                '                  "           Select ISNULL(SUM(ShippingQty), 0)   " & vbCrLf & _
                '                  "           FROM ShippingInstruction_Detail RRD   " & vbCrLf & _
                '                  "           LEFT JOIN ShippingInstruction_Master RRM ON RRM.AffiliateID = RRD.AffiliateID   " & vbCrLf & _
                '                  "                 AND RRM.ForwarderID = RRD.ForwarderID " & vbCrLf & _
                '                  "                 AND RRM.ShippingInstructionNo = RRD.ShippingInstructionNo " & vbCrLf & _
                '                  "           WHERE RRD.OrderNo = POD.OrderNo1 " & vbCrLf & _
                '                  "             AND RRD.SupplierID = POD.SupplierID  "

                'ls_sql = ls_sql + "             AND RRD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                '                  "             AND RRD.PartNo = POD.PartNo  " & vbCrLf & _
                '                  "             --AND RRM.ShippingInstructionDate <= RFM.ReceiveDate  " & vbCrLf & _
                '                  "       )) * (-1)  " & vbCrLf & _
                '                  "   FROM PO_Master_Export POM  " & vbCrLf & _
                '                  "   LEFT JOIN PO_Detail_Export POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                '                  "    						AND POM.PoNo = POD.PONo    " & vbCrLf & _
                '                  "    						AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
                '                  "  						AND POM.OrderNo1 = POD.OrderNo1  " & vbCrLf & _
                '                  "   INNER JOIN ReceiveForwarder_Detail RFD ON POD.SupplierID = RFD.SupplierID  " & vbCrLf & _
                '                  "  										AND POD.AffiliateID = RFD.AffiliateID  "

                'ls_sql = ls_sql + "  										AND POD.PONo = RFD.PONo  " & vbCrLf & _
                '                  "  										AND POD.OrderNo1 = RFD.OrderNo  " & vbCrLf & _
                '                  "  										AND POD.PartNo = RFD.PartNo  " & vbCrLf & _
                '                  "   INNER JOIN ReceiveForwarder_Master RFM ON RFD.SuratJalanNo = RFM.SuratJalanNo  " & vbCrLf & _
                '                  "  										AND RFD.SupplierID = RFM.SupplierID  " & vbCrLf & _
                '                  "  										AND RFD.AffiliateID = RFM.AffiliateID  " & vbCrLf & _
                '                  "  										AND RFD.PONo = RFM.PONo  " & vbCrLf & _
                '                  "  										AND RFD.OrderNo = RFM.OrderNo  " & vbCrLf & _
                '                  "   LEFT JOIN DOSupplier_Master_Export DSM ON RFM.suratJalanNo = DSM.SuratJalanNo   " & vbCrLf & _
                '                  "                                          AND RFM.affiliateID = DSM.affiliateID  " & vbCrLf & _
                '                  "                                          AND RFM.SupplierID = DSM.SupplierID   "

                'ls_sql = ls_sql + "  										AND RFM.PONo = DSM.PONo  " & vbCrLf & _
                '                  "  										AND RFM.OrderNo = DSM.OrderNo  " & vbCrLf & _
                '                  "   LEFT JOIN DOSupplier_Detail_Export DSD ON DSM.SuratJalanNo = DSD.SuratjalanNo    " & vbCrLf & _
                '                  "                                          AND DSM.AffiliateID = DSD.AffiliateID    " & vbCrLf & _
                '                  "                                          AND DSM.SupplierID = DSD.SupplierID    " & vbCrLf & _
                '                  "                                          AND DSM.PONO = DSD.PONO    " & vbCrLf & _
                '                  "                                          AND RFD.PartNo = DSD.PartNo    " & vbCrLf & _
                '                  "                                          AND RFD.PONO = DSD.PONO  " & vbCrLf & _
                '                  "   LEFT JOIN (Select ShippingInstructionNo,SuratJalanNo,AffiliateID,SupplierID,OrderNo,PartNo,ShippingQty=SUM(ShippingQty) " & vbCrLf & _
                '                  " 			   From ShippingInstruction_Detail " & vbCrLf & _
                '                  " 			 Group By ShippingInstructionNo,SuratJalanNo,AffiliateID,SupplierID,OrderNo,PartNo "

                'ls_sql = ls_sql + " 			 ) SD  " & vbCrLf & _
                '                  " 			ON RFD.SuratJalanNo = SD.SuratJalanNo " & vbCrLf & _
                '                  " 			AND RFD.AffiliateID = SD.AffiliateID " & vbCrLf & _
                '                  " 			AND RFD.SupplierID = SD.SupplierID " & vbCrLf & _
                '                  " 			AND RFD.OrderNo = SD.OrderNo " & vbCrLf & _
                '                  " 			AND RFD.PartNo = SD.PartNo " & vbCrLf & _
                '                  "   LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                '                  "   LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo AND MPM.AffiliateID = POD.AffiliateID AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
                '                  "   WHERE POD.TotalPOQty > 0 " & vbCrLf



                'ls_sql = ls_sql + ls_filter & vbCrLf

                'ls_sql = ls_sql + ")A " & vbCrLf & _
                '                  " Order By Period,PONo,AffiliateID,SupplierID,PASISuratJalanNo " & vbCrLf

                ls_sql = uf_SQL()

                Dim Cmd As New SqlCommand(ls_sql, cn)
                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 300
                da.Fill(dt)
                dt = Trimdata(dt)
                Return dt
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function Trimdata(dt As DataTable) As DataTable
        For Each c As DataColumn In dt.Columns
            If c.ColumnName = "PONo" Or c.ColumnName = "AffiliateID" Or c.ColumnName = "SupplierID" Or c.ColumnName = "PartNo" Or c.ColumnName = "PartName" Or c.ColumnName = "PASISuratJalanNo" Or c.ColumnName = "PASIInvoiceNo" Then
                For Each r As DataRow In dt.Rows
                    Try
                        r(c.ColumnName) = r(c.ColumnName).ToString().Trim()
                    Catch
                    End Try
                Next
            End If
        Next
        Return dt
    End Function

    Private Sub up_GridLoadWhenEventChange()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT  Top 0 " & vbCrLf &
                  "  	 Period = '' " & vbCrLf &
                  "  	,PONo = '' " & vbCrLf &
                  "  	,AffiliateID = '' " & vbCrLf &
                  "  	,SupplierID = '' " & vbCrLf &
                  "  	--,POKanban = '' " & vbCrLf &
                  "  	,PASISendAffiliateDate = '' " & vbCrLf &
                  "  	,PartNo = '' " & vbCrLf &
                  "  	,PartName = '' " & vbCrLf &
                  "  	,QtyPO = '' " & vbCrLf &
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
                     "SELECT AffiliateID = RTRIM(AffiliateID), AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate Where isnull(overseascls, '0') = '1'"
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
                     "SELECT SupplierID = RTRIM(SupplierID), SupplierName = RTRIM(SupplierName) FROM dbo.MS_supplier Where isnull(overseas, '0') = '0'"
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
            Dim tempFile As String = "TemplateSummaryOutstandingDeliveryPasiToAffiliateExp " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                .Cells(8, 6, pData.Rows.Count + 7, 6).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 13, pData.Rows.Count + 7, 13).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 14, pData.Rows.Count + 7, 14).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 15, pData.Rows.Count + 7, 15).Style.Numberformat.Format = "dd-mmm-yy"
                '.Cells(8, 19, pData.Rows.Count + 7, 19).Style.Numberformat.Format = "dd-mmm-yy"
                '.Cells(8, 21, pData.Rows.Count + 7, 21).Style.Numberformat.Format = "dd-mmm-yy"

                .Cells(8, 9, pData.Rows.Count + 7, 9).Style.Numberformat.Format = "#,##0"
                .Cells(8, 10, pData.Rows.Count + 7, 10).Style.Numberformat.Format = "#,##0"
                .Cells(8, 11, pData.Rows.Count + 7, 11).Style.Numberformat.Format = "#,##0"
                .Cells(8, 12, pData.Rows.Count + 7, 12).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 17, pData.Rows.Count + 7, 17).Style.Numberformat.Format = "#,##0"
                .Cells(8, 18, pData.Rows.Count + 7, 18).Style.Numberformat.Format = "#,##0"
                .Cells(8, 19, pData.Rows.Count + 7, 19).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 20, pData.Rows.Count + 7, 20).Style.Numberformat.Format = "#,##0"
                .Cells(8, 21, pData.Rows.Count + 7, 21).Style.Numberformat.Format = "#,##0"

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

                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 20)
                EpPlusDrawAllBorders(rgAll)
            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    'Private Sub epplusExportExcelOLD(ByVal pFilename As String, ByVal pSheetName As String,
    '                          ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

    '    Try
    '        Dim tempFile As String = "Summary Outstanding " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
    '        Dim NewFileName As String = Server.MapPath("~\ProgressReport\Import\" & tempFile & "")
    '        If (System.IO.File.Exists(pFilename)) Then
    '            System.IO.File.Copy(pFilename, NewFileName, True)
    '        End If

    '        Dim rowstart As String = Split(pCellStart, ":")(1)
    '        Dim Coltart As String = Split(pCellStart, ":")(0)
    '        Dim fi As New FileInfo(NewFileName)

    '        Dim exl As New ExcelPackage(fi)
    '        Dim ws As ExcelWorksheet

    '        ws = exl.Workbook.Worksheets(pSheetName)
    '        Dim irow As Integer = 0
    '        Dim icol As Integer = 0

    '        With ws
    '            .Cells(3, 4).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
    '            .Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

    '            For irow = 0 To pData.Rows.Count - 1
    '                For icol = 1 To pData.Columns.Count
    '                    .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
    '                    If icol = 7 Or icol = 8 Or icol = 14 Or icol = 15 Or icol = 16 Or icol = 20 Or icol = 23 Or icol = 26 Or icol = 29 Then
    '                        .Cells(irow + rowstart, icol).Style.Numberformat.Format = "dd-mmm-yy"
    '                    End If
    '                    If icol = 11 Or icol = 13 Or icol = 18 Or icol = 19 Or icol = 21 Or icol = 28 Or icol = 30 Or icol = 25 Or icol = 34 Then
    '                        .Cells(irow + rowstart, icol).Style.Numberformat.Format = "#,##0"
    '                    End If
    '                Next
    '            Next

    '            Dim rgAll As ExcelRange = .Cells(8, 1, irow + 8, 34)
    '            EpPlusDrawAllBorders(rgAll)

    '        End With

    '        exl.Save()

    '        DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

    '        exl = Nothing
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
                    Dim dtProd As DataTable = GetSummaryOutStanding()
                    FileName = "TemplateSummaryOutstandingDeliveryPasiToAffiliateExp.xlsx"
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