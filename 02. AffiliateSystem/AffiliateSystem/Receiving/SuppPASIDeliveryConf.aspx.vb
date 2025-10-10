Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing


Public Class SuppPASIDeliveryConf
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = ""

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "E01"

    Dim dt_SupplierPeriod As String = "", _
        dt_RecDateFrom As String = "", _
        dt_RecDateTo As String = ""

#End Region

#Region "Procedures"
    Private Sub up_Initialize()
        Dim script As String = _
            "var a = new Date(); " & vbCrLf & _
            "dtSupplierPeriod.SetDate(a); " & vbCrLf & _
            "rdrSADAll.SetValue(true); " & vbCrLf & _
            "rdrPADAll.SetValue(true); " & vbCrLf & _
            "rdrRRQAll.SetValue(true); " & vbCrLf & _
            "rdrSDAll.SetValue(true); " & vbCrLf & _
            "rdrPOKAll.SetValue(true); " & vbCrLf & _
            "rdrMCPAll.SetValue(true); " & vbCrLf & _
            "rdrGRSAll.SetValue(true); " & vbCrLf & _
            "chkSupplierPeriod.SetValue(false); " & vbCrLf & _
            "txtSupplierSJNo.SetText(''); " & vbCrLf & _
            "txtPONo.SetText(''); " & vbCrLf & _
            "if (cboSupplier.GetItemCount() > 1) { " & vbCrLf & _
            "   txtSupplierName.SetText('==ALL=='); " & vbCrLf & _
            "   cboSupplier.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            "if (cboPart.GetItemCount() > 1) { " & vbCrLf & _
            "   txtPartName.SetText('==ALL=='); " & vbCrLf & _
            "   cboPart.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            "cboSupplier.SetValue('==ALL=='); " & vbCrLf & _
            "var date = new Date(a.getFullYear(),a.getMonth(),1); " & vbCrLf & _
            "dtRecDateFrom.SetDate(date); " & vbCrLf & _
            "dtRecDateTo.SetDate(a); " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(chkSupplierPeriod, chkSupplierPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""

            ls_SQL = " SELECT *   " & vbCrLf & _
                  "  	FROM (  " & vbCrLf & _
                  "  		  SELECT ColNo = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PONo, SortKanbanNo,SupplierSJNo, PASISJNo)),  " & vbCrLf & _
                  "  			     Period, PONo, DeliveryLocationCode, DeliveryLocationName, SupplierCode, SupplierName, POKanban = CASE WHEN ISNULL(KanbanCls,'0') = '1' THEN 'YES' ELSE 'NO' END,   " & vbCrLf & _
                  "  				 KanbanNo, SupplierPlanDeliveryDate, SupplierDeliveryDate, SupplierSJNo, PASIDeliveryDate, PASISJNo, PartNo, PartName, UOM,   " & vbCrLf & _
                  "  				 SupplierDeliveryQty, PASIGoodReceivingQty, PASIDefectQty, PASIDeliveryQty, GoodReceivingQty, DefectReceivingQty, RemainingReceivingQty,  " & vbCrLf & _
                  "  				 ReceivedDate, ReceivedBy, GoodReceivingSent, detail, url, DeliveryByPASICls,  " & vbCrLf & _
                  "  				 SortPONo = PONo, SortKanbanNo = SortKanbanNo, SortHeader = 0,  " & vbCrLf & _
                  "                   SortDeliveryByPASICls = DeliveryByPASICls, /*SortPASISJNo = SupplierSJNO, SortSupplierSJNo = supSJ */ " & vbCrLf & _
                  "                   SortPASISJNo = PASISJNo, SortSupplierSJNo = SupplierSJNo " & vbCrLf & _
                  "  			FROM (  "

            ls_SQL = ls_SQL + "  				  SELECT DISTINCT   " & vbCrLf & _
                              "  						 Period = SUBSTRING(CONVERT(CHAR,POM.Period,106),4,9),  " & vbCrLf & _
                              "  						 PONo = POD.PONo,   " & vbCrLf & _
                              "  						 DeliveryLocationCode = ISNULL(KM.DeliveryLocationCode,''),  " & vbCrLf & _
                              "                          DeliveryLocationName = ISNULL(MD.DeliveryLocationName,''),  " & vbCrLf & _
                              "  						 SupplierCode = POM.SupplierID,  " & vbCrLf & _
                              "                          SupplierName = MS.SupplierName,  " & vbCrLf & _
                              "  						 KanbanCls = ISNULL(POD.KanbanCls,'0'),  " & vbCrLf & _
                              "  						 KanbanNo = ISNULL(KD.KanbanNo,''),  " & vbCrLf & _
                              "  						 SupplierPlanDeliveryDate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),  " & vbCrLf & _
                              "  						 SupplierDeliveryDate = ISNULL(CONVERT(CHAR,DSM.DeliveryDate,106),''),  "

            ls_SQL = ls_SQL + "  						 SupplierSJNO = ISNULL(DSD.SuratJalanNo,''),  " & vbCrLf & _
                              "  						 PASIDeliveryDate = ISNULL(CONVERT(CHAR,DPM.DeliveryDate,106),''),  " & vbCrLf & _
                              "  						 PASISJNo = ISNULL(DPD.SuratJalanNo,''),  " & vbCrLf & _
                              "  						 PartNo = '',  " & vbCrLf & _
                              "  						 PartName = '',  " & vbCrLf & _
                              "  						 UOM = '',  " & vbCrLf & _
                              "  						 SupplierDeliveryQty = '',  " & vbCrLf & _
                              "  						 PASIGoodReceivingQty = '',  " & vbCrLf & _
                              "  						 PASIDefectQty = '',  " & vbCrLf & _
                              "  						 PASIDeliveryQty = '',  " & vbCrLf & _
                              "  						 GoodReceivingQty = '',  "

            ls_SQL = ls_SQL + "  						 DefectReceivingQty = '',  " & vbCrLf & _
                              "  						 RemainingReceivingQty = '',  " & vbCrLf & _
                              "  				         ReceivedDate = ISNULL(CONVERT(CHAR(20),RAM.ReceiveDate,113),''),  " & vbCrLf & _
                              "  				         ReceivedBy = ISNULL(RAM.ReceiveBy,''),  " & vbCrLf & _
                              "  						 GoodReceivingSent = CASE WHEN ISNULL(RAM.ExcelCls,0) = 0 then 'NO' else 'YES' END,  " & vbCrLf & _
                              "  						 url = 'ReceivingEntry.aspx?prm='+Rtrim(ISNULL(CONVERT(CHAR(11),RAM.ReceiveDate,106)+' '+CONVERT(CHAR(8),RAM.ReceiveDate,108), CONVERT(CHAR(11),GETDATE(),106)+' '+CONVERT(CHAR(8),GETDATE(),108)))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(POD.SupplierID)  " & vbCrLf & _
                              "  									+ '|' +Rtrim(MS.SupplierName)  " & vbCrLf & _
                              "  									+ '|' +Rtrim(ISNULL(DSD.SuratJalanNo,''))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''))   " & vbCrLf & _
                              "  									+ '|' +Rtrim(ISNULL(CONVERT(CHAR,DSM.DeliveryDate,106),''))  "

            ls_SQL = ls_SQL + "  									+ '|' +Rtrim(ISNULL(DPD.SuratJalanNo,''))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(ISNULL(CONVERT(CHAR,DPM.DeliveryDate,106),''))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(ISNULL(KM.DeliveryLocationCode,''))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(ISNULL(MD.DeliveryLocationName,''))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(COALESCE(RAM.DriverName,DPM.DriverName,DSM.DriverName,''))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(COALESCE(RAM.DriverContact,DPM.DriverContact,DSM.DriverContact,''))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(COALESCE(RAM.NoPol,DPM.NoPol,DSM.NoPol,''))  " & vbCrLf & _
                              "  									+ '|' +Rtrim(COALESCE(RAM.JenisArmada,DPM.JenisArmada,DSM.JenisArmada,''))  " & vbCrLf & _
                              "  									+ '|' +'0' " & vbCrLf & _
                              "  									+ '|' +Rtrim(POD.PONo)   " & vbCrLf & _
                              "  									+ '|' +Rtrim(ISNULL(KD.KanbanNo,'')) + '|' + DeliveryByPASICls,  " & vbCrLf

            ls_SQL = ls_SQL + "  						 detail = (CASE WHEN isnull(RAM.SuratJalanNo,'') = ''THEN 'RECEIVE' ELSE 'DETAIL' END),  " & vbCrLf & _
                              "                          --detail = CASE WHEN ISNULL(DeliveryByPasiCls,0) = 1 THEN " & vbCrLf & _
                              "			                               --(CASE WHEN isnull(RAM.SuratJalanNo,'') = ''THEN 'RECEIVE' ELSE 'DETAIL' END)  " & vbCrLf & _
                              "	                                  --ELSE  " & vbCrLf & _
                              "			                               --(CASE WHEN isnull(DSM.SuratJalanNo,'') = ''THEN 'RECEIVE' ELSE 'DETAIL' END)  " & vbCrLf & _
                              "	                                  --END, " & vbCrLf & _
                              "                           DeliveryByPASICls, supSJ = DPM.SuratJalanNo, SortKanbanNo = ISNULL(KD.KanbanNo,'')  " & vbCrLf & _
                              "  					FROM PO_DETAIL POD   " & vbCrLf & _
                              "  						 LEFT JOIN PO_Master POM ON POM.AffiliateID =POD.AffiliateID  " & vbCrLf & _
                              "  							AND POM.SupplierID =POD.SupplierID  " & vbCrLf & _
                              "  							AND POM.PONO =POD.PONO  " & vbCrLf & _
                              "  						 LEFT JOIN Kanban_Detail KD ON KD.AffiliateID =POD.AffiliateID  " & vbCrLf & _
                              "  							AND KD.SupplierID =POD.SupplierID  "

            ls_SQL = ls_SQL + "  							AND KD.PONO =POD.PONO  " & vbCrLf & _
                              "  							AND KD.PartNo =POD.PartNo  " & vbCrLf & _
                              "  						 LEFT JOIN Kanban_Master KM ON KD.AffiliateID =KM.AffiliateID  " & vbCrLf & _
                              "  							AND KD.SupplierID =KM.SupplierID  " & vbCrLf & _
                              "  							AND KD.KanbanNo =KM.KanbanNo  " & vbCrLf & _
                              "                              AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                              "  						 /*LEFT */INNER JOIN DOSupplier_Detail DSD ON KD.AffiliateID =DSD.AffiliateID  " & vbCrLf & _
                              "  							AND KD.SupplierID =DSD.SupplierID  " & vbCrLf & _
                              "  							AND KD.PONO =DSD.PONO  " & vbCrLf & _
                              "  							AND KD.PartNo =DSD.PartNo  " & vbCrLf & _
                              "  							AND KD.KanbanNo =DSD.KanbanNo  "

            ls_SQL = ls_SQL + "  						 LEFT JOIN DOSupplier_Master DSM ON DSM.AffiliateID =DSD.AffiliateID  " & vbCrLf & _
                              "  							AND DSM.SupplierID =DSD.SupplierID  " & vbCrLf & _
                              "  							AND DSM.SuratJalanNo =DSD.SuratJalanNo  " & vbCrLf & _
                              "  						 LEFT JOIN DOPASI_Detail DPD ON KD.AffiliateID =DPD.AffiliateID  " & vbCrLf & _
                              "  							AND KD.SupplierID =DPD.SupplierID  " & vbCrLf & _
                              "  							AND KD.PONO =DPD.PONO  " & vbCrLf & _
                              "  							AND KD.PartNo =DPD.PartNo  " & vbCrLf & _
                              "  							AND KD.KanbanNo =DPD.KanbanNo  " & vbCrLf & _
                              "  							AND DPD.SuratJalanNoSupplier = DSM.SuratJalanNo " & vbCrLf & _
                              "  						 LEFT JOIN DOPASI_Master DPM ON DPM.AffiliateID =DPD.AffiliateID  " & vbCrLf & _
                              "  							--AND DPM.SupplierID =DPD.SupplierID  " & vbCrLf

            ls_SQL = ls_SQL + "  							AND DPM.SuratJalanNo =DPD.SuratJalanNo  " & vbCrLf & _
                              "  						 LEFT JOIN ReceivePASI_Detail RPD ON RPD.AffiliateID = DPM.AffiliateID  " & vbCrLf & _
                              "  							AND RPD.SupplierID = DPM.SupplierID  " & vbCrLf & _
                              "  							AND RPD.PONo = POD.PONo  " & vbCrLf & _
                              "  							AND RPD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  							AND RPD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "  							AND RPD.SuratjalanNo = DSM.SuratJalanNo " & vbCrLf & _
                              "  				         LEFT JOIN ReceiveAffiliate_Detail RAD ON RAD.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                              "  					        AND RAD.SupplierID = KD.SupplierID  " & vbCrLf & _
                              "  					        AND RAD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "  					        AND RAD.PONo = KD.PONo  " & vbCrLf & _
                              "                             AND RAD.SuratJalanNo = DPM.SuratJalanNo " & vbCrLf

            ls_SQL = ls_SQL + "  					        AND RAD.PartNo = KD.PartNo  " & vbCrLf & _
                              "  				         LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
                              "  					        --AND RAM.SupplierID = RAD.SupplierID  " & vbCrLf & _
                              "  					        AND RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
                              "  						 LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                              "  						 LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls  " & vbCrLf & _
                              "                           LEFT JOIN MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                              "                           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "                           LEFT JOIN   " & vbCrLf & _
                              "   				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(RecQty),0) + ISNULL(SUM(DefectQty),0)  " & vbCrLf & _
                              "   				            FROM ReceiveAffiliate_Detail  "

            ls_SQL = ls_SQL + "   				           GROUP BY PONo, KanbanNo  " & vbCrLf & _
                              "   				         ) SumKanban ON SumKanban.PONo = KD.PONo AND SumKanban.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "   				         LEFT JOIN   " & vbCrLf & _
                              "   				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)  " & vbCrLf & _
                              "   				            FROM DOSupplier_Detail   " & vbCrLf & _
                              "   				           GROUP BY PONo, KanbanNo  " & vbCrLf & _
                              "   				         ) SumDSD ON SumDSD.PONo = KD.PONo AND SumDSD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "   				         LEFT JOIN   " & vbCrLf & _
                              "   				         (SELECT SuratJalanNoSupplier, PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)  " & vbCrLf & _
                              "   				            FROM DOPASI_Detail  " & vbCrLf & _
                              "   				           GROUP BY PONo, KanbanNo, SuratJalanNoSupplier  "

            ls_SQL = ls_SQL + "   				         ) SumDPD ON SumDPD.PONo = KD.PONo AND SumDPD.KanbanNo = KD.KanbanNo AND sumDPD.SuratJalanNoSupplier = DSM.SuratJalanNo " & vbCrLf & _
                              " 		           WHERE POD.AffiliateID = '" & Session("AffiliateID") & "' AND KD.KanbanQty <> 0 AND DPM.DeliveryDate is not null" & vbCrLf & _
                              "                    --AND POD.pono='PC1507-IKS' " & vbCrLf
            'SUPPLIER PLAN DELIVERY DATE (UNTIL)
            If chkSupplierPeriod.Checked = True Then
                If IsNothing(Session("fromE02")) Then
                    ls_SQL = ls_SQL + _
                              "                      AND CONVERT(DATETIME,KM.KanbanDate) <= CONVERT(DATETIME,'" & Format(dtSupplierPeriod.Value, "yyyy-MM-dd") & "') " & vbCrLf
                Else
                    ls_SQL = ls_SQL + _
                              "                      AND CONVERT(DATETIME,KM.KanbanDate) <= CONVERT(DATETIME,'" & Format(CDate(dt_SupplierPeriod), "yyyy-MM-dd") & "') " & vbCrLf
                End If
                'ls_SQL = ls_SQL + "                      AND ISNULL(DPM.DeliveryDate,DSM.DeliveryDate) <= '" & Format(dtSupplierPeriod.Value, "yyyy-MM-dd") & "'"
            End If
            'SUPPLIER ALREADY DELIVER
            If rdrSADYes.Value = True Then
                ls_SQL = ls_SQL + _
                              "                      AND ISNULL(CONVERT(CHAR,DSM.DeliveryDate),'') <> '' " & vbCrLf
            ElseIf rdrSADNo.Value = True Then
                ls_SQL = ls_SQL + _
                              "                      AND ISNULL(CONVERT(CHAR,DSM.DeliveryDate),'') = '' " & vbCrLf
            End If
            'PASI ALREADY DELIVER
            If rdrPADYes.Value = True Then
                ls_SQL = ls_SQL + _
                              "                      AND ISNULL(CONVERT(CHAR,DPM.DeliveryDate),'') <> '' " & vbCrLf
            ElseIf rdrPADNo.Value = True Then
                ls_SQL = ls_SQL + _
                              "                      AND ISNULL(CONVERT(CHAR,DPM.DeliveryDate),'') = '' " & vbCrLf
            End If
            'REMAINING RECEVING QTY
            If rdrRRQYes.Value = True Then
                ls_SQL = ls_SQL + _
                              "                      /*AND ISNULL(SumKanban.TQty,0) > (CASE WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(SumDPD.TQty,0) ELSE ISNULL(SumDSD.TQty,0) END)*/ " & vbCrLf & _
                              "                      AND CONVERT(NUMERIC(18,0), " & vbCrLf & _
                              " 				                 CASE WHEN ISNULL(POM.DeliveryByPASICls,'0') = '1'  " & vbCrLf & _
                              " 					                  THEN (ISNULL(DPD.DOQty,0) - (ISNULL(RAD.RecQty,0) + ISNULL(RAD.DefectQty,0)))  " & vbCrLf & _
                              " 					                  WHEN ISNULL(POM.DeliveryByPASICls,'0') = '0'  " & vbCrLf & _
                              " 					                  THEN (ISNULL(DSD.DOQty,0) - (ISNULL(RAD.RecQty,0) + ISNULL(RAD.DefectQty,0))) END) > 0 " & vbCrLf
            ElseIf rdrRRQNo.Value = True Then
                ls_SQL = ls_SQL + _
                              "                      /*AND ISNULL(SumKanban.TQty,0) <= (CASE WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(SumDPD.TQty,0) ELSE ISNULL(SumDSD.TQty,0) END)*/ " & vbCrLf & _
                              "                      AND CONVERT(NUMERIC(18,0), " & vbCrLf & _
                              " 				                 CASE WHEN ISNULL(POM.DeliveryByPASICls,'0') = '1'  " & vbCrLf & _
                              " 					                  THEN (ISNULL(DPD.DOQty,0) - (ISNULL(RAD.RecQty,0) + ISNULL(RAD.DefectQty,0)))  " & vbCrLf & _
                              " 					                  WHEN ISNULL(POM.DeliveryByPASICls,'0') = '0'  " & vbCrLf & _
                              " 					                  THEN (ISNULL(DSD.DOQty,0) - (ISNULL(RAD.RecQty,0) + ISNULL(RAD.DefectQty,0))) END) <= 0 " & vbCrLf
            End If

            'SUPPLIER SURAT JALAN NO.
            If Trim(txtSupplierSJNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                              "                      AND ISNULL(DSD.SuratJalanNo,'') LIKE '%" & Trim(txtSupplierSJNo.Text) & "%' OR ISNULL(DPD.SuratJalanNo,'') LIKE '%" & Trim(txtSupplierSJNo.Text) & "%' " & vbCrLf
            End If
            'RECEIVED DATE
            If chkRecDate.Checked = True Then
                'If IsNothing(Session("fromE02")) Then
                '    ls_SQL = ls_SQL + _
                '              "                      AND CONVERT(DATETIME,RAM.ReceiveDate) BETWEEN CONVERT(DATETIME,'" & Format(dtRecDateFrom.Value, "yyyy-MM-dd") & "') AND CONVERT(DATETIME,'" & Format(dtRecDateTo.Value, "yyyy-MM-dd") & "') " & vbCrLf
                'Else
                '    ls_SQL = ls_SQL + _
                '              "                      AND CONVERT(DATETIME,RAM.ReceiveDate) BETWEEN CONVERT(DATETIME,'" & Format(CDate(dt_RecDateFrom), "yyyy-MM-dd") & "') AND CONVERT(DATETIME,'" & Format(CDate(dt_RecDateTo), "yyyy-MM-dd") & "') " & vbCrLf
                'End If
                ls_SQL = ls_SQL + "                      AND (ISNULL(DPM.DeliveryDate,DSM.DeliveryDate) BETWEEN '" & Format(dtRecDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtRecDateTo.Value, "yyyy-MM-dd") & "')"
            End If

            'SUPPLIER CODE/NAME
            If Trim(cboSupplier.Text) <> "==ALL==" And Trim(cboSupplier.Text) <> "" Then
                ls_SQL = ls_SQL + _
                              "                      AND POM.SupplierID = '" & Trim(cboSupplier.Text) & "' " & vbCrLf
            End If
            'PART CODE/NAME
            If Trim(cboPart.Text) <> "==ALL==" And Trim(cboPart.Text) <> "" Then
                ls_SQL = ls_SQL + _
                              "                      AND POD.PartNo = '" & Trim(cboPart.Text) & "' " & vbCrLf
            End If
            'PO NO.
            If Trim(txtPONo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                              "                      AND ISNULL(POM.PONo,'') LIKE '%" & Trim(txtPONo.Text) & "%' " & vbCrLf
            End If
            'GOOD RECEIVING SENT

            ls_SQL = ls_SQL + " ) hdr  " & vbCrLf 
            ls_SQL = ls_SQL + " 		) SPDC " & vbCrLf & _
                              "   ORDER BY SortPONo, SortKanbanNo,SortSupplierSJNo, SortPasiSJNo, SortHeader " & vbCrLf & _
                              "  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 300
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "SELECT TOP 0  " & vbCrLf & _
                     " 		 ColNo = 0, Period = '', PONo = '', DeliveryLocationCode = '', DeliveryLocationName = '', SupplierCode = '', SupplierName = '', POKanban = '',  " & vbCrLf & _
                     " 		 KanbanNo = '', SupplierPlanDeliveryDate = '', SupplierDeliveryDate = '', SupplierSJNO = '', PASIDeliveryDate = '',  " & vbCrLf & _
                     " 		 PASISJNo = '', PartNo = '', PartName = '', UOM = '', SupplierDeliveryQty = '', PASIGoodReceivingQty = '', PASIDefectQty = '',  " & vbCrLf & _
                     " 		 PASIDeliveryQty = '', GoodReceivingQty = '', DefectReceivingQty = '', RemainingReceivingQty = '', ReceivedDate = '', ReceivedBy = '',  " & vbCrLf & _
                     " 		 GoodReceivingSent = '', detail = '', CloseCls = '', CloseDate = '', SupplierPIC = '', url = '', DeliveryByPASICls = '', " & vbCrLf & _
                     " 		 SortPONo = '', SortKanbanNo = '', SortHeader = 0 " & vbCrLf & _
                     "  "
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

        'Combo Supplier
        With cboSupplier
            ls_SQL = "SELECT SupplierID = '==ALL==', SupplierName = '==ALL=='" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     "SELECT SupplierID, SupplierName FROM dbo.MS_Supplier " & vbCrLf
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

        'Combo Parts
        With cboPart
            ls_SQL = "SELECT PartNo = '==ALL==', PartName = '==ALL=='" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     "SELECT PartNo, PartName FROM dbo.MS_Parts"
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
#End Region

#Region "Functions"

#End Region

#Region "Form Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim ls_Param As String = ""

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_FillCombo()

                If Not IsNothing(Session("fromE02")) Then
                    ls_Param = Session("paramLoad")
                    
                    If Split(ls_Param, "|")(0) = "1" Then chkSupplierPeriod.Checked = True Else chkSupplierPeriod.Checked = False
                    Dim ls_SupplierPeriod As String = _
                    "var a = new Date(" & Year((CDate(Split(ls_Param, "|")(1)))) & "," & (Month((CDate(Split(ls_Param, "|")(1)))) - 1) & "," & Day((CDate(Split(ls_Param, "|")(1)))) & "); " & vbCrLf & _
                    "var b = new Date(" & Year((CDate(Split(ls_Param, "|")(7)))) & "," & (Month((CDate(Split(ls_Param, "|")(7)))) - 1) & "," & Day((CDate(Split(ls_Param, "|")(7)))) & "); " & vbCrLf & _
                    "var c = new Date(" & Year((CDate(Split(ls_Param, "|")(8)))) & "," & (Month((CDate(Split(ls_Param, "|")(8)))) - 1) & "," & Day((CDate(Split(ls_Param, "|")(8)))) & "); " & vbCrLf & _
                    "dtSupplierPeriod.SetDate(a); " & vbCrLf & _
                    "dtRecDateFrom.SetDate(b); " & vbCrLf & _
                    "dtRecDateTo.SetDate(c); " & vbCrLf & _
                    "lblInfo.SetText(''); " & vbCrLf
                    ScriptManager.RegisterStartupScript(chkSupplierPeriod, chkSupplierPeriod.GetType(), "Initialize", ls_SupplierPeriod, True)
                    'dtSupplierPeriod.Value = Split(ls_Param, "|")(1)

                    'Log to variable for GridLoad ++++++++++++++++++
                    dt_SupplierPeriod = Split(ls_Param, "|")(1)
                    dt_RecDateFrom = Split(ls_Param, "|")(7)
                    dt_RecDateTo = Split(ls_Param, "|")(8)
                    '+++++++++++++++++++++++++++++++++++++++++++++++

                    If Split(ls_Param, "|")(2) = "0" Then rdrSADAll.Checked = True Else rdrSADAll.Checked = False
                    If Split(ls_Param, "|")(2) = "1" Then rdrSADYes.Checked = True Else rdrSADYes.Checked = False
                    If Split(ls_Param, "|")(2) = "2" Then rdrSADNo.Checked = True Else rdrSADNo.Checked = False

                    If Split(ls_Param, "|")(3) = "0" Then rdrPADAll.Checked = True Else rdrPADAll.Checked = False
                    If Split(ls_Param, "|")(3) = "1" Then rdrPADYes.Checked = True Else rdrPADYes.Checked = False
                    If Split(ls_Param, "|")(3) = "2" Then rdrPADNo.Checked = True Else rdrPADNo.Checked = False

                    If Split(ls_Param, "|")(4) = "0" Then rdrRRQAll.Checked = True Else rdrRRQAll.Checked = False
                    If Split(ls_Param, "|")(4) = "1" Then rdrRRQYes.Checked = True Else rdrRRQYes.Checked = False
                    If Split(ls_Param, "|")(4) = "2" Then rdrRRQNo.Checked = True Else rdrRRQNo.Checked = False

                    txtSupplierSJNo.Text = Split(ls_Param, "|")(5)

                    If Split(ls_Param, "|")(6) = "1" Then chkRecDate.Checked = True Else chkRecDate.Checked = False
                    'Dim ls_RecDateFrom As String = _
                    '"var b = new Date(" & Year((CDate(Split(ls_Param, "|")(7)))) & "," & (Month((CDate(Split(ls_Param, "|")(7)))) - 1) & "," & Day((CDate(Split(ls_Param, "|")(7)))) & "); " & vbCrLf & _
                    '"dtRecDateFrom.SetDate(b); " & vbCrLf
                    'ScriptManager.RegisterStartupScript(dtRecDateFrom, dtRecDateFrom.GetType(), "Initialize", ls_RecDateFrom, True)
                    ''dtRecDateFrom.Value = Split(ls_Param, "|")(7)
                    'Dim ls_RecDateTo As String = _
                    '"var c = new Date(" & Year((CDate(Split(ls_Param, "|")(8)))) & "," & (Month((CDate(Split(ls_Param, "|")(8)))) - 1) & "," & Day((CDate(Split(ls_Param, "|")(8)))) & "); " & vbCrLf & _
                    '"dtRecDateTo.SetDate(c); " & vbCrLf
                    'ScriptManager.RegisterStartupScript(dtRecDateTo, dtRecDateTo.GetType(), "Initialize", ls_RecDateTo, True)
                    ''dtRecDateTo.Value = Split(ls_Param, "|")(8)

                    cboSupplier.Text = Split(ls_Param, "|")(9)
                    txtSupplierName.Text = Split(ls_Param, "|")(10)

                    cboPart.Text = Split(ls_Param, "|")(11)
                    txtPartName.Text = Split(ls_Param, "|")(12)

                    txtPONo.Text = Split(ls_Param, "|")(13)

                    If Split(ls_Param, "|")(14) = "0" Then rdrSDAll.Checked = True Else rdrSDAll.Checked = False
                    If Split(ls_Param, "|")(14) = "1" Then rdrSDDirect.Checked = True Else rdrSDDirect.Checked = False
                    If Split(ls_Param, "|")(14) = "2" Then rdrSDPasi.Checked = True Else rdrSDPasi.Checked = False

                    If Split(ls_Param, "|")(15) = "0" Then rdrPOKAll.Checked = True Else rdrPOKAll.Checked = False
                    If Split(ls_Param, "|")(15) = "1" Then rdrPOKYes.Checked = True Else rdrPOKYes.Checked = False
                    If Split(ls_Param, "|")(15) = "2" Then rdrPOKNo.Checked = True Else rdrPOKNo.Checked = False

                    If Split(ls_Param, "|")(16) = "0" Then rdrMCPAll.Checked = True Else rdrMCPAll.Checked = False
                    If Split(ls_Param, "|")(16) = "1" Then rdrMCPYes.Checked = True Else rdrMCPYes.Checked = False
                    If Split(ls_Param, "|")(16) = "2" Then rdrMCPNo.Checked = True Else rdrMCPNo.Checked = False

                    If Split(ls_Param, "|")(17) = "0" Then rdrGRSAll.Checked = True Else rdrGRSAll.Checked = False
                    If Split(ls_Param, "|")(17) = "1" Then rdrGRSYes.Checked = True Else rdrGRSYes.Checked = False
                    If Split(ls_Param, "|")(17) = "2" Then rdrGRSNo.Checked = True Else rdrGRSNo.Checked = False

                    Call up_GridLoad()

                    'Session.Remove("paramLoad")
                    Session.Remove("fromE02")

                Else
                    Call up_GridLoadWhenEventChange()
                    Call up_Initialize()
                End If

            Else
                'SAVE PARAMETER
                If chkSupplierPeriod.Checked = True Then ls_Param = "1|" Else ls_Param = "0|"
                ls_Param = ls_Param & Format(dtSupplierPeriod.Value, "dd MMMM yyyy") & "|"

                If rdrSADAll.Checked = True Then ls_Param = ls_Param & "0|"
                If rdrSADYes.Checked = True Then ls_Param = ls_Param & "1|"
                If rdrSADNo.Checked = True Then ls_Param = ls_Param & "2|"

                If rdrPADAll.Checked = True Then ls_Param = ls_Param & "0|"
                If rdrPADYes.Checked = True Then ls_Param = ls_Param & "1|"
                If rdrPADNo.Checked = True Then ls_Param = ls_Param & "2|"

                If rdrRRQAll.Checked = True Then ls_Param = ls_Param & "0|"
                If rdrRRQYes.Checked = True Then ls_Param = ls_Param & "1|"
                If rdrRRQNo.Checked = True Then ls_Param = ls_Param & "2|"

                ls_Param = ls_Param & Trim(txtSupplierSJNo.Text) & "|"

                If chkRecDate.Checked = True Then ls_Param = ls_Param & "1|" Else ls_Param = ls_Param & "0|"
                ls_Param = ls_Param & Format(dtRecDateFrom.Value, "dd MMMM yyyy") & "|"
                ls_Param = ls_Param & Format(dtRecDateTo.Value, "dd MMMM yyyy") & "|"

                ls_Param = ls_Param & Trim(cboSupplier.Text) & "|"
                ls_Param = ls_Param & Trim(txtSupplierName.Text) & "|"

                ls_Param = ls_Param & Trim(cboPart.Text) & "|"
                ls_Param = ls_Param & Trim(txtPartName.Text) & "|"

                ls_Param = ls_Param & Trim(txtPONo.Text) & "|"

                If rdrSDAll.Checked = True Then ls_Param = ls_Param & "0|"
                If rdrSDDirect.Checked = True Then ls_Param = ls_Param & "1|"
                If rdrSDPasi.Checked = True Then ls_Param = ls_Param & "2|"

                If rdrPOKAll.Checked = True Then ls_Param = ls_Param & "0|"
                If rdrPOKYes.Checked = True Then ls_Param = ls_Param & "1|"
                If rdrPOKNo.Checked = True Then ls_Param = ls_Param & "2|"

                If rdrMCPAll.Checked = True Then ls_Param = ls_Param & "0|"
                If rdrMCPYes.Checked = True Then ls_Param = ls_Param & "1|"
                If rdrMCPNo.Checked = True Then ls_Param = ls_Param & "2|"

                If rdrGRSAll.Checked = True Then ls_Param = ls_Param & "0|"
                If rdrGRSYes.Checked = True Then ls_Param = ls_Param & "1|"
                If rdrGRSNo.Checked = True Then ls_Param = ls_Param & "2|"

                Session("paramLoad") = ls_Param
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("E01Msg")
        Session.Remove("paramLoad")
        Session.Remove("fromE02")

        Response.Redirect("~/MainMenu.aspx")
    End Sub

    'Private Sub grid_BatchUpdate(sender As Object, e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
    '    Try
    '        Dim ls_CloseCls As String = "0", ls_SupplierPIC As String = "", _
    '            ls_PONo As String = "", ls_AffilateID As String = "", ls_SupplierID As String = ""
    '        Dim iLoop As Integer = 0

    '        Using scope As New TransactionScope

    '            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '                sqlConn.Open()

    '                For iLoop = 0 To e.UpdateValues.Count
    '                    ls_CloseCls = e.UpdateValues(iLoop).NewValues("CloseCls").ToString()
    '                    ls_PONo = e.UpdateValues(iLoop).NewValues("PONo").ToString()
    '                    ls_SupplierID = e.UpdateValues(iLoop).NewValues("SupplierCode").ToString()

    '                    ls_SQL = "UPDATE dbo.PO_Master " & vbCrLf & _
    '                             "   SET CloseCls = '" & ls_CloseCls & "', CloseDate = GETDATE(), CloseSupplierPIC = '" & ls_SupplierPIC & "' " & vbCrLf & _
    '                             " WHERE PONo = '" & ls_PONo & "'" & vbCrLf & _
    '                             "   AND SupplierID = '" & ls_SupplierID & "'" & vbCrLf & _
    '                             "   AND AffilateID = '" & ls_AffilateID & "'"
    '                    Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
    '                    sqlCmd.ExecuteNonQuery()
    '                    sqlCmd.Dispose()
    '                Next iLoop

    '            End Using
    '            scope.Complete()
    '        End Using

    '    Catch ex As Exception
    '        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '        Session("E01Msg") = lblInfo.Text
    '    End Try
    'End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        Session("E01Msg") = lblInfo.Text
                    Else
                        grid.PageIndex = 0
                    End If
                Case "clear"
                    Call up_GridLoadWhenEventChange()
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("E01Msg") = lblInfo.Text
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
        If (Not IsNothing(Session("E01Msg"))) Then grid.JSProperties("cpMessage") = Session("E01Msg") : Session.Remove("E01Msg")
        grid.SettingsPager.PageSize = grid.VisibleRowCount + 1
    End Sub

    Private Sub grid_CustomColumnDisplayText(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewColumnDisplayTextEventArgs) Handles grid.CustomColumnDisplayText
        If (e.Column.FieldName = "SupplierDeliveryQty" Or e.Column.FieldName = "PASIGoodReceivingQty" Or _
            e.Column.FieldName = "PASIDefectQty" Or e.Column.FieldName = "PASIDeliveryQty" Or _
            e.Column.FieldName = "GoodReceivingQty" Or e.Column.FieldName = "DefectReceivingQty" Or _
            e.Column.FieldName = "RemainingReceivingQty") And e.GetFieldValue("ColNo") = "" Then

            Dim ls_Value As String = e.GetFieldValue(e.Column.FieldName)
            If IsNothing(ls_Value) Then ls_Value = "0"
            e.DisplayText = FormatNumber(ls_Value.Trim, 0, TriState.True)
            If ls_Value = "" Then e.DisplayText = "0"
        End If
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        Select Case e.DataColumn.FieldName
            Case "url"
                If e.GetValue("SortHeader") = "0" Then 'HEADER DATA
                    If e.GetValue("DeliveryByPASICls") = "1" And e.GetValue("PASISJNo").ToString().Trim() = "" Then
                        'DELIVERY BY PASI (SUPPLIER to PASI to AFFILIATE)
                        e.Cell.Controls.Item("0").Controls.Clear()
                    ElseIf e.GetValue("DeliveryByPASICls") = "0" And e.GetValue("SupplierSJNo").ToString().Trim() = "" Then
                        'DIRECT DELIVERY (SUPPLIER to AFFILIATE)
                        e.Cell.Controls.Item("0").Controls.Clear()
                    End If
                End If

            Case "RemainingReceivingQty"
                If e.GetValue("SortHeader") = "1" Then 'DETAIL DATA
                    If e.GetValue("SortDeliveryByPASICls") = "1" Then
                        'DELIVERY BY PASI (SUPPLIER to PASI to AFFILIATE)
                        If CDbl(e.GetValue("PASIDeliveryQty")) > (CDbl(e.GetValue("GoodReceivingQty")) + CDbl(e.GetValue("DefectReceivingQty"))) Then
                            e.Cell.BackColor = Color.Fuchsia
                        End If

                    ElseIf e.GetValue("SortDeliveryByPASICls") = "0" Then
                        'DIRECT DELIVERY (SUPPLIER to AFFILIATE)
                        If CDbl(e.GetValue("SupplierDeliveryQty")) > (CDbl(e.GetValue("GoodReceivingQty")) + CDbl(e.GetValue("DefectReceivingQty"))) Then
                            e.Cell.BackColor = Color.Fuchsia
                        End If
                    End If
                End If

        End Select
    End Sub

    Private Sub grid_HtmlRowPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowPrepared
        Try
            Dim getRowValues As String = e.GetValue("PONo")
            If Not IsNothing(getRowValues) Then
                If getRowValues.Trim() <> "" Then
                    e.Row.BackColor = Color.FromName("#E0E0E0")
                End If
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("E01Msg") = lblInfo.Text
        End Try
    End Sub

    Private Sub grid_PageIndexChanged(sender As Object, e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region
    
End Class