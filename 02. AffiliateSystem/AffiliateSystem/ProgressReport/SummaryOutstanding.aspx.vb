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
    Const colPOKanban As Byte = 4
    Const colKanbanNo As Byte = 5
    Const colPartNo As Byte = 6
    Const colPartName As Byte = 7
    Const colQtyPO As Byte = 8
    Const colRemainingQtyPASI As Byte = 9
    Const colPASIDelDate As Byte = 10
    Const colPASISJNo As Byte = 11
    Const colPASIDeliveryQty As Byte = 12
    Const colAffiliateRecDate As Byte = 13
    Const colAffiliateReceivingQty As Byte = 14
    Const colPASIInvNo As Byte = 15
    Const colPASIInvDate As Byte = 16
    Const colPASIInvCurr As Byte = 17
    Const colPASIInvAmount As Byte = 18

    Const colCount As Byte = 18

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "Procedures"
    Private Sub up_Initialize()
        Dim script As String = _
            "var PeriodTo = new Date(); " & vbCrLf & _
            "var PeriodFrom = new Date(PeriodTo.getFullYear(),PeriodTo.getMonth(),1); " & vbCrLf & _
            "dtPOPeriodFrom.SetDate(PeriodFrom); " & vbCrLf & _
            "dtPOPeriodTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtPASIDelDateFrom.SetDate(PeriodFrom); " & vbCrLf & _
            "dtPASIDelDateTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtAffiliateRecDateFrom.SetDate(PeriodFrom); " & vbCrLf & _
            "dtAffiliateRecDateTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtPASIInvDateFrom.SetDate(PeriodFrom); " & vbCrLf & _
            "dtPASIInvDateTo.SetDate(PeriodTo); " & vbCrLf & _
            " " & vbCrLf & _
            "txtPONo.SetText(''); " & vbCrLf & _
            "txtPASISJNo.SetText(''); " & vbCrLf & _
            "txtPASIInvNo.SetText(''); " & vbCrLf & _
            " " & vbCrLf & _
            "chkPASIDelDate.SetValue(false); " & vbCrLf & _
            "chkAffiliateRecDate.SetValue(false); " & vbCrLf & _
            "chkPASIInvDate.SetValue(false); " & vbCrLf & _
            " " & vbCrLf & _            
            "if (cboPart.GetItemCount() > 1) { " & vbCrLf & _
            "   txtPartName.SetText('==ALL=='); " & vbCrLf & _
            "   cboPart.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(chkPOPeriod, chkPOPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""
            Dim ls_Filter As String = ""

            Dim ls_End As String = ""
            ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

            ''PO PERIOD
            'If chkPOPeriod.Checked = True Then
            '    ls_Filter = ls_Filter + _
            '                  "                      AND CONVERT(CHAR(8),POM.Period,112) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyyMM01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyyMM" & ls_End) & "' " & vbCrLf
            'End If

            ''PASI DELIVERY DATE
            'If chkPASIDelDate.Checked = True Then
            '    ls_Filter = ls_Filter + _
            '                  "                      AND CONVERT(CHAR(8),PDM.DeliveryDate,112) BETWEEN '" & Format(dtPASIDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
            'End If

            ''AFFILIATE RECEIVE DATE
            'If chkAffiliateRecDate.Checked = True Then
            '    ls_Filter = ls_Filter + _
            '                  "                      AND CONVERT(CHAR(8),RAM.ReceiveDate,112) BETWEEN '" & Format(dtAffiliateRecDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtAffiliateRecDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
            'End If

            ''PASI INVOICE DATE
            'If chkPASIInvDate.Checked = True Then
            '    ls_Filter = ls_Filter + _
            '                  "                      AND CONVERT(CHAR(8),IPM.DeliveryDate,112) BETWEEN '" & Format(dtPASIInvDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIInvDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
            'End If

            ''PART CODE
            'If Trim(cboPart.Text) <> "==ALL==" And Trim(cboPart.Text) <> "" Then
            '    ls_Filter = ls_Filter + _
            '                  "                      AND POD.PartNo = '" & Trim(cboPart.Text) & "' " & vbCrLf
            'End If

            ''PONO
            'If Trim(txtPONo.Text) <> "" Then
            '    ls_Filter = ls_Filter + _
            '                  "                      AND ISNULL(POM.PONo,'') LIKE '%" & Trim(txtPONo.Text) & "%' " & vbCrLf
            'End If

            ''PASI SJ NO
            'If Trim(txtPASISJNo.Text) <> "" Then
            '    ls_Filter = ls_Filter + _
            '                  "                      AND ISNULL(PDM.SuratJalanNo,'') LIKE '%" & Trim(txtPASISJNo.Text) & "%'" & vbCrLf
            'End If

            ''PASI INV NO
            'If Trim(txtPASIInvNo.Text) <> "" Then
            '    ls_Filter = ls_Filter + _
            '                  "                      AND ISNULL(IPM.InvoiceNo,'') LIKE '%" & Trim(txtPASIInvNo.Text) & "%'" & vbCrLf
            'End If

            'ls_SQL = "  SELECT  ColNo = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY Period, AffiliateCode, SupplierCode, PONo, KanbanNo, PartNo)) ,    " & vbCrLf & _
            '      "  		Period ,    " & vbCrLf & _
            '      " 		PONo , " & vbCrLf & _
            '      "  		SupplierCode ,    		 " & vbCrLf & _
            '      "  		POKanban = CASE WHEN ISNULL(KanbanCls, '0') = '1' THEN 'YES'    " & vbCrLf & _
            '      "  						ELSE 'NO'    " & vbCrLf & _
            '      "  					END ,    " & vbCrLf & _
            '      "  		KanbanNo ,  " & vbCrLf & _
            '      " 		PartNo ,    " & vbCrLf & _
            '      "  		PartName ,    " & vbCrLf & _
            '      " 		QtyPO = QtyPO ,    "

            'ls_SQL = ls_SQL + " 		PASIDeliveryDate ,    " & vbCrLf & _
            '                  "  		PASISJNo ,    " & vbCrLf & _
            '                  " 		PASIDeliveryQty ,    " & vbCrLf & _
            '                  "  		AffiliateReceiveDate ,    		  				 " & vbCrLf & _
            '                  "  		AffiliateReceivingQty ,     " & vbCrLf & _
            '                  "  		InvoiceToAffiliateQty ,      " & vbCrLf & _
            '                  "  		InvoiceNoToAffiliate ,    " & vbCrLf & _
            '                  "  		InvoiceDateToAffiliate ,    		 " & vbCrLf & _
            '                  " 		InvoiceToAffiliateCurr ,    " & vbCrLf & _
            '                  "  		InvoiceToAffiliateAmount " & vbCrLf & _
            '                  "  FROM    ( SELECT DISTINCT    "

            'ls_SQL = ls_SQL + "  			Period = SUBSTRING(CONVERT(CHAR, POM.Period, 106), 4, 9) ,    " & vbCrLf & _
            '                  "  			PONo = POM.PONo ,    " & vbCrLf & _
            '                  "  			AffiliateCode = POM.AffiliateID ,    " & vbCrLf & _
            '                  "  			AffiliateName = MA.AffiliateName ,    " & vbCrLf & _
            '                  "  			SupplierCode = POM.SupplierID ,    " & vbCrLf & _
            '                  "  			SupplierName = MS.SupplierName ,    " & vbCrLf & _
            '                  "  			KanbanCls = ISNULL(POD.KanbanCls, '0') ,    			KanbanNo = ISNULL(KD.KanbanNo,''),      " & vbCrLf & _
            '                  "  			SupplierPlanDeliveryDate = '',      " & vbCrLf & _
            '                  "  			SupplierDeliveryDate = '',      " & vbCrLf & _
            '                  "  			SupplierSJNO = '',      " & vbCrLf & _
            '                  "  			PASIReceiveDate = '',  "

            'ls_SQL = ls_SQL + "              PASIDeliveryDate = ISNULL(CONVERT(CHAR,PDM.DeliveryDate,106),'') ,  " & vbCrLf & _
            '                  "              PASISJNo = ISNULL(PDM.SuratJalanNo,'') ,    " & vbCrLf & _
            '                  "              AffiliateReceiveDate = ISNULL(CONVERT(CHAR,RAM.ReceiveDate,106),'') ,  " & vbCrLf & _
            '                  "              PartNo = POD.PartNo,      " & vbCrLf & _
            '                  "  			PartName = MP.PartName,      " & vbCrLf & _
            '                  "  			UOM = UC.Description,     			SupplierDeliveryQty = 0,      " & vbCrLf & _
            '                  "  			PASIReceivingQty = 0,     " & vbCrLf & _
            '                  "  			PASIDeliveryQty = ISNULL(PDD.DOQty,0),      " & vbCrLf & _
            '                  "  			AffiliateReceivingQty = ISNULL(RAD.RecQty,0),      " & vbCrLf & _
            '                  "  			InvoiceFromSupplierQty = 0,       " & vbCrLf & _
            '                  "  			InvoiceToAffiliateQty = ISNULL(IPD.DOQty,0),      "

            'ls_SQL = ls_SQL + "  			InvoiceNoFromSupplier = '',     " & vbCrLf & _
            '                  "  			InvoiceDateFromSupplier = '',     " & vbCrLf & _
            '                  "  			InvoiceFromSupplierCurr = '',     " & vbCrLf & _
            '                  "  			InvoiceFromSupplierAmount = '',     " & vbCrLf & _
            '                  "  			InvoiceNoToAffiliate = ISNULL(IPM.InvoiceNo,''),      			InvoiceDateToAffiliate = ISNULL(CONVERT(CHAR,IPM.DeliveryDate,106),''),     " & vbCrLf & _
            '                  "  			InvoiceToAffiliateCurr = 'IDR',     " & vbCrLf & _
            '                  "  			InvoiceToAffiliateAmount = ISNULL(IPD.DOQty,0) * ISNULL(MSP.Price,0),      " & vbCrLf & _
            '                  "              PODelivery = 0 ,    " & vbCrLf & _
            '                  "              RemainingQtyPOPASI = 0 ,    " & vbCrLf & _
            '                  "              RemainingQtyPOSupplier = 0,   " & vbCrLf & _
            '                  "              QtyPO = ISNULL(POD.POQty,0)  ,    "

            'ls_SQL = ls_SQL + "              h_affiliateorder = KM.AffiliateID ,    " & vbCrLf & _
            '                  "              H_PasiSJ = PDD.SuratJalanNo ,    " & vbCrLf & _
            '                  "              h_poorder = KD.PONo ,    " & vbCrLf & _
            '                  "              h_idxorder = '0' ,                h_kanbanorder = KM.KanbanNo,h_partno ='0'    " & vbCrLf & _
            '                  "              FROM      dbo.PO_Master POM    " & vbCrLf & _
            '                  "                        LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
            '                  "                                                   AND POM.PONo = POD.PONo    " & vbCrLf & _
            '                  "                                                   AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID    " & vbCrLf & _
            '                  "                                                          AND KD.PoNo = POD.PONo    " & vbCrLf & _
            '                  "                                                          AND KD.SupplierID = POD.SupplierID    "

            'ls_SQL = ls_SQL + "                                                          AND KD.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf & _
            '                  "                                                          AND KD.KanbanNo = KM.KanbanNo                                                            AND KD.SupplierID = KM.SupplierID    " & vbCrLf & _
            '                  "                                                          AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
            '                  "                        LEFT JOIN DOPasi_Detail PDD ON KD.AffiliateID = PDD.AffiliateID    " & vbCrLf & _
            '                  "                                           AND KD.KanbanNo = PDD.KanbanNo    " & vbCrLf & _
            '                  "                                           AND KD.SupplierID = PDD.SupplierID    " & vbCrLf & _
            '                  "                                           AND KD.PartNo = PDD.PartNo    " & vbCrLf & _
            '                  "                                           AND KD.PoNo = PDD.PoNo       " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf & _
            '                  "                                                           AND PDD.SuratJalanNo = PDM.SuratJalanNo          "

            'ls_SQL = ls_SQL + "                        LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
            '                  "                                                                  AND PDD.KanbanNo = RAD.KanbanNo                                                                    AND PDD.SupplierID = RAD.SupplierID    " & vbCrLf & _
            '                  "                                                                  AND PDD.PartNo = RAD.PartNo    " & vbCrLf & _
            '                  "                                                                  AND PDD.PoNo = RAD.PONo    " & vbCrLf & _
            '                  "  																AND PDD.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo    " & vbCrLf & _
            '                  "                                                                  AND RAM.AffiliateID = RAD.AffiliateID         " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.PLPASI_Detail IPD ON PDD.AffiliateID = IPD.AffiliateID    " & vbCrLf & _
            '                  "                                                                AND PDD.KanbanNo = IPD.KanbanNo    " & vbCrLf & _
            '                  "                                                                AND PDD.PartNo = IPD.PartNo    " & vbCrLf & _
            '                  "                                                                AND PDD.PONo = IPD.PONo    "

            'ls_SQL = ls_SQL + "                                                                AND PDD.SuratJalanNo = IPD.SuratJalanNo                          LEFT JOIN dbo.PLPASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID                                                                  " & vbCrLf & _
            '                  "                                                                AND IPD.SuratJalanNo = IPM.SuratJalanNo    		  " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls    " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID    " & vbCrLf & _
            '                  "                        LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode                          " & vbCrLf & _
            '                  "                        LEFT JOIN MS_Price MSP ON MSP.AffiliateID = IPD.AffiliateID and MSP.PartNo = IPD.PartNo and (IPM.DeliveryDate between MSP.StartDate and MSP.EndDate)                         " & vbCrLf & _
            '                  "              WHERE     KD.KanbanQty <> 0  AND POM.AffiliateID = '" & Session("AffiliateID") & "'  " & vbCrLf

            'ls_SQL = ls_SQL + ls_Filter & vbCrLf

            'ls_SQL = ls_SQL + "  " & vbCrLf & _
            '                  "          ) Hdr   " & vbCrLf & _
            '                  " ORDER BY CONVERT(numeric,CONVERT(char, ROW_NUMBER() OVER ( ORDER BY Period, AffiliateCode, SupplierCode, PONo, KanbanNo, PartNo)))           " & vbCrLf & _
            '                  "  "


            ls_SQL = "sp_Affiliate_SummaryOutstanding_GridLoad"
            Dim cmd As New SqlCommand(ls_SQL, sqlConn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@AffiliateID", Session("AffiliateID"))

            'PO PERIOD
            If chkPOPeriod.Checked = True Then
                cmd.Parameters.AddWithValue("@POPeriodFrom", Format(dtPOPeriodFrom.Value, "yyyyMM01"))
                cmd.Parameters.AddWithValue("@POPeriodTo", Format(dtPOPeriodTo.Value, "yyyyMM" & ls_End))
            End If

            'PASI DELIVERY DATE
            If chkPASIDelDate.Checked = True Then
                cmd.Parameters.AddWithValue("@PASIDeliveryDateFrom", Format(dtPASIDelDateFrom.Value, "yyyyMMdd"))
                cmd.Parameters.AddWithValue("@PASIDeliveryDateTo", Format(dtPASIDelDateTo.Value, "yyyyMMdd"))
            End If

            'AFFILIATE RECEIVE DATE
            If chkAffiliateRecDate.Checked = True Then
                cmd.Parameters.AddWithValue("@ReceiveDateFrom", Format(dtAffiliateRecDateFrom.Value, "yyyyMMdd"))
                cmd.Parameters.AddWithValue("@ReceiveDateTo", Format(dtAffiliateRecDateTo.Value, "yyyyMMdd"))
            End If

            'PASI INVOICE DATE
            If chkPASIInvDate.Checked = True Then
                cmd.Parameters.AddWithValue("@PASIInvoiceDateFrom", Format(dtPASIInvDateFrom.Value, "yyyyMMdd"))
                cmd.Parameters.AddWithValue("@PASIInvoiceDateTo", Format(dtPASIInvDateTo.Value, "yyyyMMdd"))
            End If

            'PART CODE , PONo ,PASI SJ NO,PASI INV NO
            cmd.Parameters.AddWithValue("@PartNo", Trim(cboPart.Text))
            cmd.Parameters.AddWithValue("@PONo", Trim(txtPONo.Text))
            cmd.Parameters.AddWithValue("@PASISJNo", Trim(txtPASISJNo.Text))
            cmd.Parameters.AddWithValue("@PASIInvNo", Trim(txtPASIInvNo.Text))

            Dim sqlDA As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
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

            ls_SQL = "SELECT TOP 0 " & vbCrLf & _
                     " 	     ColNo = 0,  " & vbCrLf & _
                     "       Period = '', PONo = '', SupplierCode = '', SupplierName = '', POKanban = '',   " & vbCrLf & _
                     " 	     KanbanNo = '', /*KanbanSeqNo = '',*/ SupplierPlanDeliveryDate = '', SupplierDeliveryDate = '', SupplierSJNo = '', PASIReceiveDate = '', PASIDeliveryDate = '',  " & vbCrLf & _
                     " 	     PASISJNo = '', AffiliateReceiveDate = '', PartNo = '', PartName = '', UOM = '', SupplierDeliveryQty = '', PASIReceivingQty = '', PASIDeliveryQty = '', AffiliateReceivingQty = '', PASIInvoiceQty = '',  " & vbCrLf & _
                     " 	     PASIInvoiceNo = '', PASIInvoiceDate = '', PASIInvoiceCurr = '', PASIInvoiceAmount = '', " & vbCrLf & _
                     " 	     SortPONo = '', SortKanbanNo = '', SortHeader = 0, PODelivery = '', PODelivery2 = '', PASIReceiveDate2 = '' "

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
            Dim irow As Integer = 0
            Dim icol As Integer = 0

            With ws
                .Cells(3, 4).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
                .Cells(4, 4).Value = ": " & Trim(Session("AffiliateID"))

                For irow = 0 To pData.Rows.Count - 1
                    For icol = 1 To pData.Columns.Count
                        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                    Next
                Next

                Dim rgAll As ExcelRange = .Cells(8, 1, irow + 8, 18)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

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

                'PO PERIOD
                If chkPOPeriod.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(CHAR(8),POM.Period,112) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyyMM01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyyMM" & ls_End) & "' " & vbCrLf
                End If

                'PASI DELIVERY DATE
                If chkPASIDelDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(CHAR(8),PDM.DeliveryDate,112) BETWEEN '" & Format(dtPASIDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If

                'AFFILIATE RECEIVE DATE
                If chkAffiliateRecDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(CHAR(8),RAM.ReceiveDate,112) BETWEEN '" & Format(dtAffiliateRecDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtAffiliateRecDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If

                'PASI INVOICE DATE
                If chkPASIInvDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(CHAR(8),IPM.DeliveryDate,112) BETWEEN '" & Format(dtPASIInvDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIInvDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
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

                'PASI SJ NO
                If Trim(txtPASISJNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(PDM.SuratJalanNo,'') LIKE '%" & Trim(txtPASISJNo.Text) & "%'" & vbCrLf
                End If

                'PASI INV NO
                If Trim(txtPASIInvNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(IPM.InvoiceNo,'') LIKE '%" & Trim(txtPASIInvNo.Text) & "%'" & vbCrLf
                End If

                '------------------------------- QUERY --------------------------------------
                ls_sql = " SELECT  ColNo = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY Period, AffiliateCode, SupplierCode, PONo, KanbanNo, PartNo)) ,  " & vbCrLf & _
                      "          Period ,  " & vbCrLf & _
                      "          PONo ,  " & vbCrLf & _
                      "          SupplierCode ,  " & vbCrLf & _
                      "          POKanban = CASE WHEN ISNULL(KanbanCls, '0') = '1' THEN 'YES'  " & vbCrLf & _
                      "                          ELSE 'NO'  " & vbCrLf & _
                      "                     END ,  " & vbCrLf & _
                      "          KanbanNo ,  " & vbCrLf

                ls_sql = ls_sql + "          PartNo ,  " & vbCrLf & _
                                  "          PartName ,  " & vbCrLf & _
                                  "          QtyPO = QtyPO ,  " & vbCrLf & _
                                  "          PASIDeliveryDate ,  " & vbCrLf & _
                                  "          PASISJNo ,  " & vbCrLf & _
                                  "          PASIDeliveryQty ,  " & vbCrLf & _
                                  "          AffiliateReceiveDate ,  " & vbCrLf

                ls_sql = ls_sql + "          AffiliateReceivingQty ,  " & vbCrLf & _
                                  "          InvoiceNoToAffiliate ,  " & vbCrLf & _
                                  "          InvoiceDateToAffiliate ,  " & vbCrLf

                ls_sql = ls_sql + "          InvoiceToAffiliateCurr ,  " & vbCrLf & _
                                  "          InvoiceToAffiliateAmount   " & vbCrLf

                ls_sql = ls_sql + "  FROM    ( SELECT DISTINCT  " & vbCrLf & _
                                  "                      Period = SUBSTRING(CONVERT(CHAR, POM.Period, 106), 4, 9) ,  " & vbCrLf & _
                                  "                      PONo = POD.PONo ,  " & vbCrLf & _
                                  "                      AffiliateCode = MA.AffiliateID ,  " & vbCrLf & _
                                  "                      AffiliateName = MA.AffiliateName ,  " & vbCrLf & _
                                  "                      SupplierCode = POM.SupplierID ,  " & vbCrLf & _
                                  "                      SupplierName = MS.SupplierName ,  " & vbCrLf

                ls_sql = ls_sql + "                      KanbanCls = ISNULL(POD.KanbanCls, '0') ,  " & vbCrLf & _
                                  "                      KanbanNo = ISNULL(KD.KanbanNo,''),    " & vbCrLf & _
                                  "  					 SupplierPlanDeliveryDate = ISNULL(CONVERT(CHAR,SDM.DeliveryDate,106),''),    " & vbCrLf & _
                                  "                      SupplierDeliveryDate = ISNULL(CONVERT(CHAR,SDM.DeliveryDate,106),''),    " & vbCrLf & _
                                  "  				 	 SupplierSJNO = ISNULL(SDD.SuratJalanNo,''),    " & vbCrLf & _
                                  "  				     PASIReceiveDate = ISNULL(CONVERT(CHAR,PRM.ReceiveDate,106),''),   " & vbCrLf & _
                                  "                      PASIDeliveryDate = ISNULL(CONVERT(CHAR,PDM.DeliveryDate,106),''),  " & vbCrLf & _
                                  "                      PASISJNo = ISNULL(PDM.SuratJalanNo,'') ,  " & vbCrLf & _
                                  "                      AffiliateReceiveDate = ISNULL(CONVERT(CHAR,RAM.ReceiveDate,106),'') ,  " & vbCrLf & _
                                  "                      PartNo = POD.PartNo,    " & vbCrLf & _
                                  "  					 PartName = MP.PartName,    " & vbCrLf

                ls_sql = ls_sql + "  					 UOM = UC.Description,   " & vbCrLf & _
                                  "                      SupplierDeliveryQty = ISNULL(SDD.DOQty,0),    " & vbCrLf & _
                                  "  					 PASIReceivingQty = ISNULL(PRD.GoodRecQty,0),   " & vbCrLf & _
                                  "  					 PASIDeliveryQty = ISNULL(PDD.DOQty,0),    " & vbCrLf & _
                                  "  					 AffiliateReceivingQty = ISNULL(RAD.RecQty,0),    " & vbCrLf & _
                                  "  					 InvoiceFromSupplierQty = ISNULL(ISD.InvQty,0),     " & vbCrLf & _
                                  "  					 InvoiceToAffiliateQty = ISNULL(IPD.InvQty,0),    " & vbCrLf & _
                                  "  					 InvoiceNoFromSupplier = ISNULL(ISM.InvoiceNo,''),   " & vbCrLf & _
                                  "  					 InvoiceDateFromSupplier = ISNULL(CONVERT(CHAR,ISM.InvoiceDate,106),''),   " & vbCrLf & _
                                  "  					 InvoiceFromSupplierCurr = ISNULL(MCS.Description,''),   " & vbCrLf & _
                                  "  					 InvoiceFromSupplierAmount = ISNULL(ISD.InvAmount,0),   " & vbCrLf

                ls_sql = ls_sql + "  					 InvoiceNoToAffiliate = ISNULL(IPM.InvoiceNo,''),    " & vbCrLf & _
                                  "  					 InvoiceDateToAffiliate = ISNULL(CONVERT(CHAR,IPM.InvoiceDate,106),''),   " & vbCrLf & _
                                  "  					 InvoiceToAffiliateCurr = ISNULL(MC.Description,''),   " & vbCrLf & _
                                  "  					 InvoiceToAffiliateAmount = ISNULL(IPD.InvAmount,0),    " & vbCrLf & _
                                  "                      PODelivery = 0 ,  " & vbCrLf & _
                                  "                      RemainingQtyPOPASI = 0 ,  " & vbCrLf & _
                                  "                      RemainingQtyPOSupplier = ISNULL(SDD.DOQty,0), " & vbCrLf & _
                                  "                      QtyPO = ISNULL(POD.POQty,0)  ,  " & vbCrLf & _
                                  "                      h_affiliateorder = KM.AffiliateID ,  " & vbCrLf & _
                                  "                      H_PasiSJ = PDD.SuratJalanNo ,  " & vbCrLf & _
                                  "                      h_poorder = KD.PONo ,  " & vbCrLf

                ls_sql = ls_sql + "                      h_idxorder = '0' ,  " & vbCrLf & _
                                  "                      h_kanbanorder = KM.KanbanNo,h_partno ='0'  " & vbCrLf & _
                                  "            FROM      dbo.PO_Master POM  " & vbCrLf & _
                                  "                      LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                                  "                                                 AND POM.PoNo = POD.PONo  " & vbCrLf & _
                                  "                                                 AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                                  "                                                        AND KD.PoNo = POD.PONo  " & vbCrLf & _
                                  "                                                        AND KD.SupplierID = POD.SupplierID  " & vbCrLf & _
                                  "                                                        AND KD.PartNo = POD.PartNo  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID  " & vbCrLf

                ls_sql = ls_sql + "                                                        AND KD.KanbanNo = KM.KanbanNo  " & vbCrLf & _
                                  "                                                        AND KD.SupplierID = KM.SupplierID  " & vbCrLf & _
                                  "                                                        AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
                                  "                                                             AND KD.KanbanNo = SDD.KanbanNo  " & vbCrLf & _
                                  "                                                             AND KD.SupplierID = SDD.SupplierID  " & vbCrLf & _
                                  "                                                             AND KD.PartNo = SDD.PartNo  " & vbCrLf & _
                                  "                                                             AND KD.PoNo = SDD.PoNo  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
                                  "                                                             AND SDM.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
                                  "                                                             AND SDM.SupplierID = SDD.SupplierID  " & vbCrLf

                ls_sql = ls_sql + "                      LEFT JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                                  "                                                               AND SDD.KanbanNo = PRD.KanbanNo  " & vbCrLf & _
                                  "                                                               AND SDD.SupplierID = PRD.SupplierID  " & vbCrLf & _
                                  "                                                               AND SDD.PartNo = PRD.PartNo  " & vbCrLf & _
                                  "                                                               AND SDD.PONo = PRD.PONo  " & vbCrLf & _
                                  "                                                               AND PRD.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                                  "                                                              AND PRM.SuratJalanNo = PRD.SuratJalanNo  " & vbCrLf & _
                                  "                      LEFT JOIN ( SELECT  SuratJalanno ,  " & vbCrLf & _
                                  "                                          SupplierID ,  " & vbCrLf & _
                                  "                                          AffiliateID ,  " & vbCrLf

                ls_sql = ls_sql + "                                          PONO ,  " & vbCrLf & _
                                  "                                          KanbanNO ,  " & vbCrLf & _
                                  "                                          Partno ,  " & vbCrLf & _
                                  "                                          UnitCls ,  " & vbCrLf & _
                                  "                                          DoQty = SUM(ISNULL(DoQty, 0))  " & vbCrLf & _
                                  "                                  FROM    DOPasi_Detail  " & vbCrLf & _
                                  "                                  GROUP BY SuratJalanno ,  " & vbCrLf & _
                                  "                                          SupplierID ,  " & vbCrLf & _
                                  "                                          AffiliateID ,  " & vbCrLf & _
                                  "                                          PONO ,  " & vbCrLf & _
                                  "                                          KanbanNO ,  " & vbCrLf

                ls_sql = ls_sql + "                                          Partno ,  " & vbCrLf & _
                                  "                                          UnitCls  " & vbCrLf & _
                                  "                                ) PDD ON PRD.AffiliateID = PDD.AffiliateID  " & vbCrLf & _
                                  "                                         AND PRD.KanbanNo = PDD.KanbanNo  " & vbCrLf & _
                                  "                                         AND PRD.SupplierID = PDD.SupplierID  " & vbCrLf & _
                                  "                                         AND PRD.PartNo = PDD.PartNo  " & vbCrLf & _
                                  "                                         AND PRD.PoNo = PDD.PoNo     " & vbCrLf & _
                                  "                                                --AND PDD.SuratJalanNoSupplier = SDM.SuratJalanNo    " & vbCrLf & _
                                  "                      LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
                                  "                                                         AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf & _
                                  "                                                --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf

                ls_sql = ls_sql + "                      LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
                                  "                                                                AND PDD.KanbanNo = RAD.KanbanNo  " & vbCrLf & _
                                  "                                                                AND PDD.SupplierID = RAD.SupplierID  " & vbCrLf & _
                                  "                                                                AND PDD.PartNo = RAD.PartNo  " & vbCrLf & _
                                  "                                                                AND PDD.PoNo = RAD.PoNo  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
                                  "                                                                AND RAM.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
                                  "                                                         --AND RAM.SupplierID = RAD.SupplierID    " & vbCrLf & _
                                  "                      LEFT JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
                                  "                                                              AND RAD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
                                  "                                                              AND RAD.PartNo = IPD.PartNo  " & vbCrLf

                ls_sql = ls_sql + "                                                              AND RAD.PONo = IPD.PONo  " & vbCrLf & _
                                  "                                                              AND RAD.SuratJalanNo = IPD.SuratJalanNo  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
                                  "                                                              AND IPD.InvoiceNo = IPM.InvoiceNo  " & vbCrLf & _
                                  "                                                              AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf & _
                                  "  		LEFT JOIN InvoiceSupplier_Detail ISD ON ISD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                                  "  												AND ISD.SupplierID = PRD.SupplierID  " & vbCrLf & _
                                  "  												AND ISD.SuratJalanNo = PRD.SuratJalanNo  " & vbCrLf & _
                                  "  												AND ISD.PONo = PRD.POno  " & vbCrLf & _
                                  "  												AND ISD.PartNo = PRD.PartNo  " & vbCrLf & _
                                  "                                                 AND ISD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                                  "  		LEFT JOIN InvoiceSupplier_Master ISM ON ISM.InvoiceNo = ISD.InvoiceNo  " & vbCrLf

                ls_sql = ls_sql + "  												AND ISM.AffiliateID = ISD.AffiliateID  " & vbCrLf & _
                                  "  												AND ISM.SupplierID = ISD.SupplierID  " & vbCrLf & _
                                  "  												AND ISM.suratJalanno = ISD.SuratJalanNo  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID  " & vbCrLf & _
                                  "                      LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                                  "                      LEFT JOIN MS_CurrCls MCS ON MCS.CurrCls = ISD.InvCurrCls  " & vbCrLf & _
                                  "                      LEFT JOIN MS_CurrCls MC ON MC.CurrCls = IPD.InvCurrCls  " & vbCrLf & _
                                  "            WHERE     KD.KanbanQty <> 0  AND POM.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf

                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + "  " & vbCrLf & _
                                  "          ) Hdr   " & vbCrLf & _
                                  " ORDER BY CONVERT(numeric,CONVERT(char, ROW_NUMBER() OVER ( ORDER BY Period, AffiliateCode, SupplierCode, PONo, KanbanNo, PartNo)))           " & vbCrLf & _
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

                'PO PERIOD
                If chkPOPeriod.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(CHAR(8),POM.Period,112) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyyMM01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyyMM" & ls_End) & "' " & vbCrLf
                End If

                'PASI DELIVERY DATE
                If chkPASIDelDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(CHAR(8),PDM.DeliveryDate,112) BETWEEN '" & Format(dtPASIDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If

                'AFFILIATE RECEIVE DATE
                If chkAffiliateRecDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(CHAR(8),RAM.ReceiveDate,112) BETWEEN '" & Format(dtAffiliateRecDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtAffiliateRecDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If

                'PASI INVOICE DATE
                If chkPASIInvDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(CHAR(8),IPM.DeliveryDate,112) BETWEEN '" & Format(dtPASIInvDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIInvDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
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

                'PASI SJ NO
                If Trim(txtPASISJNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(PDM.SuratJalanNo,'') LIKE '%" & Trim(txtPASISJNo.Text) & "%'" & vbCrLf
                End If

                'PASI INV NO
                If Trim(txtPASIInvNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(IPM.InvoiceNo,'') LIKE '%" & Trim(txtPASIInvNo.Text) & "%'" & vbCrLf
                End If

                ls_sql = "  SELECT  ColNo = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY Period, AffiliateCode, SupplierCode, PONo, KanbanNo, PartNo)) ,    " & vbCrLf & _
                      "  		Period ,    " & vbCrLf & _
                      " 		PONo , " & vbCrLf & _
                      "  		SupplierCode ,    		 " & vbCrLf & _
                      "  		POKanban = CASE WHEN ISNULL(KanbanCls, '0') = '1' THEN 'YES'    " & vbCrLf & _
                      "  						ELSE 'NO'    " & vbCrLf & _
                      "  					END ,    " & vbCrLf & _
                      "  		KanbanNo ,  " & vbCrLf & _
                      " 		PartNo ,    " & vbCrLf & _
                      "  		PartName ,    " & vbCrLf & _
                      " 		QtyPO = QtyPO ,    "

                ls_sql = ls_sql + " 		PASIDeliveryDate ,    " & vbCrLf & _
                                  "  		PASISJNo ,    " & vbCrLf & _
                                  " 		PASIDeliveryQty ,    " & vbCrLf & _
                                  "  		AffiliateReceiveDate ,    		  				 " & vbCrLf & _
                                  "  		AffiliateReceivingQty ,     " & vbCrLf & _
                                  "  		InvoiceToAffiliateQty ,      " & vbCrLf & _
                                  "  		InvoiceNoToAffiliate ,    " & vbCrLf & _
                                  "  		InvoiceDateToAffiliate ,    		 " & vbCrLf & _
                                  " 		InvoiceToAffiliateCurr ,    " & vbCrLf & _
                                  "  		InvoiceToAffiliateAmount " & vbCrLf & _
                                  "  FROM    ( SELECT DISTINCT    "

                ls_sql = ls_sql + "  			Period = SUBSTRING(CONVERT(CHAR, POM.Period, 106), 4, 9) ,    " & vbCrLf & _
                                  "  			PONo = POM.PONo ,    " & vbCrLf & _
                                  "  			AffiliateCode = POM.AffiliateID ,    " & vbCrLf & _
                                  "  			AffiliateName = MA.AffiliateName ,    " & vbCrLf & _
                                  "  			SupplierCode = POM.SupplierID ,    " & vbCrLf & _
                                  "  			SupplierName = MS.SupplierName ,    " & vbCrLf & _
                                  "  			KanbanCls = ISNULL(POD.KanbanCls, '0') ,    			KanbanNo = ISNULL(KD.KanbanNo,''),      " & vbCrLf & _
                                  "  			SupplierPlanDeliveryDate = '',      " & vbCrLf & _
                                  "  			SupplierDeliveryDate = '',      " & vbCrLf & _
                                  "  			SupplierSJNO = '',      " & vbCrLf & _
                                  "  			PASIReceiveDate = '',  "

                ls_sql = ls_sql + "              PASIDeliveryDate = ISNULL(CONVERT(CHAR,PDM.DeliveryDate,106),'') ,  " & vbCrLf & _
                                  "              PASISJNo = ISNULL(PDM.SuratJalanNo,'') ,    " & vbCrLf & _
                                  "              AffiliateReceiveDate = ISNULL(CONVERT(CHAR,RAM.ReceiveDate,106),'') ,  " & vbCrLf & _
                                  "              PartNo = POD.PartNo,      " & vbCrLf & _
                                  "  			PartName = MP.PartName,      " & vbCrLf & _
                                  "  			UOM = UC.Description,     			SupplierDeliveryQty = 0,      " & vbCrLf & _
                                  "  			PASIReceivingQty = 0,     " & vbCrLf & _
                                  "  			PASIDeliveryQty = ISNULL(PDD.DOQty,0),      " & vbCrLf & _
                                  "  			AffiliateReceivingQty = ISNULL(RAD.RecQty,0),      " & vbCrLf & _
                                  "  			InvoiceFromSupplierQty = 0,       " & vbCrLf & _
                                  "  			InvoiceToAffiliateQty = ISNULL(IPD.DOQty,0),      "

                ls_sql = ls_sql + "  			InvoiceNoFromSupplier = '',     " & vbCrLf & _
                                  "  			InvoiceDateFromSupplier = '',     " & vbCrLf & _
                                  "  			InvoiceFromSupplierCurr = '',     " & vbCrLf & _
                                  "  			InvoiceFromSupplierAmount = '',     " & vbCrLf & _
                                  "  			InvoiceNoToAffiliate = ISNULL(IPM.InvoiceNo,''),      			InvoiceDateToAffiliate = ISNULL(CONVERT(CHAR,IPM.DeliveryDate,106),''),     " & vbCrLf & _
                                  "  			InvoiceToAffiliateCurr = 'IDR',     " & vbCrLf & _
                                  "  			InvoiceToAffiliateAmount = ISNULL(IPD.DOQty,0) * ISNULL(PDD.Price,0),      " & vbCrLf & _
                                  "              PODelivery = 0 ,    " & vbCrLf & _
                                  "              RemainingQtyPOPASI = 0 ,    " & vbCrLf & _
                                  "              RemainingQtyPOSupplier = 0,   " & vbCrLf & _
                                  "              QtyPO = ISNULL(POD.POQty,0)  ,    "

                ls_sql = ls_sql + "              h_affiliateorder = KM.AffiliateID ,    " & vbCrLf & _
                                  "              H_PasiSJ = PDD.SuratJalanNo ,    " & vbCrLf & _
                                  "              h_poorder = KD.PONo ,    " & vbCrLf & _
                                  "              h_idxorder = '0' ,                h_kanbanorder = KM.KanbanNo,h_partno ='0'    " & vbCrLf & _
                                  "              FROM      dbo.PO_Master POM    " & vbCrLf & _
                                  "                        LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                                  "                                                   AND POM.PONo = POD.PONo    " & vbCrLf & _
                                  "                                                   AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
                                  "                        LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                                  "                                                          AND KD.PoNo = POD.PONo    " & vbCrLf & _
                                  "                                                          AND KD.SupplierID = POD.SupplierID    "

                ls_sql = ls_sql + "                                                          AND KD.PartNo = POD.PartNo    " & vbCrLf & _
                                  "                        LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf & _
                                  "                                                          AND KD.KanbanNo = KM.KanbanNo                                                            AND KD.SupplierID = KM.SupplierID    " & vbCrLf & _
                                  "                                                          AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
                                  "                        LEFT JOIN DOPasi_Detail PDD ON KD.AffiliateID = PDD.AffiliateID    " & vbCrLf & _
                                  "                                           AND KD.KanbanNo = PDD.KanbanNo    " & vbCrLf & _
                                  "                                           AND KD.SupplierID = PDD.SupplierID    " & vbCrLf & _
                                  "                                           AND KD.PartNo = PDD.PartNo    " & vbCrLf & _
                                  "                                           AND KD.PoNo = PDD.PoNo       " & vbCrLf & _
                                  "                        LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf & _
                                  "                                                           AND PDD.SuratJalanNo = PDM.SuratJalanNo          "

                ls_sql = ls_sql + "                        LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
                                  "                                                                  AND PDD.KanbanNo = RAD.KanbanNo                                                                    AND PDD.SupplierID = RAD.SupplierID    " & vbCrLf & _
                                  "                                                                  AND PDD.PartNo = RAD.PartNo    " & vbCrLf & _
                                  "                                                                  AND PDD.PoNo = RAD.PONo    " & vbCrLf & _
                                  "  																AND PDD.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
                                  "                        LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo    " & vbCrLf & _
                                  "                                                                  AND RAM.AffiliateID = RAD.AffiliateID         " & vbCrLf & _
                                  "                        LEFT JOIN dbo.PLPASI_Detail IPD ON PDD.AffiliateID = IPD.AffiliateID    " & vbCrLf & _
                                  "                                                                AND PDD.KanbanNo = IPD.KanbanNo    " & vbCrLf & _
                                  "                                                                AND PDD.PartNo = IPD.PartNo    " & vbCrLf & _
                                  "                                                                AND PDD.PONo = IPD.PONo    "

                ls_sql = ls_sql + "                                                                AND PDD.SuratJalanNo = IPD.SuratJalanNo                          LEFT JOIN dbo.PLPASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID                                                                  " & vbCrLf & _
                                  "                                                                AND IPD.SuratJalanNo = IPM.SuratJalanNo    		  " & vbCrLf & _
                                  "                        LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                                  "                        LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls    " & vbCrLf & _
                                  "                        LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                                  "                        LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID    " & vbCrLf & _
                                  "                        LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode                          " & vbCrLf & _
                                  "                        LEFT JOIN MS_Price MSP ON MSP.AffiliateID = IPD.AffiliateID and MSP.PartNo = IPD.PartNo and (IPM.DeliveryDate between MSP.StartDate and MSP.EndDate)                         " & vbCrLf & _
                                  "              WHERE     KD.KanbanQty <> 0  AND POM.AffiliateID = '" & Session("AffiliateID") & "'  " & vbCrLf

                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + "  " & vbCrLf & _
                                  "          ) Hdr   " & vbCrLf & _
                                  " ORDER BY CONVERT(numeric,CONVERT(char, ROW_NUMBER() OVER ( ORDER BY Period, AffiliateCode, SupplierCode, PONo, KanbanNo, PartNo)))           " & vbCrLf & _
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
                        Call EpPlusExportExcel(FilePath, "Sheet1", dtProd, "A:8", psERR)
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("G01Msg") = lblInfo.Text
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowPager)

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