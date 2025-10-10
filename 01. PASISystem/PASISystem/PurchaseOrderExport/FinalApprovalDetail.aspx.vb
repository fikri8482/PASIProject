Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO

Public Class FinalApprovalDetail
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = ""

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "O01"
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim dtExcel As DataTable
    Dim pPeriod As Date

#End Region

#Region "CONTROL EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim param As String = ""

        lblInfo.Text = ""

        Try
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                param = Request.QueryString("prm").ToString
                If param <> "" Then
                    pPeriod = Split(param, "|")(0)
                    txtperiod.Text = Format(pPeriod, "MMM yyyy")
                    txtoriginal.Text = Split(param, "|")(1)
                    txtorderno.Text = Split(param, "|")(2)
                    txtaffiliate.Text = Split(param, "|")(3)
                    txtaffname.Text = Split(param, "|")(4)
                    txtsupplier.Text = Split(param, "|")(5)
                    txtsuppname.Text = Split(param, "|")(6)

                    Call up_GridLoad()
                End If
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("G01Msg")
        Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If
                
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub up_GridLoad()
        Dim ls_sql As String = ""
        Dim ls_Filter As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_sql = " SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY AllowAccess DESC, PartNo, AffiliateID, SupplierID) NoUrut, *  " & vbCrLf & _
                  " FROM( " & vbCrLf & _
                  " Select DISTINCT  " & vbCrLf & _
                  " '1' AllowAccess,  " & vbCrLf & _
                  " '1' AdaData,  " & vbCrLf & _
                  " RTRIM(B.PartNo)PartNo,  " & vbCrLf & _
                  " RTRIM(C.PartName)PartName,  " & vbCrLf & _
                  " box1 = isnull(PL1.BoxNo,''), " & vbCrLf & _
                  " box2 = isnull(PL2.BoxNo,''), " & vbCrLf & _
                  " RTRIM(ISNULL(d.Description, UPO.UOM))UOM,  " & vbCrLf & _
                  " MOQ = CONVERT(NUMERIC(18,0), ISNULL(b.POMOQ,MPM.MOQ)),  "

            ls_sql = ls_sql + " QtyBox = CONVERT(NUMERIC(18,0), ISNULL(b.POQtyBox,MPM.QtyBox)),  " & vbCrLf & _
                              " Week1 = CONVERT(NUMERIC(18,0), B.Week1),  " & vbCrLf & _
                              " B.Week2,  " & vbCrLf & _
                              " B.Week3,  " & vbCrLf & _
                              " B.Week4,  " & vbCrLf & _
                              " B.Week5,  " & vbCrLf & _
                              " TotalPOQty = CONVERT(NUMERIC(18,0), B.Week1),  " & vbCrLf & _
                              " PreviousForecast = ISNULL(PrevQty.Forecast1,0),  " & vbCrLf & _
                              " B.Forecast1,  " & vbCrLf & _
                              " B.Forecast2,  " & vbCrLf & _
                              " B.Forecast3,  "

            ls_sql = ls_sql + " Variance = CASE WHEN ISNULL(PrevQty.Forecast1,0) = 0 THEN 0 ELSE B.Week1 - PrevQty.Forecast1 END,  " & vbCrLf & _
                              " VariancePercentage = CASE WHEN ISNULL(PrevQty.Forecast1,0) = 0 THEN 0 ELSE ((B.Week1 - PrevQty.Forecast1) / PrevQty.Forecast1) * 100 END,  " & vbCrLf & _
                              " a.PONo,  " & vbCrLf & _
                              " a.ShipCls,  " & vbCrLf & _
                              " a.CommercialCls,  " & vbCrLf & _
                              " a.ForwarderID,  " & vbCrLf & _
                              " a.Period,  " & vbCrLf & _
                              " RTRIM(a.AffiliateID)AffiliateID,  " & vbCrLf & _
                              " RTRIM(a.SupplierID)SupplierID,  " & vbCrLf & _
                              " ErrorStatus = ISNULL(UPO.errorCls,'')  " & vbCrLf & _
                              " FROM PO_Master_Export a  "

            ls_sql = ls_sql + " INNER JOIN PO_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = B.AffiliateID AND a.SupplierID = B.SupplierID AND a.OrderNo1 = b.OrderNo1  " & vbCrLf & _
                              " LEFT JOIN (  " & vbCrLf & _
                              "    SELECT Forecast1, PartNo, a.AffiliateID, a.PONo, a.OrderNo1 FROM PO_Detail_Export a  " & vbCrLf & _
                              "    INNER JOIN PO_Master_Export b ON a.PONo = b.PONo and a.OrderNo1 = b.OrderNo1 and a.AffiliateID = b.AffiliateID  and a.SupplierID = b.SupplierID  " & vbCrLf & _
                              "    WHERE Period = '" & DateAdd(DateInterval.Month, -1, pPeriod) & "' and a.PONo = a.PONo and b.EmergencyCls <> 'E' and a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              " )PrevQty ON PrevQty.PartNo = b.PartNo and PrevQty.AffiliateID = b.AffiliateID --and PrevQty.PONo = b.PONo and PrevQty.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              " LEFT JOIN MS_Parts c ON c.PartNo = B.PartNo  " & vbCrLf & _
                              " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = b.PartNo AND MPM.AffiliateID = b.AffiliateID AND MPM.SupplierID= b.SupplierID  " & vbCrLf & _
                              " LEFT JOIN MS_UnitCls d ON d.UnitCls = c.UnitCls  " & vbCrLf & _
                              " LEFT JOIN UploadPOExport UPO ON UPO.PONo = a.Pono AND a.AffiliateID = UPO.AffiliateID AND UPO.SupplierID = a.supplierID AND UPO.ForwarderID = a.ForwarderID AND UPO.Partno = b.PartNo  " & vbCrLf & _
                              " LEFT JOIN (select boxNo = Min(LabelNo), PoNo, AffiliateID, SupplierID, PartNo from PrintLabelExport group by PoNo, AffiliateID, SupplierID, PartNo) PL1 ON PL1.PONo = a.POno and PL1.AffiliateID = a.AffiliateID and PL1.supplierID = a.supplierID and PL1.PartNo = b.partNo "

            ls_sql = ls_sql + " LEFT JOIN (select boxNo = Max(LabelNo), PoNo, AffiliateID, SupplierID, PartNo from PrintLabelExport group by PoNo, AffiliateID, SupplierID, PartNo) PL2 ON PL2.PONo = a.POno and PL2.AffiliateID = a.AffiliateID and PL2.supplierID = a.supplierID and PL2.PartNo = b.partNo " & vbCrLf & _
                              " WHERE a.AffiliateID = '" & Trim(txtaffiliate.Text) & "' " & vbCrLf & _
                              " and a.SupplierID = '" & Trim(txtsupplier.Text) & "' " & vbCrLf & _
                              " and a.POno = '" & Trim(txtoriginal.Text) & "' " & vbCrLf & _
                              " and a.OrderNo1 = '" & Trim(txtorderno.Text) & "' " & vbCrLf & _
                              " )x "


            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            dtExcel = ds.Tables(0)
            sqlConn.Close()

        End Using
    End Sub

#End Region

#Region "Excel"

    'Private Sub GetExcel()
    '    Call up_GridLoad()
    '    FileName = "REMAINING REPORT.xlsx"
    '    FilePath = Server.MapPath("~\Template\" & FileName)
    '    If grid.VisibleRowCount - 1 > 0 Then
    '        Call epplusExportHeaderExcel(FilePath, "", dtExcel, "D:3", "")
    '    Else
    '        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
    '        grid.JSProperties("cpMessage") = lblInfo.Text
    '    End If
    'End Sub

    'Private Sub EpPlusDrawAllBorders(ByVal Rg As ExcelRange)
    '    With Rg
    '        .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
    '        .Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
    '        .Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
    '        .Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
    '        .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
    '    End With
    'End Sub

    'Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
    '    With Rg
    '        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '    End With
    'End Sub

    'Private Sub epplusExportHeaderExcel(ByVal pFilename As String, ByVal pSheetName As String,
    '                          ByVal pData1 As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

    '    Try

    '        Dim NewFileName As String = Server.MapPath("~\ProgressReportExport\Remaining Report.xlsx")
    '        If (System.IO.File.Exists(pFilename)) Then
    '            System.IO.File.Copy(pFilename, NewFileName, True)
    '        End If

    '        Dim rowstart As String = Split(pCellStart, ":")(1)
    '        Dim Coltart As String = Split(pCellStart, ":")(0)
    '        Dim fi As New FileInfo(NewFileName)

    '        Dim exl As New ExcelPackage(fi)
    '        Dim ws As ExcelWorksheet

    '        ws = exl.Workbook.Worksheets("REMAINING")
    '        Dim irow As Long = 0
    '        Dim iRowTmp As Long = 0
    '        Dim icol As Long = 0

    '        With ws
    '            ws.Cells("K4").Value = Format(dtPeriodFrom.Value, "MMM yyyy") + "-" + Format(dtPeriodTo.Value, "MMM yyyy")
    '            ws.Cells("K4").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '            If rdrRAll.Value = 1 Then ws.Cells("K5").Value = "ALL"
    '            If rdrRYes.Value = 1 Then ws.Cells("K5").Value = "YES"
    '            If rdrRNo.Value = 1 Then ws.Cells("K5").Value = "NO"
    '            ws.Cells("K5").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '            ws.Cells("K6").Value = IIf(Trim(cboPart.Text) = "==ALL==", "ALL", Trim(cboPart.Text) + "-" + Trim(txtPartName.Text))
    '            ws.Cells("K6").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '            ws.Cells("K7").Value = IIf(Trim(txtboxno.Text) = "", "-", Trim(txtboxno.Text))
    '            ws.Cells("K7").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left

    '            ws.Cells("AA4").Value = IIf(Trim(cbosupplier.Text) = "==ALL==", "ALL", Trim(cbosupplier.Text) + "-" + Trim(txtsupplier.Text))
    '            ws.Cells("AA4").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '            ws.Cells("AA5").Value = IIf(Trim(cboAffiliate.Text) = "==ALL==", "ALL", Trim(cboAffiliate.Text) + "'-" + Trim(txtAffiliate.Text))
    '            ws.Cells("AA5").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '            If rdrEAll.Value = 1 Then ws.Cells("AA6").Value = "ALL"
    '            If rdrEyes.Value = 1 Then ws.Cells("AA6").Value = "YES"
    '            If rdrENo.Value = 1 Then ws.Cells("AA6").Value = "NO"
    '            ws.Cells("AA6").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '            ws.Cells("AA7").Value = IIf(Trim(txtpono.Text) = "", "-", Trim(txtpono.Text))
    '            ws.Cells("AA7").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '        End With

    '        iRowTmp = 11
    '        For irow = 0 To pData1.Rows.Count - 1
    '            If pData1.Rows.Count > 0 Then
    '                ws.Cells("B" & iRowTmp).Value = irow + 1
    '                ws.Cells("B" & iRowTmp & ":" & "C" & iRowTmp).Merge = True
    '                ws.Cells("D" & iRowTmp).Value = pData1.Rows(irow)("period")

    '                ws.Cells("D" & iRowTmp & ":" & "F" & iRowTmp).Merge = True
    '                ws.Cells("G" & iRowTmp).Value = pData1.Rows(irow)("AffiliateID")

    '                ws.Cells("K" & iRowTmp).Value = pData1.Rows(irow)("AffiliateName")
    '                ws.Cells("K" & iRowTmp & ":" & "S" & iRowTmp).Merge = True

    '                ws.Cells("T" & iRowTmp).Value = pData1.Rows(irow)("Orderno")
    '                ws.Cells("T" & iRowTmp & ":" & "Y" & iRowTmp).Merge = True

    '                ws.Cells("Z" & iRowTmp).Value = pData1.Rows(irow)("EmergencyCls")
    '                ws.Cells("Z" & iRowTmp & ":" & "AC" & iRowTmp).Merge = True
    '                ws.Cells("Z" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center

    '                ws.Cells("AD" & iRowTmp).Value = pData1.Rows(irow)("SupplierID")
    '                ws.Cells("AD" & iRowTmp & ":" & "AG" & iRowTmp).Merge = True

    '                ws.Cells("AH" & iRowTmp).Value = pData1.Rows(irow)("SupplierName")
    '                ws.Cells("AH" & iRowTmp & ":" & "AO" & iRowTmp).Merge = True

    '                ws.Cells("AP" & iRowTmp).Value = pData1.Rows(irow)("ETDVendor")
    '                ws.Cells("AP" & iRowTmp & ":" & "AS" & iRowTmp).Merge = True
    '                ws.Cells("AP" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center

    '                ws.Cells("AT" & iRowTmp).Value = pData1.Rows(irow)("ETDPORT")
    '                ws.Cells("AT" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
    '                ws.Cells("AT" & iRowTmp & ":" & "AW" & iRowTmp).Merge = True

    '                ws.Cells("AX" & iRowTmp).Value = pData1.Rows(irow)("ETAPORT")
    '                ws.Cells("AX" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
    '                ws.Cells("AX" & iRowTmp & ":" & "BA" & iRowTmp).Merge = True

    '                ws.Cells("BB" & iRowTmp).Value = pData1.Rows(irow)("ETAFACTORY")
    '                ws.Cells("BB" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
    '                ws.Cells("BB" & iRowTmp & ":" & "BE" & iRowTmp).Merge = True

    '                ws.Cells("BF" & iRowTmp).Value = pData1.Rows(irow)("PartNo")
    '                ws.Cells("BF" & iRowTmp & ":" & "BJ" & iRowTmp).Merge = True

    '                ws.Cells("BK" & iRowTmp).Value = pData1.Rows(irow)("PartName")
    '                ws.Cells("BK" & iRowTmp & ":" & "BR" & iRowTmp).Merge = True

    '                ws.Cells("BS" & iRowTmp).Value = Trim(pData1.Rows(irow)("UOm"))
    '                ws.Cells("BS" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
    '                ws.Cells("BS" & iRowTmp & ":" & "BT" & iRowTmp).Merge = True

    '                ws.Cells("BU" & iRowTmp).Value = pData1.Rows(irow)("QtyBox")
    '                ws.Cells("BU" & iRowTmp & ":" & "BW" & iRowTmp).Merge = True

    '                ws.Cells("BX" & iRowTmp).Value = pData1.Rows(irow)("BoxNo")
    '                ws.Cells("BX" & iRowTmp & ":" & "CD" & iRowTmp).Merge = True

    '                ws.Cells("CE" & iRowTmp).Value = pData1.Rows(irow)("POQty")
    '                ws.Cells("CE" & iRowTmp & ":" & "CH" & iRowTmp).Merge = True
    '                ws.Cells("CE" & iRowTmp).Style.Numberformat.Format = "###,##0"
    '                ws.Cells("CE" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

    '                ws.Cells("CI" & iRowTmp).Value = pData1.Rows(irow)("DOQty")
    '                ws.Cells("CI" & iRowTmp & ":" & "CM" & iRowTmp).Merge = True
    '                ws.Cells("Ci" & iRowTmp).Style.Numberformat.Format = "###,##0"
    '                ws.Cells("CI" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

    '                ws.Cells("CN" & iRowTmp).Value = pData1.Rows(irow)("GoodRecQty")
    '                ws.Cells("CN" & iRowTmp & ":" & "CR" & iRowTmp).Merge = True
    '                ws.Cells("CN" & iRowTmp).Style.Numberformat.Format = "###,##0"
    '                ws.Cells("CN" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

    '                ws.Cells("CS" & iRowTmp).Value = pData1.Rows(irow)("DefectRecQty")
    '                ws.Cells("CS" & iRowTmp & ":" & "CW" & iRowTmp).Merge = True
    '                ws.Cells("CS" & iRowTmp).Style.Numberformat.Format = "###,##0"
    '                ws.Cells("CS" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

    '                ws.Cells("CX" & iRowTmp).Value = pData1.Rows(irow)("Remaining")
    '                ws.Cells("CX" & iRowTmp & ":" & "DB" & iRowTmp).Merge = True
    '                ws.Cells("CX" & iRowTmp).Style.Numberformat.Format = "###,##0"
    '                ws.Cells("CX" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

    '                'ws.Cells("C18" & ":J" & iRowTmp).Style.Numberformat.Format = "#,###"
    '            End If
    '            iRowTmp = iRowTmp + 1
    '        Next


    '        Dim rgAll As ExcelRange = ws.Cells(11, 2, iRowTmp - 1, 106)
    '        EpPlusDrawAllBorders(rgAll)

    '        exl.Save()

    '        DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

    '        exl = Nothing
    '    Catch ex As Exception
    '        pErr = ex.Message
    '    End Try

    'End Sub

#End Region

End Class