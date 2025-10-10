Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports DevExpress.Web.ASPxMenu
Imports OfficeOpenXml
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.Net

Public Class PrintPreBookingList
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_KanbanDate As String
    Dim ls_approve As Boolean

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim dtHeader As DataTable
    Dim dtHeader2 As DataTable
    Dim dtDetail As DataTable
#End Region

#Region "CONTROL EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If

            Session("M01Url") = Request.QueryString("Session")
            Session("MenuDesc") = "PRINT PRE-BOOKING VESSEL"

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Clear()
                dtPeriodFrom.Text = Format(Now, "yyyy-MM")
                Call up_fillcombo()
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
        End Try
    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'Affiliate Code
        ls_sql = "SELECT [Affiliate Code] = RTRIM(AffiliateID) ,[Affiliate Name] = RTRIM(Affiliatename) FROM MS_Affiliate  where isnull(overseascls, '0') = '1'" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Affiliate Code")
                .Columns(0).Width = 90
                .Columns.Add("Affiliate Name")
                .Columns(1).Width = 240

                .TextField = "Affiliate Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub Clear()
        lblerrmessage.Text = ""
    End Sub

#End Region

#Region "EXCEL"
    Private Sub GetExcel()
        Call GridLoadExcel_Header()
        Call GridLoadExcel_Detail()
        FileName = "PRE-BOOKING.xlsx"
        FilePath = Server.MapPath("~\Template\" & FileName)
        If dtHeader.Rows.Count > 0 Then
            Call epplusExportHeaderExcel(FilePath, "", dtHeader, dtDetail, "C:7", "")
        Else
            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
            Approve.JSProperties("cpMessage") = lblerrmessage.Text
        End If
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

    Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
        With Rg
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
    End Sub

    Private Sub epplusExportHeaderExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData1 As DataTable, ByVal pData2 As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try

            Dim NewFileName As String = Server.MapPath("~\PurchaseOrderExport\PreBooking.xlsx")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet
            Dim sumQtyBox1 As Long = 0
            Dim sumPalletBox1 As Single = 0
            Dim sumQtyBox2 As Long = 0
            Dim sumPalletBox2 As Single = 0
            Dim sumQtyBox3 As Long = 0
            Dim sumPalletBox3 As Single = 0
            Dim sumQtyBox4 As Long = 0
            Dim sumPalletBox4 As Single = 0
            Dim sumQtyBox5 As Long = 0
            Dim sumPalletBox5 As Single = 0
            Dim sumCbm1 As Long = 0
            Dim sumCbm2 As Long = 0
            Dim sumCbm3 As Long = 0
            Dim sumCbm4 As Long = 0
            Dim sumCbm5 As Long = 0

            ws = exl.Workbook.Worksheets("PRE-BOOKING")
            Dim irow As Long = 0
            Dim iRowTmp As Long = 0
            Dim icol As Long = 0

            With ws
                For irow = 0 To pData1.Rows.Count - 1
                    ws.Cells("A10").Value = "ETA PORT (" & pData1.Rows(irow)("DestinationPort") & ")"
                    ws.Cells("A12").Value = "VESSEL NAME 1"
                    ws.Cells("A13").Value = "VESSEL NAME 2"
                    If pData1.Rows.Count > 0 Then
                        If pData1.Rows(irow)("Week") = "1" Then
                            ws.Cells("C7").Value = Trim(pData1.Rows(irow)("OrderNo1"))
                            ws.Cells("C8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("C9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("C10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("C11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("C12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("C13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 3).Style.Font.Size = 11
                            .Cells(7, 3).Style.Font.Name = "Calibri"
                        ElseIf pData1.Rows(irow)("Week") = "2" Then
                            ws.Cells("D7").Value = Trim(pData1.Rows(irow)("OrderNo2"))
                            ws.Cells("D8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("D9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("D10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("D11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("D12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("D13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 4).Style.Font.Size = 11
                            .Cells(7, 4).Style.Font.Name = "Calibri"
                        ElseIf pData1.Rows(irow)("Week") = "3" Then
                            ws.Cells("E7").Value = Trim(pData1.Rows(irow)("OrderNo3"))
                            ws.Cells("E8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("E9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("E10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("E11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("E12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("E13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 5).Style.Font.Size = 11
                            .Cells(7, 5).Style.Font.Name = "Calibri"
                        ElseIf pData1.Rows(irow)("Week") = "4" Then
                            ws.Cells("F7").Value = Trim(pData1.Rows(irow)("OrderNo4"))
                            ws.Cells("F8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("F9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("F10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("F11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("F12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("F13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 6).Style.Font.Size = 11
                            .Cells(7, 6).Style.Font.Name = "Calibri"
                        ElseIf pData1.Rows(irow)("Week") = "5" Then
                            ws.Cells("F7").Value = Trim(pData1.Rows(irow)("OrderNo5"))
                            ws.Cells("F8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("F9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("F10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("F11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("F12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("F13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 6).Style.Font.Size = 11
                            .Cells(7, 6).Style.Font.Name = "Calibri"
                        End If
                    End If
                Next
            End With

            iRowTmp = 18
            For irow = 0 To pData2.Rows.Count - 1
                If pData2.Rows.Count > 0 Then
                    ws.Cells("B" & iRowTmp).Value = Trim(pData2.Rows(irow)("SupplierID"))
                    ws.Cells("C" & iRowTmp).Value = pData2.Rows(irow)("QtyBox1")
                    ws.Cells("D" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet1")
                    ws.Cells("E" & iRowTmp).Value = pData2.Rows(irow)("QtyBox2")
                    ws.Cells("F" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet2")
                    ws.Cells("G" & iRowTmp).Value = pData2.Rows(irow)("QtyBox3")
                    ws.Cells("H" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet3")
                    ws.Cells("I" & iRowTmp).Value = pData2.Rows(irow)("QtyBox4")
                    ws.Cells("J" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet4")
                    ws.Cells("K" & iRowTmp).Value = pData2.Rows(irow)("QtyBox5")
                    ws.Cells("L" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet5")
                    'ws.Cells("C18" & ":J" & iRowTmp).Style.Numberformat.Format = "#,###"

                    sumQtyBox1 = sumQtyBox1 + pData2.Rows(irow)("QtyBox1")
                    sumPalletBox1 = CDbl(sumPalletBox1) + CDbl(pData2.Rows(irow)("QtyPallet1"))
                    sumQtyBox2 = sumQtyBox2 + pData2.Rows(irow)("QtyBox2")
                    sumPalletBox2 = CDbl(sumPalletBox2) + CDbl(pData2.Rows(irow)("QtyPallet2"))
                    sumQtyBox3 = sumQtyBox3 + pData2.Rows(irow)("QtyBox3")
                    sumPalletBox3 = CDbl(sumPalletBox3) + CDbl(pData2.Rows(irow)("QtyPallet3"))
                    sumQtyBox4 = sumQtyBox4 + pData2.Rows(irow)("QtyBox4")
                    sumPalletBox4 = CDbl(sumPalletBox4) + CDbl(pData2.Rows(irow)("QtyPallet4"))
                    sumQtyBox5 = sumQtyBox5 + pData2.Rows(irow)("QtyBox5")
                    sumPalletBox5 = CDbl(sumPalletBox5) + CDbl(pData2.Rows(irow)("QtyPallet5"))

                    sumCbm1 = sumCbm1 + pData2.Rows(irow)("CBM1")
                    sumCbm2 = sumCbm2 + pData2.Rows(irow)("CBM2")
                    sumCbm3 = sumCbm3 + pData2.Rows(irow)("CBM3")
                    sumCbm4 = sumCbm4 + pData2.Rows(irow)("CBM4")
                    sumCbm5 = sumCbm5 + pData2.Rows(irow)("CBM5")

                    'ALIGNMENT
                    ws.Cells("B" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("C" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("D" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("E" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("F" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("G" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("H" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("I" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("J" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("K" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("L" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'FORMAT
                    ws.Cells("C" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("D" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("E" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("F" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("G" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("H" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("I" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("J" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("K" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("L" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                End If
                iRowTmp = iRowTmp + 1
            Next

            iRowTmp = iRowTmp
            ws.Cells("B" & iRowTmp).Value = "TOTAL CBM"
            ws.Cells("D" & iRowTmp).Value = sumCbm1
            ws.Cells("D" & iRowTmp).Style.Numberformat.Format = "###,##0"
            ws.Cells("F" & iRowTmp).Value = sumCbm2
            ws.Cells("F" & iRowTmp).Style.Numberformat.Format = "###,##0"
            ws.Cells("H" & iRowTmp).Value = sumCbm3
            ws.Cells("H" & iRowTmp).Style.Numberformat.Format = "###,##0"
            ws.Cells("J" & iRowTmp).Value = sumCbm4
            ws.Cells("J" & iRowTmp).Style.Numberformat.Format = "###,##0"
            ws.Cells("L" & iRowTmp).Value = sumCbm5
            ws.Cells("L" & iRowTmp).Style.Numberformat.Format = "###,##0"

            ws.Cells("B" & iRowTmp + 1).Value = "TOTAL PALLET"
            ws.Cells("D" & iRowTmp + 1).Value = sumPalletBox1 'Format(sumPalletBox1, "###,##0.00")
            ws.Cells("D" & iRowTmp + 1).Style.Numberformat.Format = "###,##0.00"
            ws.Cells("F" & iRowTmp + 1).Value = sumPalletBox2 'Format(sumPalletBox2, "###,##0.00")
            ws.Cells("F" & iRowTmp + 1).Style.Numberformat.Format = "###,##0.00"
            ws.Cells("H" & iRowTmp + 1).Value = sumPalletBox3 'Format(sumPalletBox3, "###,##0.00")
            ws.Cells("H" & iRowTmp + 1).Style.Numberformat.Format = "###,##0.00"
            ws.Cells("J" & iRowTmp + 1).Value = sumPalletBox4 'Format(sumPalletBox4, "###,##0.00")
            ws.Cells("J" & iRowTmp + 1).Style.Numberformat.Format = "###,##0.00"
            ws.Cells("L" & iRowTmp + 1).Value = sumPalletBox5 'Format(sumPalletBox4, "###,##0.00")
            ws.Cells("L" & iRowTmp + 1).Style.Numberformat.Format = "###,##0.00"

            ws.Cells("A" & iRowTmp + 2).Value = "CONTAINER"

            'rumus 40FT
            Dim ls_40FT As Single
            ls_40FT = 15
            Dim ls_40FTEx As Single
            ls_40FTEx = 7.6

            'rumus 20FT
            Dim ls_20FT As Single
            ls_20FT = 7.5

            Dim JmlPallet As Long = 0
            Dim JmlContainer40 As Long = 0
            Dim JmlContainer20 As Single = 0

            '=============WEEK 1=================='
            If CDbl(sumPalletBox1) >= ls_40FTEx Then
                'menggunakan 40FT
                JmlContainer40 = IIf((sumPalletBox1 / ls_40FT) < 1, 1, Math.Floor(sumPalletBox1 / ls_40FT))
                ws.Cells("D" & iRowTmp + 3).Value = JmlContainer40
                If CDbl(sumPalletBox1) >= ls_40FT Then
                    JmlPallet = CInt(sumPalletBox1 Mod ls_40FT)
                    If CDbl(JmlPallet) > 0 Then
                        JmlContainer20 = JmlPallet / ls_20FT
                        If JmlContainer20 > 0 And JmlContainer20 < 1 Then JmlContainer20 = 1
                        ws.Cells("D" & iRowTmp + 2).Value = JmlContainer20
                        ws.Cells("D" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
                    End If
                End If
            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox1 / ls_20FT)
                ws.Cells("D" & iRowTmp + 2).Value = JmlContainer20
            End If

            If CDbl(sumPalletBox1) > 0 Then
                If CDbl(sumPalletBox1) < 5 Then
                    ws.Cells("D" & iRowTmp + 4).Value = "LCL"
                Else
                    ws.Cells("D" & iRowTmp + 4).Value = "FCL"
                End If
            End If

            '=============WEEK 2=================='
            If CDbl(sumPalletBox2) >= ls_40FTEx Then
                'menggunakan 40FT
                JmlContainer40 = IIf((sumPalletBox2 / ls_40FT) < 1, 1, Math.Floor(sumPalletBox2 / ls_40FT))
                ws.Cells("F" & iRowTmp + 3).Value = JmlContainer40
                If CDbl(sumPalletBox2) >= ls_40FT Then
                    JmlPallet = CInt(sumPalletBox2 Mod ls_40FT)
                    If CDbl(JmlPallet) > 0 Then
                        JmlContainer20 = JmlPallet / ls_20FT
                        If JmlContainer20 > 0 And JmlContainer20 < 1 Then JmlContainer20 = 1
                        ws.Cells("F" & iRowTmp + 2).Value = JmlContainer20
                        ws.Cells("F" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
                    End If
                End If
            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox2 / ls_20FT)
                ws.Cells("F" & iRowTmp + 2).Value = JmlContainer20
            End If

            If CDbl(sumPalletBox2) > 0 Then
                If CDbl(sumPalletBox2) < 5 Then
                    ws.Cells("F" & iRowTmp + 4).Value = "LCL"
                Else
                    ws.Cells("F" & iRowTmp + 4).Value = "FCL"
                End If
            End If

            '=============WEEK 3=================='
            If CDbl(sumPalletBox3) >= ls_40FTEx Then
                'menggunakan 40FT
                JmlContainer40 = IIf((sumPalletBox3 / ls_40FT) < 1, 1, Math.Floor(sumPalletBox3 / ls_40FT))
                ws.Cells("H" & iRowTmp + 3).Value = JmlContainer40

                If CDbl(sumPalletBox3) >= ls_40FT Then
                    JmlPallet = CInt(sumPalletBox3 Mod ls_40FT)
                    If CDbl(JmlPallet) > 0 Then
                        JmlContainer20 = JmlPallet / ls_20FT
                        If JmlContainer20 > 0 And JmlContainer20 < 1 Then JmlContainer20 = 1
                        ws.Cells("H" & iRowTmp + 2).Value = JmlContainer20
                        ws.Cells("H" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
                    End If
                End If
            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox3 / ls_20FT)
                ws.Cells("H" & iRowTmp + 2).Value = JmlContainer20
                ws.Cells("H" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
            End If

            If CDbl(sumPalletBox3) > 0 Then
                If CDbl(sumPalletBox3) < 5 Then
                    ws.Cells("H" & iRowTmp + 4).Value = "LCL"
                Else
                    ws.Cells("H" & iRowTmp + 4).Value = "FCL"
                End If
            End If

            '=============WEEK 4=================='
            If CDbl(sumPalletBox4) >= ls_40FTEx Then
                'menggunakan 40FT
                JmlContainer40 = IIf((sumPalletBox4 / ls_40FT) < 1, 1, Math.Floor(sumPalletBox4 / ls_40FT))
                ws.Cells("J" & iRowTmp + 3).Value = JmlContainer40
                ws.Cells("J" & iRowTmp + 3).Style.Numberformat.Format = "#,##0"
                If CDbl(sumPalletBox4) >= ls_40FT Then
                    JmlPallet = CInt(sumPalletBox4 Mod ls_40FT)
                    If CDbl(JmlPallet) > 0 Then
                        JmlContainer20 = JmlPallet / ls_20FT
                        If JmlContainer20 > 0 And JmlContainer20 < 1 Then JmlContainer20 = 1
                        ws.Cells("J" & iRowTmp + 2).Value = JmlContainer20
                        ws.Cells("J" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
                    End If
                End If
            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox4 / ls_20FT)
                ws.Cells("J" & iRowTmp + 2).Value = JmlContainer20
                ws.Cells("J" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
            End If

            If CDbl(sumPalletBox4) > 0 Then
                If CDbl(sumPalletBox4) < 5 Then
                    ws.Cells("J" & iRowTmp + 4).Value = "LCL"
                Else
                    ws.Cells("J" & iRowTmp + 4).Value = "FCL"
                End If
            End If

            '=============WEEK 5=================='
            If CDbl(sumPalletBox5) >= ls_40FTEx Then
                'menggunakan 40FT
                JmlContainer40 = IIf((sumPalletBox5 / ls_40FT) < 1, 1, Math.Floor(sumPalletBox5 / ls_40FT))
                ws.Cells("L" & iRowTmp + 3).Value = JmlContainer40
                ws.Cells("L" & iRowTmp + 3).Style.Numberformat.Format = "#,##0"
                If CDbl(sumPalletBox5) >= ls_40FT Then
                    JmlPallet = CInt(sumPalletBox5 Mod ls_40FT)
                    If CDbl(JmlPallet) > 0 Then
                        JmlContainer20 = JmlPallet / ls_20FT
                        If JmlContainer20 > 0 And JmlContainer20 < 1 Then JmlContainer20 = 1
                        ws.Cells("L" & iRowTmp + 2).Value = JmlContainer20
                        ws.Cells("L" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
                    End If
                End If
            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox5 / ls_20FT)
                ws.Cells("L" & iRowTmp + 2).Value = JmlContainer20
                ws.Cells("L" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
            End If

            If CDbl(sumPalletBox5) > 0 Then
                If CDbl(sumPalletBox5) < 5 Then
                    ws.Cells("L" & iRowTmp + 4).Value = "LCL"
                Else
                    ws.Cells("L" & iRowTmp + 4).Value = "FCL"
                End If
            End If

            ws.Cells("B" & iRowTmp + 2).Value = "20FT"
            ws.Cells("B" & iRowTmp + 3).Value = "40FT"

            Dim rgAll As ExcelRange = ws.Cells(18, 2, iRowTmp + 2, 12)
            EpPlusDrawAllBorders(rgAll)

            Dim rgAll2 As ExcelRange = ws.Cells(iRowTmp + 2, 1, iRowTmp + 3, 12)
            EpPlusDrawAllBorders(rgAll2)

            Dim rgAll3 As ExcelRange = ws.Cells(iRowTmp + 4, 1, iRowTmp + 4, 12)
            EpPlusDrawAllBorders(rgAll3)

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
            Approve.JSProperties("cpMessage") = pErr
        End Try

    End Sub

    Private Sub GridLoadExcel_Header()

        Dim ds As New DataSet
        Dim ls_sql As String = ""

        ls_sql = " SELECT distinct  Vessel1 = '', Vessel2 = '', " & vbCrLf & _
                  " OrderNo1 = (SELECT (STUFF((SELECT distinct ', ' + RTrim(orderNo1) " & vbCrLf & _
                  " 			FROM dbo.PO_Master_Export a   " & vbCrLf & _
                  " 				LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                  " 							UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                  " 							UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                  " 							UNION ALL Select * from MS_ETD_Export where week = '4'  " & vbCrLf & _
                  "                             UNION ALL Select * from MS_ETD_Export where week = '5')b " & vbCrLf & _
                  " 				ON a.AffiliateID = b.AffiliateID   " & vbCrLf & _
                  " 				AND a.SupplierID = b.SupplierID   " & vbCrLf & _
                  " 				AND a.Period = b.Period   " & vbCrLf & _
                  " 				AND b.ETAForwarder = a.ETDVendor1 " & vbCrLf

        ls_sql = ls_sql + " 				WHERE a.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _
                          " 				and week = '1'  " & vbCrLf & _
                          "                 AND a.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 			FOR XML PATH('')), 1, 2, ''))),  " & vbCrLf & _
                          " OrderNo2 = (SELECT (STUFF((SELECT distinct ', ' + RTrim(orderNo1) " & vbCrLf & _
                          " 			FROM dbo.PO_Master_Export a   " & vbCrLf & _
                          " 				LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '4'  " & vbCrLf & _
                          "                             UNION ALL Select * from MS_ETD_Export where week = '5')b " & vbCrLf & _
                          " 				ON a.AffiliateID = b.AffiliateID   " & vbCrLf

        ls_sql = ls_sql + " 				AND a.SupplierID = b.SupplierID   " & vbCrLf & _
                          " 				AND a.Period = b.Period   " & vbCrLf & _
                          " 				AND b.ETAForwarder = a.ETDVendor1 " & vbCrLf & _
                          " 				WHERE a.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _                          
                          "                 AND a.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 				and week = '2'  " & vbCrLf & _
                          " 			FOR XML PATH('')), 1, 2, ''))),  " & vbCrLf & _
                          " OrderNo3 = (SELECT (STUFF((SELECT distinct ', ' + RTrim(orderNo1) " & vbCrLf & _
                          " 			FROM dbo.PO_Master_Export a   " & vbCrLf & _
                          " 				LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf

        ls_sql = ls_sql + " 							UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '4'  " & vbCrLf & _
                          "                             UNION ALL Select * from MS_ETD_Export where week = '5')b " & vbCrLf & _
                          " 				ON a.AffiliateID = b.AffiliateID   " & vbCrLf & _
                          " 				AND a.SupplierID = b.SupplierID   " & vbCrLf & _
                          " 				AND a.Period = b.Period   " & vbCrLf & _
                          " 				AND b.ETAForwarder = a.ETDVendor1 " & vbCrLf & _
                          " 				WHERE a.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
                          "                 AND a.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 				and week = '3'  " & vbCrLf & _
                          " 			FOR XML PATH('')), 1, 2, ''))),  " & vbCrLf & _
                          " OrderNo4 = (SELECT (STUFF((SELECT distinct ', ' + RTrim(orderNo1) " & vbCrLf

        ls_sql = ls_sql + " 			FROM dbo.PO_Master_Export a   " & vbCrLf & _
                          " 				LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '4'  " & vbCrLf & _
                          "                             UNION ALL Select * from MS_ETD_Export where week = '5')b " & vbCrLf & _
                          " 				ON a.AffiliateID = b.AffiliateID   " & vbCrLf & _
                          " 				AND a.SupplierID = b.SupplierID   " & vbCrLf & _
                          " 				AND a.Period = b.Period   " & vbCrLf & _
                          " 				AND b.ETAForwarder = a.ETDVendor1 " & vbCrLf & _
                          " 				WHERE a.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'    " & vbCrLf & _
                          "                 AND a.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf

        ls_sql = ls_sql + " 				and week = '4'  " & vbCrLf & _
                          " 			FOR XML PATH('')), 1, 2, ''))),  " & vbCrLf & _
                          " OrderNo5 = (SELECT (STUFF((SELECT distinct ', ' + RTrim(orderNo1) " & vbCrLf & _
                          " 			FROM dbo.PO_Master_Export a   " & vbCrLf & _
                          " 				LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                          " 							UNION ALL Select * from MS_ETD_Export where week = '4'  " & vbCrLf & _
                          "                             UNION ALL Select * from MS_ETD_Export where week = '5')b " & vbCrLf & _
                          " 				ON a.AffiliateID = b.AffiliateID   " & vbCrLf & _
                          " 				AND a.SupplierID = b.SupplierID   " & vbCrLf & _
                          " 				AND a.Period = b.Period   " & vbCrLf & _
                          " 				AND b.ETAForwarder = a.ETDVendor1 " & vbCrLf & _
                          " 				WHERE a.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'    " & vbCrLf & _
                          "                 AND a.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 				and week = '5'  " & vbCrLf & _
                          " 			FOR XML PATH('')), 1, 2, ''))),  " & vbCrLf & _
                          " MEE.Week,  " & vbCrLf & _
                          " ETAForwarder = MEE.ETAForwarder,  " & vbCrLf & _
                          " ETDPORT = MEE.ETDPORT,  " & vbCrLf & _
                          " ETAPORT = MEE.ETAPORT,  " & vbCrLf & _
                          " ETAFACTORY = MEE.ETAFACTORY,  " & vbCrLf & _
                          " ISNULL(MA.DestinationPort,'')DestinationPort  " & vbCrLf & _
                          " FROM dbo.PO_Master_Export ME   " & vbCrLf & _
                          " LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf

        ls_sql = ls_sql + " UNION ALL Select * from MS_ETD_Export where week = '4' UNION ALL Select * from MS_ETD_Export where week = '5') MEE  " & vbCrLf & _
                          " ON MEE.AffiliateID = ME.AffiliateID   " & vbCrLf & _
                          " AND MEE.SupplierID = ME.SupplierID   " & vbCrLf & _
                          " AND MEE.Period = ME.Period   " & vbCrLf & _
                          " AND MEE.ETAForwarder = ME.ETDVendor1   " & vbCrLf & _
                          " LEFT JOIN MS_Affiliate MA on MA.AffiliateID = ME.AffiliateID " & vbCrLf & _
                          " WHERE ME.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
                          " AND ME.AffiliateID = '" & Trim(cboaffiliate.Text) & "' and Week is not null and ShipCls = 'B'" & vbCrLf & _
                          " " & vbCrLf



        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using
        dtHeader = ds.Tables(0)
    End Sub

    Private Sub GridLoadExcel_Detail()

        Dim ds As New DataSet
        Dim ls_sql As String = ""

        ls_sql = " SELECT distinct ME.SupplierID, QtyBox1=isnull(W1.QtyBox,0),QtyPallet1=isnull(W1.QtyPallet,0), " & vbCrLf & _
                  " QtyBox2=isnull(W2.QtyBox,0),QtyPallet2=isnull(W2.QtyPallet,0), " & vbCrLf & _
                  " QtyBox3=isnull(W3.QtyBox,0),QtyPallet3=isnull(W3.QtyPallet,0), " & vbCrLf & _
                  " QtyBox4=isnull(W4.QtyBox,0),QtyPallet4=isnull(W4.QtyPallet,0), " & vbCrLf & _
                  " QtyBox5=isnull(W5.QtyBox,0),QtyPallet5=isnull(W5.QtyPallet,0), " & vbCrLf & _
                  " CBM1 = isnull(W1.CBM,0) , CBM2 = isnull(W2.CBM,0), CBM3 = isnull(W3.CBM,0), CBM4 = isnull(W4.CBM,0), CBM5 = isnull(W5.CBM,0) " & vbCrLf & _
                  " FROM dbo.PO_Master_Export ME  " & vbCrLf & _
                  " LEFT JOIN PO_Detail_Export DE  " & vbCrLf & _
                  " ON ME.PONo = DE.PONo  " & vbCrLf & _
                  " 	AND ME.OrderNo1 = DE.OrderNo1  " & vbCrLf & _
                  " 	AND ME.SupplierID = DE.SupplierID  " & vbCrLf & _
                  " 	AND ME.AffiliateID = DE.AffiliateID  "

        ls_sql = ls_sql + " LEFT JOIN(SELECT distinct ME.SupplierID,QtyBox = (SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))), QtyPallet = (SUM((Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))/BoxPallet)),week " & vbCrLf & _
                          " 			--,CBM = SUM(week1*(Length*width*height)) * (SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox)))/ 1000000 " & vbCrLf & _
                          "             ,CBM =  SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox)) * SUM((Length*width*height)/6000) " & vbCrLf & _
                          " 			FROM dbo.PO_Master_Export ME  " & vbCrLf & _
                          " 			LEFT JOIN PO_Detail_Export DE  " & vbCrLf & _
                          " 			ON ME.PONo = DE.PONo  " & vbCrLf & _
                          " 				AND ME.OrderNo1 = DE.OrderNo1  " & vbCrLf & _
                          " 				AND ME.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND ME.AffiliateID = DE.AffiliateID   " & vbCrLf & _
                          " 			LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '3'  "

        ls_sql = ls_sql + " 						UNION ALL Select * from MS_ETD_Export where week = '4' UNION ALL Select * from MS_ETD_Export where week = '5')MEE " & vbCrLf & _
                          " 			ON MEE.AffiliateID = ME.AffiliateID   " & vbCrLf & _
                          "    				AND MEE.SupplierID = ME.SupplierID   " & vbCrLf & _
                          "    				AND MEE.Period = ME.Period   " & vbCrLf & _
                          "    				AND MEE.ETAForwarder = ME.ETDVendor1  " & vbCrLf & _
                          " 			LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DE.PartNo  " & vbCrLf & _
                          " 				AND MPM.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND MPM.AffiliateID = DE.AffiliateID " & vbCrLf & _
                          " 			WHERE ME.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _
                          "             AND ME.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 			and week = '1' "

        ls_sql = ls_sql + " 			Group By ME.SupplierID, week) W1 " & vbCrLf & _
                          " ON W1.SupplierID = ME.SupplierID " & vbCrLf & _
                          " 			 " & vbCrLf & _
                          " LEFT JOIN (SELECT distinct ME.SupplierID,QtyBox = (SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))), QtyPallet = (SUM((Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))/BoxPallet)),week " & vbCrLf & _
                          " 			--,CBM = SUM(week1*(Length*width*height)) * (SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox)))/ 1000000 " & vbCrLf & _
                          "             ,CBM =  SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox)) * SUM((Length*width*height)/6000) " & vbCrLf & _
                          " 			FROM dbo.PO_Master_Export ME  " & vbCrLf & _
                          " 			LEFT JOIN PO_Detail_Export DE  " & vbCrLf & _
                          " 			ON ME.PONo = DE.PONo  " & vbCrLf & _
                          " 				AND ME.OrderNo1 = DE.OrderNo1  " & vbCrLf & _
                          " 				AND ME.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND ME.AffiliateID = DE.AffiliateID   "

        ls_sql = ls_sql + " 			LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '4' UNION ALL Select * from MS_ETD_Export where week = '5')MEE " & vbCrLf & _
                          " 			ON MEE.AffiliateID = ME.AffiliateID   " & vbCrLf & _
                          "    				AND MEE.SupplierID = ME.SupplierID   " & vbCrLf & _
                          "    				AND MEE.Period = ME.Period   " & vbCrLf & _
                          "    				AND MEE.ETAForwarder = ME.ETDVendor1  " & vbCrLf & _
                          " 			LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DE.PartNo  " & vbCrLf & _
                          " 				AND MPM.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND MPM.AffiliateID = DE.AffiliateID "

        ls_sql = ls_sql + " 			WHERE ME.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _
                          "             AND ME.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 			and week = '2'  " & vbCrLf & _
                          " 			Group By ME.SupplierID, week) W2 " & vbCrLf & _
                          " ON W2.SupplierID = ME.SupplierID " & vbCrLf & _
                          " LEFT JOIN (SELECT distinct ME.SupplierID,QtyBox = (SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))), QtyPallet = (SUM((Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))/BoxPallet)),week " & vbCrLf & _
                          " 			--,CBM = SUM(week1*(Length*width*height)) * (SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox)))/ 1000000 " & vbCrLf & _
                          "             ,CBM =  SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox)) * SUM((Length*width*height)/6000) " & vbCrLf & _
                          " 			FROM dbo.PO_Master_Export ME  " & vbCrLf & _
                          " 			LEFT JOIN PO_Detail_Export DE  " & vbCrLf & _
                          " 			ON ME.PONo = DE.PONo  " & vbCrLf & _
                          " 				AND ME.OrderNo1 = DE.OrderNo1  "

        ls_sql = ls_sql + " 				AND ME.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND ME.AffiliateID = DE.AffiliateID   " & vbCrLf & _
                          " 			LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '4' UNION ALL Select * from MS_ETD_Export where week = '5')MEE " & vbCrLf & _
                          " 			ON MEE.AffiliateID = ME.AffiliateID   " & vbCrLf & _
                          "    				AND MEE.SupplierID = ME.SupplierID   " & vbCrLf & _
                          "    				AND MEE.Period = ME.Period   " & vbCrLf & _
                          "    				AND MEE.ETAForwarder = ME.ETDVendor1  " & vbCrLf & _
                          " 			LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DE.PartNo  "

        ls_sql = ls_sql + " 				AND MPM.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND MPM.AffiliateID = DE.AffiliateID " & vbCrLf & _
                          " 			WHERE ME.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _
                          "             AND ME.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 			and week = '3' " & vbCrLf & _
                          " 			Group By ME.SupplierID, week)W3 " & vbCrLf & _
                          " ON W3.SupplierID = ME.SupplierID " & vbCrLf & _
                          " LEFT JOIN (SELECT distinct ME.SupplierID,QtyBox = (SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))), QtyPallet = (SUM((Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))/BoxPallet)),week " & vbCrLf & _
                          " 			--,CBM = SUM(week1*(Length*width*height)) * (SUM(Week1/QtyBox))/ 1000000 " & vbCrLf & _
                          "             ,CBM =  SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox)) * SUM((Length*width*height)/6000) " & vbCrLf & _
                          " 			FROM dbo.PO_Master_Export ME  " & vbCrLf & _
                          " 			LEFT JOIN PO_Detail_Export DE  "

        ls_sql = ls_sql + " 			ON ME.PONo = DE.PONo  " & vbCrLf & _
                          " 				AND ME.OrderNo1 = DE.OrderNo1  " & vbCrLf & _
                          " 				AND ME.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND ME.AffiliateID = DE.AffiliateID   " & vbCrLf & _
                          " 			LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '4' UNION ALL Select * from MS_ETD_Export where week = '5')MEE " & vbCrLf & _
                          " 			ON MEE.AffiliateID = ME.AffiliateID   " & vbCrLf & _
                          "    				AND MEE.SupplierID = ME.SupplierID   " & vbCrLf & _
                          "    				AND MEE.Period = ME.Period   "

        ls_sql = ls_sql + "    				AND MEE.ETAForwarder = ME.ETDVendor1  " & vbCrLf & _
                          " 			LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DE.PartNo  " & vbCrLf & _
                          " 				AND MPM.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND MPM.AffiliateID = DE.AffiliateID " & vbCrLf & _
                          " 			WHERE ME.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _
                          "             AND ME.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 			and week = '4'   " & vbCrLf & _
                          " 			Group By ME.SupplierID, week) W4 " & vbCrLf & _
                          "  ON W4.SupplierID = ME.SupplierID " & vbCrLf & _
                          " LEFT JOIN (SELECT distinct ME.SupplierID,QtyBox = (SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))), QtyPallet = (SUM((Week1/ISNULL(DE.POQtyBox,MPM.QtyBox))/BoxPallet)),week " & vbCrLf & _
                          " 			--,CBM = SUM(week1*(Length*width*height)) * (SUM(Week1/QtyBox))/ 1000000 " & vbCrLf & _
                          "             ,CBM =  SUM(Week1/ISNULL(DE.POQtyBox,MPM.QtyBox)) * SUM((Length*width*height)/6000) " & vbCrLf & _
                          " 			FROM dbo.PO_Master_Export ME  " & vbCrLf & _
                          " 			LEFT JOIN PO_Detail_Export DE  "

        ls_sql = ls_sql + " 			ON ME.PONo = DE.PONo  " & vbCrLf & _
                          " 				AND ME.OrderNo1 = DE.OrderNo1  " & vbCrLf & _
                          " 				AND ME.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND ME.AffiliateID = DE.AffiliateID   " & vbCrLf & _
                          " 			LEFT JOIN (Select * from MS_ETD_Export where week = '1'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '2'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '3'  " & vbCrLf & _
                          " 						UNION ALL Select * from MS_ETD_Export where week = '4' UNION ALL Select * from MS_ETD_Export where week = '5')MEE " & vbCrLf & _
                          " 			ON MEE.AffiliateID = ME.AffiliateID   " & vbCrLf & _
                          "    				AND MEE.SupplierID = ME.SupplierID   " & vbCrLf & _
                          "    				AND MEE.Period = ME.Period   "

        ls_sql = ls_sql + "    				AND MEE.ETAForwarder = ME.ETDVendor1  " & vbCrLf & _
                          " 			LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DE.PartNo  " & vbCrLf & _
                          " 				AND MPM.SupplierID = DE.SupplierID  " & vbCrLf & _
                          " 				AND MPM.AffiliateID = DE.AffiliateID " & vbCrLf & _
                          " 			WHERE ME.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _
                          "             AND ME.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf & _
                          " 			and week = '5'   " & vbCrLf & _
                          " 			Group By ME.SupplierID, week) W5 " & vbCrLf & _
                          "  ON W5.SupplierID = ME.SupplierID " & vbCrLf & _
                          " WHERE ME.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
                          " and ShipCls = 'B' " & vbCrLf & _
                          " AND ME.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using
        dtDetail = ds.Tables(0)
    End Sub

#End Region

    Private Sub Approve_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles Approve.Callback
        GetExcel()
    End Sub
End Class