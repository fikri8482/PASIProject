Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel


Public Class ForecastRegularMonthly
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
            " " & vbCrLf & _
            "var PeriodTo = new Date(); " & vbCrLf & _
            "dtPOPeriod.SetDate(PeriodTo); " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(dtPOPeriod, dtPOPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_GridHeader(ByVal pYear As Integer)
        grid.VisibleColumns(4).Caption = "JUL " & pYear
        grid.VisibleColumns(5).Caption = "AUG " & pYear
        grid.VisibleColumns(6).Caption = "SEP " & pYear
        grid.VisibleColumns(7).Caption = "OCT " & pYear
        grid.VisibleColumns(8).Caption = "NOV " & pYear
        grid.VisibleColumns(9).Caption = "DEC " & pYear

        grid.VisibleColumns(10).Caption = "JAN " & CInt(pYear) + 1
        grid.VisibleColumns(11).Caption = "FEB " & CInt(pYear) + 1
        grid.VisibleColumns(12).Caption = "MAR " & CInt(pYear) + 1
        grid.VisibleColumns(13).Caption = "APR " & CInt(pYear) + 1
        grid.VisibleColumns(14).Caption = "MAY " & CInt(pYear) + 1
        grid.VisibleColumns(15).Caption = "JUN " & CInt(pYear) + 1
    End Sub
    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Dim ls_filter As String = ""
            ls_SQL = ""

            'For i = 1 To 12
            '    grid.VisibleColumns(4 + i).Caption = Format("2018-" & i, "MMM")
            'Next
            grid.VisibleColumns(4).Caption = "JUL " & dtPOPeriod.Text
            grid.VisibleColumns(5).Caption = "AUG " & dtPOPeriod.Text
            grid.VisibleColumns(6).Caption = "SEP " & dtPOPeriod.Text
            grid.VisibleColumns(7).Caption = "OCT " & dtPOPeriod.Text
            grid.VisibleColumns(8).Caption = "NOV " & dtPOPeriod.Text
            grid.VisibleColumns(9).Caption = "DEC " & dtPOPeriod.Text

            grid.VisibleColumns(10).Caption = "JAN " & CInt(dtPOPeriod.Text) + 1
            grid.VisibleColumns(11).Caption = "FEB " & CInt(dtPOPeriod.Text) + 1
            grid.VisibleColumns(12).Caption = "MAR " & CInt(dtPOPeriod.Text) + 1
            grid.VisibleColumns(13).Caption = "APR " & CInt(dtPOPeriod.Text) + 1
            grid.VisibleColumns(14).Caption = "MAY " & CInt(dtPOPeriod.Text) + 1
            grid.VisibleColumns(15).Caption = "JUN " & CInt(dtPOPeriod.Text) + 1

            ls_SQL = "EXEC sp_SelectForecastMonthly '" & dtPOPeriod.Text & "','" & cbopart.Text & "'"

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

    Private Sub up_FillCombo(ByVal pYear As String)
        ls_SQL = ""
        ls_SQL = "SELECT Distinct RTRIM(FM.PartNo) PartCode, PartName  " & vbCrLf & _
                 "from ForecastMonthly FM" & vbCrLf & _
                 "left join MS_Parts MP ON FM.PartNo = MP.PartNo" & vbCrLf & _
                 "Where Year = '" & pYear & "'" & vbCrLf & _
                 "order by PartCode" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbopart
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartCode")
                .Columns(0).Width = 90
                .Columns.Add("PartName")
                .Columns(1).Width = 400

                .TextField = "PartCode"
                .DataBind()
                .SelectedIndex = 0
                'txtPartNo.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
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

    Private Function GetSummaryOutStanding() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""

                'SUPPLIER CODE
                If Trim(cbopart.Text) <> "==ALL==" And Trim(cbopart.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FD.PartNo = '" & Trim(cbopart.Text) & "' " & vbCrLf
                End If

                ls_sql = " Select FD.* " & vbCrLf & _
                      " From ForecastMonthly FD " & vbCrLf & _
                      " Left Join MS_PartMapping MPM ON FD.PartNo = MPM.PartNo And FD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                      " Left Join MS_Parts MP ON FD.PartNo = MP.PartNo " & vbCrLf & _
                      " Where Year = '" & dtPOPeriod.Text & "' " & vbCrLf & _
                      "  "


                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + " " & vbCrLf & _
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

                ls_sql = "EXEC sp_SelectForecastMonthly '" & dtPOPeriod.Text & "','" & cbopart.Text & "'"

                Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
                Dim ds As New DataSet
                sqlDA.SelectCommand.CommandTimeout = 200
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

    Private Sub EpPlusDrawAllBorders(ByVal Rg As ExcelRange)
        With Rg
            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
        End With
    End Sub

    Private Sub epplusExportExcelNew(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "TemplateForecastReportRegularMonthly " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Forecast\Import\" & tempFile & "")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            'Dim fi As New FileInfo(NewFileName)

            'Dim exl As New ExcelPackage(fi)
            'Dim ws As ExcelWorksheet
            Dim ExcelBook As excel.Workbook
            Dim ExcelSheet As excel.Worksheet

            Dim xlApp = New excel.Application
            Dim ls_file As String = NewFileName
            '
            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(pSheetName), excel.Worksheet)

            'ws = exl.Workbook.Worksheets(pSheetName)
            With ExcelSheet
                .Cells(3, 4).Value = ": " & Format(dtPOPeriod.Value, "yyyy")
                '.Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text)
                .Cells(4, 4).Value = ": " & Trim(cbopart.Text)
                '.Cells(6, 4).Value = ": " & Trim(cboAffiliateCode.Text)


                '.Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
                Dim ds As New DataSet
                ds = GetSummaryOutStanding2()
                If ds.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        ExcelSheet.Range("A" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("PartNo"))
                        ExcelSheet.Range("B" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("AffiliateID"))
                        ExcelSheet.Range("C" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("SupplierID"))
                        ExcelSheet.Range("D" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("Data"))
                        ExcelSheet.Range("E" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F1"))
                        ExcelSheet.Range("F" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F2"))
                        ExcelSheet.Range("G" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F3"))
                        ExcelSheet.Range("H" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F4"))
                        ExcelSheet.Range("I" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F5"))
                        ExcelSheet.Range("J" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F6"))
                        ExcelSheet.Range("K" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F7"))
                        ExcelSheet.Range("L" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F8"))
                        ExcelSheet.Range("M" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F9"))
                        ExcelSheet.Range("N" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F9"))
                        ExcelSheet.Range("O" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F11"))
                        ExcelSheet.Range("P" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("F12"))
                        
                        ExcelSheet.Range("E" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("F" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("G" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("H" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("I" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("J" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("K" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("L" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("M" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("N" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("O" & i + 9).NumberFormat = "#,##0"
                        ExcelSheet.Range("P" & i + 9).NumberFormat = "#,##0"
                        
                        If Trim(ds.Tables(0).Rows(i)("Data")) = "Diff Firm vs Act" Then
                            ExcelSheet.Range("E" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F1")) & " %"
                            ExcelSheet.Range("F" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F2")) & " %"
                            ExcelSheet.Range("G" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F3")) & " %"
                            ExcelSheet.Range("H" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F4")) & " %"
                            ExcelSheet.Range("I" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F5")) & " %"
                            ExcelSheet.Range("J" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F6")) & " %"
                            ExcelSheet.Range("K" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F7")) & " %"
                            ExcelSheet.Range("L" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F8")) & " %"
                            ExcelSheet.Range("M" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F9")) & " %"
                            ExcelSheet.Range("N" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F9")) & " %"
                            ExcelSheet.Range("O" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F11")) & " %"
                            ExcelSheet.Range("P" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F12")) & " %"
                        End If

                        If Trim(ds.Tables(0).Rows(i)("Data")) = "Diff Firm vs Last FC" Then
                            ExcelSheet.Range("E" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F1")) & " %"
                            ExcelSheet.Range("F" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F2")) & " %"
                            ExcelSheet.Range("G" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F3")) & " %"
                            ExcelSheet.Range("H" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F4")) & " %"
                            ExcelSheet.Range("I" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F5")) & " %"
                            ExcelSheet.Range("J" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F6")) & " %"
                            ExcelSheet.Range("K" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F7")) & " %"
                            ExcelSheet.Range("L" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F8")) & " %"
                            ExcelSheet.Range("M" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F9")) & " %"
                            ExcelSheet.Range("N" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F9")) & " %"
                            ExcelSheet.Range("O" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F11")) & " %"
                            ExcelSheet.Range("P" & i + 9).Value = "'" & CInt(ds.Tables(0).Rows(i)("F12")) & " %"
                        End If



                        ''Looping Column
                        'For z = 0 To 30
                        '    'Cek Cls
                        '    If Trim(ds.Tables(0).Rows(i)("C" & z + 1)) = 1 Then
                        '        'Cek rev
                        '        If .Cells(10 + i, 2).Value = 1 Then
                        '            'ExcelSheet.Range(10 + i, 13 + z).Interior.Color = Color.Yellow
                        '            ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Yellow
                        '        ElseIf .Cells(10 + i, 2).Value = 2 Then
                        '            ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Orange
                        '        ElseIf .Cells(10 + i, 2).Value = 2 Then
                        '            ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Green
                        '        End If
                        '    End If
                        'Next
                    Next
                End If

                xlApp.DisplayAlerts = False

                'Dim s As Integer = (ds.Tables(0).Rows.Count / 7) - 1
                'For M = 0 To s
                '    Dim x As Integer = M * 7
                '    Dim A As Integer = x + 10
                '    Dim B As Integer = x + 16
                '    'ExcelSheet.Range("A10:A16").Merge()
                '    ExcelSheet.Range("A" & A & ":A" & B).Merge()
                '    ExcelSheet.Range("B" & A & ":B" & B).Merge()
                '    ExcelSheet.Range("C" & A & ":C" & B).Merge()
                '    ExcelSheet.Range("D" & A & ":D" & B).Merge()
                'Next

                ExcelSheet.Range("A9:C" & ds.Tables(0).Rows.Count + 8).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                ExcelSheet.Range("A9:C" & ds.Tables(0).Rows.Count + 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("D9:P" & ds.Tables(0).Rows.Count + 8).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                ExcelSheet.Range("D9:P" & ds.Tables(0).Rows.Count + 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight

                'ExcelSheet.Range("M9: AJ" & ds.Tables(0).Rows.Count + 8).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("A9: P" & ds.Tables(0).Rows.Count + 8).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("A9: P" & ds.Tables(0).Rows.Count + 8).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("A9: P" & ds.Tables(0).Rows.Count + 8).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("A9: P" & ds.Tables(0).Rows.Count + 8).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("A9: P" & ds.Tables(0).Rows.Count + 8).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("A9: P" & ds.Tables(0).Rows.Count + 8).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            End With


            'exl.Save()
            ExcelBook.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

            'exl = Nothing
            xlApp.Workbooks.Close()
            xlApp.Quit()
            xlApp = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "TemplateForecastReportRegularMonthly " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                .Cells(4, 4).Value = ": " & Trim(cbopart.Text)
               
                '.Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
                '.Cells(10, 1, pData.Rows.Count + 9, 36).AutoFitColumns()
                '.Cells(10, 1, pData.Rows.Count + 9, 36).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center

                .Cells("E8").Value = "JUL " & dtPOPeriod.Text
                .Cells("F8").Value = "AUG " & dtPOPeriod.Text
                .Cells("G8").Value = "SEP " & dtPOPeriod.Text
                .Cells("H8").Value = "OCT " & dtPOPeriod.Text
                .Cells("I8").Value = "NOV " & dtPOPeriod.Text
                .Cells("J8").Value = "DEC " & dtPOPeriod.Text

                .Cells("K8").Value = "JAN " & CInt(dtPOPeriod.Text) + 1
                .Cells("L8").Value = "FEB " & CInt(dtPOPeriod.Text) + 1
                .Cells("M8").Value = "MAR " & CInt(dtPOPeriod.Text) + 1
                .Cells("N8").Value = "APR " & CInt(dtPOPeriod.Text) + 1
                .Cells("O8").Value = "MAY " & CInt(dtPOPeriod.Text) + 1
                .Cells("P8").Value = "JUN " & CInt(dtPOPeriod.Text) + 1

                Dim ds As New DataSet
                ds = GetSummaryOutStanding2()
                If ds.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        .Cells("A" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("PartNo"))

                        If IsDBNull(ds.Tables(0).Rows(i)("AffiliateID")) = True Then
                            .Cells("B" & i + 9).Value = ""
                        Else
                            .Cells("B" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("AffiliateID"))
                        End If

                        .Cells("C" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("SupplierID"))
                        .Cells("D" & i + 9).Value = Trim(ds.Tables(0).Rows(i)("Data"))
                        .Cells("E" & i + 9).Value = ds.Tables(0).Rows(i)("F7")
                        .Cells("F" & i + 9).Value = ds.Tables(0).Rows(i)("F8")
                        .Cells("G" & i + 9).Value = ds.Tables(0).Rows(i)("F9")
                        .Cells("H" & i + 9).Value = ds.Tables(0).Rows(i)("F10")
                        .Cells("I" & i + 9).Value = ds.Tables(0).Rows(i)("F11")
                        .Cells("J" & i + 9).Value = ds.Tables(0).Rows(i)("F12")
                        .Cells("K" & i + 9).Value = ds.Tables(0).Rows(i)("F1")
                        .Cells("L" & i + 9).Value = ds.Tables(0).Rows(i)("F2")
                        .Cells("M" & i + 9).Value = ds.Tables(0).Rows(i)("F3")
                        .Cells("N" & i + 9).Value = ds.Tables(0).Rows(i)("F4")
                        .Cells("O" & i + 9).Value = ds.Tables(0).Rows(i)("F5")
                        .Cells("P" & i + 9).Value = ds.Tables(0).Rows(i)("F6")

                        .Cells("E" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("F" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("G" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("H" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("I" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("J" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("K" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("L" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("M" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("N" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("O" & i + 9).Style.Numberformat.Format = "#,##0"
                        .Cells("P" & i + 9).Style.Numberformat.Format = "#,##0"
                    Next
                End If


                For x = 5 To 16
                    .Cells(9, x, ds.Tables(0).Rows.Count + 8, x).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                Next

                For irow = 0 To ds.Tables(0).Rows.Count - 1
                    If Trim(.Cells("D" & irow + 9).Value) = "Diff Firm vs Act" Then
                        For d = 5 To 16
                            .Cells(irow + 9, d).Value = "" & CInt(.Cells(irow + 9, d).Value) & " %"
                        Next
                    End If
                    If Trim(.Cells("D" & irow + 9).Value) = "Diff Firm vs Last FC" Then
                        For d = 5 To 16
                            .Cells(irow + 9, d).Value = "" & CInt(.Cells(irow + 9, d).Value) & " %"
                        Next
                    End If
                Next

                'xlApp.DisplayAlerts = False

                'Dim s As Integer = (ds.Tables(0).Rows.Count / 7) - 1
                'For M = 0 To s
                '    Dim x As Integer = M * 7
                '    Dim A As Integer = x + 10
                '    Dim B As Integer = x + 16
                '    'ExcelSheet.Range("A10:A16").Merge()
                '    .Cells(A, 1, B, 1).Merge = True
                '    .Cells(A, 2, B, 2).Merge = True
                '    .Cells(A, 3, B, 3).Merge = True
                '    '.Cells(A, 4, B, 4).Merge = True
                'Next

                Dim rgAll As ExcelRange = .Cells(9, 1, ds.Tables(0).Rows.Count + 8, 16)
                EpPlusDrawAllBorders(rgAll)
            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

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
                Call up_GridLoadWhenEventChange()
                Call up_Initialize()
                Call up_FillCombo(Year(Now))
                Call up_GridHeader(Year(Now))
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 5, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("G01Msg")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 5, False, clsAppearance.PagerMode.ShowPager)

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
                    FileName = "TemplateForecastReportRegularMonthly.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:10", psERR)
                        'Call epplusExportExcelNew(FilePath, "Sheet1", dtProd, "A:9", psERR)
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("G01Msg") = lblInfo.Text
        End Try

        If (Not IsNothing(Session("G01Msg"))) Then grid.JSProperties("cpMessage") = Session("G01Msg") : Session.Remove("G01Msg")

    End Sub

    Private Sub grid_CustomColumnDisplayText(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewColumnDisplayTextEventArgs) Handles grid.CustomColumnDisplayText
        With e.Column
            For i = 1 To 12
                If .FieldName = "F" & i Then
                    If e.GetFieldValue("Data") = "Diff Firm vs Act" Then
                        e.DisplayText = e.GetFieldValue("F" & i) & " %"
                    End If
                    If e.GetFieldValue("Data") = "Diff Firm vs Last FC" Then
                        e.DisplayText = e.GetFieldValue("F" & i) & " %"
                    End If
                End If
            Next
        End With
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        With e.DataColumn
            'If .FieldName = "F1" Then
            '    If e.GetValue("Data") = "Cumm FC vs PO Affiliate" Then

            '    End If
            'End If
            '    If .FieldName = "F1" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 1) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F2" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 2) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F3" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 3) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F4" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 4) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F5" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 5) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F6" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 6) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F7" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 7) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F8" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 8) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F9" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 9) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F10" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 10) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F11" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 11) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F12" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 12) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F13" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 13) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F14" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 14) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F15" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 15) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F16" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 16) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F17" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 17) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F18" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 18) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F19" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 19) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F20" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 20) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F21" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 21) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F22" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 22) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F23" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 23) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F24" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 24) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F25" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 25) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F26" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 26) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F27" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 27) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F28" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 28) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F29" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 29) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F30" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 30) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
            '    If .FieldName = "F31" Then
            '        If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 31) = True Then
            '            If e.GetValue("Rev") = "1" Then
            '                e.Cell.BackColor = Color.Yellow
            '            ElseIf e.GetValue("Rev") = "2" Then
            '                e.Cell.BackColor = Color.Orange
            '            ElseIf e.GetValue("Rev") = "3" Then
            '                e.Cell.BackColor = Color.Green
            '            End If
            '        End If
            '    End If
        End With
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region

    Private Sub cboPart_Callback(sender As Object, e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cbopart.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)

        Call up_FillCombo(pAction)
        Call up_GridHeader(pAction)
    End Sub
End Class