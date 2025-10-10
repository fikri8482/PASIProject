Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel


Public Class ForecastRegular
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
            "    " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "var PeriodTo = new Date(); " & vbCrLf & _
            "dtPOPeriod.SetDate(PeriodTo); " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(dtPOPeriod, dtPOPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_InsertDiff()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""
            Dim ls_rev As String = ""

            ls_SQL = " Select ISNULL(Max(Rev),'') Rev From ForecastDaily Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' And AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' "
            Dim cmd As New SqlCommand(ls_SQL, sqlConn)
            cmd.CommandTimeout = 300
            Dim sqlDA As New SqlDataAdapter
            sqlDA.SelectCommand = cmd
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 300
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                ls_rev = ds.Tables(0).Rows(0)("Rev")
                If ls_rev <> "" Then
                    ls_SQL = "Delete ForecastDailyDiff Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' And AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf & _
                             "Exec sp_InsertDiff '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01','" & ls_rev & "','" & Trim(cboAffiliateCode.Text) & "' "
                End If
            End If

            Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()
        End Using
    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Dim ls_filter As String = ""
            ls_SQL = ""

            For i = 1 To 31
                grid.VisibleColumns(4 + i).Caption = i & "-" & Format(dtPOPeriod.Value, "MMM")
            Next

            up_InsertDiff()

            ls_SQL = uf_Sql()

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
                If Trim(txtPartNo.Text) <> "==ALL==" And Trim(txtPartNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FD.PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf
                End If

                ''REVISION
                'If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                '    ls_filter = ls_filter + _
                '                  "                      AND FD.Rev = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
                'End If

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

                ls_sql = uf_Sql()

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
            'ls_SQL = "--SELECT AffiliateID = '==ALL==', AffiliateName = '==ALL=='" & vbCrLf & _
            '         " --UNION ALL " & vbCrLf & _
            '         "SELECT AffiliateID = RTRIM(AffiliateID), AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate Where isnull(overseascls, '0') = '0'"
            ls_SQL = "SELECT AffiliateID = RTRIM(AffiliateID), AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate Where isnull(overseascls, '0') = '0'"
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
                .SelectedIndex = 0

                .TextField = "AffiliateID"
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

    'Private Sub epplusExportExcelNew(ByVal pFilename As String, ByVal pSheetName As String,
    '                          ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

    '    Try
    '        Dim tempFile As String = "TemplateForecastReportRegular " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
    '            .Cells(5, 4).Value = ": " & Trim(txtPartNo.Text)
    '            '.Cells(6, 4).Value = ": " & Trim(cboAffiliateCode.Text)

    '            'ExcelSheet.Range("I8").Value = "Forecast Quantity " & Format(dtPOPeriod.Value, "MMM yyyy")
    '            'ExcelSheet.Range("J8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Day, 1, dtPOPeriod.Value), "MMM yyyy")
    '            'ExcelSheet.Range("K8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Day, 2, dtPOPeriod.Value), "MMM yyyy")
    '            'ExcelSheet.Range("L8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Day, 3, dtPOPeriod.Value), "MMM yyyy")

    '            '.Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
    '            Dim ds As New DataSet
    '            ds = GetSummaryOutStanding2()
    '            If ds.Tables(0).Rows.Count > 0 Then
    '                For i = 0 To ds.Tables(0).Rows.Count - 1
    '                    ExcelSheet.Range("A" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("PartNo"))
    '                    ExcelSheet.Range("B" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("SupplierID"))
    '                    ExcelSheet.Range("C" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("AffiliateID"))
    '                    ExcelSheet.Range("D" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("MPQ"))
    '                    ExcelSheet.Range("E" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Data"))
    '                    ExcelSheet.Range("F" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F1"))
    '                    ExcelSheet.Range("G" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F2"))
    '                    ExcelSheet.Range("H" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F3"))
    '                    ExcelSheet.Range("I" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F4"))
    '                    ExcelSheet.Range("J" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F5"))
    '                    ExcelSheet.Range("K" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F6"))
    '                    ExcelSheet.Range("L" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F7"))
    '                    ExcelSheet.Range("M" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F8"))
    '                    ExcelSheet.Range("N" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F9"))
    '                    ExcelSheet.Range("O" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F10"))
    '                    ExcelSheet.Range("P" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F11"))
    '                    ExcelSheet.Range("Q" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F12"))
    '                    ExcelSheet.Range("R" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F13"))
    '                    ExcelSheet.Range("S" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F14"))
    '                    ExcelSheet.Range("T" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F15"))
    '                    ExcelSheet.Range("U" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F16"))
    '                    ExcelSheet.Range("V" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F17"))
    '                    ExcelSheet.Range("W" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F18"))
    '                    ExcelSheet.Range("X" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F19"))
    '                    ExcelSheet.Range("Y" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F20"))
    '                    ExcelSheet.Range("Z" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F21"))
    '                    ExcelSheet.Range("AA" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F22"))
    '                    ExcelSheet.Range("AB" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F23"))
    '                    ExcelSheet.Range("AC" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F24"))
    '                    ExcelSheet.Range("AD" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F25"))
    '                    ExcelSheet.Range("AE" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F26"))
    '                    ExcelSheet.Range("AF" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F27"))
    '                    ExcelSheet.Range("AG" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F28"))
    '                    ExcelSheet.Range("AH" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F29"))
    '                    ExcelSheet.Range("AI" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F30"))
    '                    ExcelSheet.Range("AJ" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F31"))

    '                    'ExcelSheet.Range("A" & i + 10).NumberFormat = "#,##0"
    '                    'ExcelSheet.Range("B" & i + 10).NumberFormat = "#,##0"
    '                    'ExcelSheet.Range("C" & i + 10).NumberFormat = "#,##0"
    '                    'ExcelSheet.Range("D" & i + 10).NumberFormat = "#,##0"
    '                    'ExcelSheet.Range("E" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("F" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("G" & i + 10).NumberFormat = "#,##0"
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

    '                    If Trim(ds.Tables(0).Rows(i)("Data")) = "Cumm FC vs PO Affiliate" Then
    '                        ExcelSheet.Range("F" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F1")) & " %"
    '                        ExcelSheet.Range("G" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F2")) & " %"
    '                        ExcelSheet.Range("H" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F3")) & " %"
    '                        ExcelSheet.Range("I" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F4")) & " %"
    '                        ExcelSheet.Range("J" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F5")) & " %"
    '                        ExcelSheet.Range("K" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F6")) & " %"
    '                        ExcelSheet.Range("L" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F7")) & " %"
    '                        ExcelSheet.Range("M" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F8")) & " %"
    '                        ExcelSheet.Range("N" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F9")) & " %"
    '                        ExcelSheet.Range("O" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F10")) & " %"
    '                        ExcelSheet.Range("P" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F11")) & " %"
    '                        ExcelSheet.Range("Q" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F12")) & " %"
    '                        ExcelSheet.Range("R" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F13")) & " %"
    '                        ExcelSheet.Range("S" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F14")) & " %"
    '                        ExcelSheet.Range("T" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F15")) & " %"
    '                        ExcelSheet.Range("U" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F16")) & " %"
    '                        ExcelSheet.Range("V" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F17")) & " %"
    '                        ExcelSheet.Range("W" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F18")) & " %"
    '                        ExcelSheet.Range("X" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F19")) & " %"
    '                        ExcelSheet.Range("Y" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F20")) & " %"
    '                        ExcelSheet.Range("Z" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F21")) & " %"
    '                        ExcelSheet.Range("AA" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F22")) & " %"
    '                        ExcelSheet.Range("AB" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F23")) & " %"
    '                        ExcelSheet.Range("AC" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F24")) & " %"
    '                        ExcelSheet.Range("AD" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F25")) & " %"
    '                        ExcelSheet.Range("AE" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F26")) & " %"
    '                        ExcelSheet.Range("AF" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F27")) & " %"
    '                        ExcelSheet.Range("AG" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F28")) & " %"
    '                        ExcelSheet.Range("AH" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F29")) & " %"
    '                        ExcelSheet.Range("AI" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F30")) & " %"
    '                        ExcelSheet.Range("AJ" & i + 10).Value = "'" & CInt(ds.Tables(0).Rows(i)("F31")) & " %"
    '                    End If



    '                    ''Looping Column
    '                    'For z = 0 To 30
    '                    '    'Cek Cls
    '                    '    If Trim(ds.Tables(0).Rows(i)("C" & z + 1)) = 1 Then
    '                    '        'Cek rev
    '                    '        If .Cells(10 + i, 2).Value = 1 Then
    '                    '            'ExcelSheet.Range(10 + i, 13 + z).Interior.Color = Color.Yellow
    '                    '            ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Yellow
    '                    '        ElseIf .Cells(10 + i, 2).Value = 2 Then
    '                    '            ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Orange
    '                    '        ElseIf .Cells(10 + i, 2).Value = 2 Then
    '                    '            ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Green
    '                    '        End If
    '                    '    End If
    '                    'Next
    '                Next
    '            End If

    '            'ExcelSheet.Range("A10:A16").Merge()

    '            xlApp.DisplayAlerts = False

    '            Dim s As Integer = (ds.Tables(0).Rows.Count / 7) - 1
    '            For M = 0 To s
    '                Dim x As Integer = M * 7
    '                Dim A As Integer = x + 10
    '                Dim B As Integer = x + 16
    '                'ExcelSheet.Range("A10:A16").Merge()
    '                ExcelSheet.Range("A" & A & ":A" & B).Merge()
    '                ExcelSheet.Range("B" & A & ":B" & B).Merge()
    '                ExcelSheet.Range("C" & A & ":C" & B).Merge()
    '                ExcelSheet.Range("D" & A & ":D" & B).Merge()
    '            Next

    '            ExcelSheet.Range("A10:D" & ds.Tables(0).Rows.Count + 8).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
    '            ExcelSheet.Range("A10:D" & ds.Tables(0).Rows.Count + 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

    '            ExcelSheet.Range("M10: AJ" & ds.Tables(0).Rows.Count + 8).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

    '            ExcelSheet.Range("A10: AJ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AJ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AJ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AJ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AJ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AJ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

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

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "TemplateForecastReportRegular " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                .Cells(3, 4).Value = ": " & Format(dtPOPeriod.Value, "MMM yyyy")
                .Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text)
                .Cells(5, 4).Value = ": " & Trim(txtPartNo.Text)

                '.Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
                '.Cells(10, 1, pData.Rows.Count + 9, 36).AutoFitColumns()
                '.Cells(10, 1, pData.Rows.Count + 9, 36).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center

                Dim ds As New DataSet
                ds = GetSummaryOutStanding2()
                If ds.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        .Cells("A" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("PartNo"))

                        If IsDBNull(ds.Tables(0).Rows(i)("SupplierID")) = True Then
                            .Cells("B" & i + 10).Value = ""
                        Else
                            .Cells("B" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("SupplierID"))
                        End If
                        '.Cells("B" & i + 10).Value = IIf(IsDBNull(ds.Tables(0).Rows(i)("SupplierID")) = True, " ", Trim(ds.Tables(0).Rows(i)("SupplierID")))
                        .Cells("C" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("AffiliateID"))

                        If IsDBNull(ds.Tables(0).Rows(i)("MPQ")) = True Then
                            .Cells("D" & i + 10).Value = ""
                        Else
                            .Cells("D" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("MPQ"))
                        End If
                        '.Cells("D" & i + 10).Value = IIf(IsDBNull(ds.Tables(0).Rows(i)("MPQ")) = True, "0", Trim(ds.Tables(0).Rows(i)("MPQ")))

                        .Cells("E" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Data"))
                        .Cells("F" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F1"))
                        .Cells("G" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F2"))
                        .Cells("H" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F3"))
                        .Cells("I" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F4"))
                        .Cells("J" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F5"))
                        .Cells("K" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F6"))
                        .Cells("L" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F7"))
                        .Cells("M" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F8"))
                        .Cells("N" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F9"))
                        .Cells("O" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F10"))
                        .Cells("P" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F11"))
                        .Cells("Q" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F12"))
                        .Cells("R" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F13"))
                        .Cells("S" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F14"))
                        .Cells("T" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F15"))
                        .Cells("U" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F16"))
                        .Cells("V" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F17"))
                        .Cells("W" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F18"))
                        .Cells("X" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F19"))
                        .Cells("Y" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F20"))
                        .Cells("Z" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F21"))
                        .Cells("AA" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F22"))
                        .Cells("AB" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F23"))
                        .Cells("AC" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F24"))
                        .Cells("AD" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F25"))
                        .Cells("AE" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F26"))
                        .Cells("AF" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F27"))
                        .Cells("AG" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F28"))
                        .Cells("AH" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F29"))
                        .Cells("AI" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F30"))
                        .Cells("AJ" & i + 10).Value = CInt(ds.Tables(0).Rows(i)("F31"))
                    Next
                End If


                'For x = 6 To 36
                '    .Cells(10, x, ds.Tables(0).Rows.Count + 9, x).Style.Numberformat.Format = "#,##0"
                'Next

                For irow = 0 To ds.Tables(0).Rows.Count - 1
                    If Trim(.Cells("E" & irow + 10).Value) = "Cumm FC vs PO Affiliate" Then
                        For d = 6 To 36
                            '.Cells(irow + 10, d).Value = "" & CInt(.Cells(irow + 10, d).Value) & "%"
                            .Cells(irow + 10, d).Value = CInt(.Cells(irow + 10, d).Value) / 100
                            .Cells(irow + 10, d).Style.Numberformat.Format = "0%"
                        Next
                    End If
                    If Trim(.Cells("E" & irow + 10).Value) = "Diff FC vs PO Affiliate" Then
                        For d = 6 To 36
                            '.Cells(irow + 10, d).Value = "" & CInt(.Cells(irow + 10, d).Value) & "%"
                            .Cells(irow + 10, d).Value = CInt(.Cells(irow + 10, d).Value) / 100
                            .Cells(irow + 10, d).Style.Numberformat.Format = "0%"
                        Next
                    End If
                Next

                'https://stackoverflow.com/questions/40209636/epplus-number-format/40214134?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
                'xlApp.DisplayAlerts = False

                Dim s As Integer = (ds.Tables(0).Rows.Count / 7) - 1
                For M = 0 To s
                    Dim x As Integer = M * 7
                    Dim A As Integer = x + 10
                    Dim B As Integer = x + 16
                    'ExcelSheet.Range("A10:A16").Merge()
                    .Cells(A, 1, B, 1).Merge = True
                    .Cells(A, 2, B, 2).Merge = True
                    .Cells(A, 3, B, 3).Merge = True
                    .Cells(A, 4, B, 4).Merge = True
                Next

                Dim rgAll As ExcelRange = .Cells(10, 1, ds.Tables(0).Rows.Count + 9, 36)
                EpPlusDrawAllBorders(rgAll)
            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Function uf_Sql() As String
        
            Dim ls_filter As String = ""
            ls_SQL = ""

            'AFFILIATE CODE
            If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            End If
            'PART CODE
            If Trim(txtPartNo.Text) <> "==ALL==" And Trim(txtPartNo.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf
            End If


        ls_SQL = " WITH IBD (Period,Rev,PartNo, SupplierID, AffiliateID, MPQ , Data , Total  " & vbCrLf & _
                  " ,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15 " & vbCrLf & _
                  " ,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31,Seq) " & vbCrLf & _
                  " AS ( " & vbCrLf & _
                  " --Forecast Rev 0 " & vbCrLf & _
                  " Select FD0.Period,FD0.Rev,FD.PartNo, PM.SupplierID, FD.AffiliateID, MPQ = PM.MOQ, Data = 'Forecast rev00', Total = ISNULL(FD0.ForecastQty1,0) " & vbCrLf & _
                  " ,F1=ISNULL(FD0.F1,0),F2=ISNULL(FD0.F2,0),F3=ISNULL(FD0.F3,0),F4=ISNULL(FD0.F4,0),F5=ISNULL(FD0.F5,0),F6=ISNULL(FD0.F6,0),F7=ISNULL(FD0.F7,0),F8=ISNULL(FD0.F8,0),F9=ISNULL(FD0.F9,0),F10=ISNULL(FD0.F10,0) " & vbCrLf & _
                  " ,F11=ISNULL(FD0.F11,0),F12=ISNULL(FD0.F12,0),F13=ISNULL(FD0.F13,0),F14=ISNULL(FD0.F14,0),F15=ISNULL(FD0.F15,0),F16=ISNULL(FD0.F16,0),F17=ISNULL(FD0.F17,0),F18=ISNULL(FD0.F18,0),F19=ISNULL(FD0.F19,0),F20=ISNULL(FD0.F20,0) " & vbCrLf & _
                  " ,F21=ISNULL(FD0.F21,0),F22=ISNULL(FD0.F22,0),F23=ISNULL(FD0.F23,0),F24=ISNULL(FD0.F24,0),F25=ISNULL(FD0.F25,0),F26=ISNULL(FD0.F26,0),F27=ISNULL(FD0.F27,0),F28=ISNULL(FD0.F28,0),F29=ISNULL(FD0.F29,0),F30=ISNULL(FD0.F30,0),F31=ISNULL(FD0.F31,0),Seq=0 " & vbCrLf & _
                  " From (Select PartNo, AffiliateID  " & vbCrLf & _
                  " 	  From ForecastDaily  " & vbCrLf & _
                  " 	  Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + " 	  Group By PartNo, AffiliateID " & vbCrLf & _
                          " 	  ) FD " & vbCrLf & _
                          " Left Join MS_PartMapping PM ON FD.PartNo = PM.PartNo And FD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " Left Join  " & vbCrLf & _
                          " 	( Select Period,Rev,PartNo, AffiliateID,ForecastQty1,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31 " & vbCrLf & _
                          " 	  From ForecastDaily " & vbCrLf & _
                          " 	  Where Rev = '0' " & vbCrLf & _
                          " 		And Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

            ls_SQL = ls_SQL + " 	) FD0 ON FD.PartNo = FD0.PartNo And FD.AffiliateID = FD0.AffiliateID " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " --Forecast Rev 1 " & vbCrLf

        ls_SQL = ls_SQL + " Select FD0.Period,FD0.Rev,FD.PartNo, PM.SupplierID, FD.AffiliateID, MPQ = PM.MOQ, Data = 'Forecast rev01', Total = ISNULL(FD0.ForecastQty1,0) " & vbCrLf & _
                              " ,F1=ISNULL(FD0.F1,0),F2=ISNULL(FD0.F2,0),F3=ISNULL(FD0.F3,0),F4=ISNULL(FD0.F4,0),F5=ISNULL(FD0.F5,0),F6=ISNULL(FD0.F6,0),F7=ISNULL(FD0.F7,0),F8=ISNULL(FD0.F8,0),F9=ISNULL(FD0.F9,0),F10=ISNULL(FD0.F10,0) " & vbCrLf & _
                              " ,F11=ISNULL(FD0.F11,0),F12=ISNULL(FD0.F12,0),F13=ISNULL(FD0.F13,0),F14=ISNULL(FD0.F14,0),F15=ISNULL(FD0.F15,0),F16=ISNULL(FD0.F16,0),F17=ISNULL(FD0.F17,0),F18=ISNULL(FD0.F18,0),F19=ISNULL(FD0.F19,0),F20=ISNULL(FD0.F20,0) " & vbCrLf & _
                              " ,F21=ISNULL(FD0.F21,0),F22=ISNULL(FD0.F22,0),F23=ISNULL(FD0.F23,0),F24=ISNULL(FD0.F24,0),F25=ISNULL(FD0.F25,0),F26=ISNULL(FD0.F26,0),F27=ISNULL(FD0.F27,0),F28=ISNULL(FD0.F28,0),F29=ISNULL(FD0.F29,0),F30=ISNULL(FD0.F30,0),F31=ISNULL(FD0.F31,0),Seq=1 " & vbCrLf & _
                              " From (Select PartNo, AffiliateID  " & vbCrLf & _
                              " 	  From ForecastDaily  " & vbCrLf & _
                              " 	  Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + " 	  Group By PartNo, AffiliateID " & vbCrLf & _
                          " 	  ) FD " & vbCrLf & _
                          " Left Join MS_PartMapping PM ON FD.PartNo = PM.PartNo And FD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " Left Join  " & vbCrLf & _
                          " 	( Select Period,Rev,PartNo, AffiliateID,ForecastQty1,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31 " & vbCrLf

            ls_SQL = ls_SQL + " 	  From ForecastDaily " & vbCrLf & _
                              " 	  Where Rev = '1' " & vbCrLf & _
                              " 		And Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + " 	) FD0 ON FD.PartNo = FD0.PartNo And FD.AffiliateID = FD0.AffiliateID " & vbCrLf & _
                          " UNION ALL " & vbCrLf & _
                          " --Forecast Rev 2 " & vbCrLf & _
                          " Select FD0.Period,FD0.Rev,FD.PartNo, PM.SupplierID, FD.AffiliateID, MPQ = PM.MOQ, Data = 'Forecast rev02', Total = ISNULL(FD0.ForecastQty1,0) " & vbCrLf & _
                          " ,F1=ISNULL(FD0.F1,0),F2=ISNULL(FD0.F2,0),F3=ISNULL(FD0.F3,0),F4=ISNULL(FD0.F4,0),F5=ISNULL(FD0.F5,0),F6=ISNULL(FD0.F6,0),F7=ISNULL(FD0.F7,0),F8=ISNULL(FD0.F8,0),F9=ISNULL(FD0.F9,0),F10=ISNULL(FD0.F10,0) " & vbCrLf & _
                          " ,F11=ISNULL(FD0.F11,0),F12=ISNULL(FD0.F12,0),F13=ISNULL(FD0.F13,0),F14=ISNULL(FD0.F14,0),F15=ISNULL(FD0.F15,0),F16=ISNULL(FD0.F16,0),F17=ISNULL(FD0.F17,0),F18=ISNULL(FD0.F18,0),F19=ISNULL(FD0.F19,0),F20=ISNULL(FD0.F20,0) " & vbCrLf & _
                          " ,F21=ISNULL(FD0.F21,0),F22=ISNULL(FD0.F22,0),F23=ISNULL(FD0.F23,0),F24=ISNULL(FD0.F24,0),F25=ISNULL(FD0.F25,0),F26=ISNULL(FD0.F26,0),F27=ISNULL(FD0.F27,0),F28=ISNULL(FD0.F28,0),F29=ISNULL(FD0.F29,0),F30=ISNULL(FD0.F30,0),F31=ISNULL(FD0.F31,0),Seq=2 " & vbCrLf & _
                          " From (Select PartNo, AffiliateID  " & vbCrLf & _
                          " 	  From ForecastDaily  " & vbCrLf

            ls_SQL = ls_SQL + " 	  Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + " 	  Group By PartNo, AffiliateID " & vbCrLf & _
                          " 	  ) FD " & vbCrLf & _
                          " Left Join MS_PartMapping PM ON FD.PartNo = PM.PartNo And FD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " Left Join  " & vbCrLf & _
                          " 	( Select Period,Rev,PartNo, AffiliateID,ForecastQty1,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31 " & vbCrLf & _
                          " 	  From ForecastDaily " & vbCrLf & _
                          " 	  Where Rev = '2' " & vbCrLf & _
                          " 		And Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

            ls_SQL = ls_SQL + " 	) FD0 ON FD.PartNo = FD0.PartNo And FD.AffiliateID = FD0.AffiliateID " & vbCrLf & _
                              " UNION ALL " & vbCrLf

        ls_SQL = ls_SQL + " --Forecast Rev 3 " & vbCrLf & _
                          " Select FD0.Period,FD0.Rev,FD.PartNo, PM.SupplierID, FD.AffiliateID, MPQ = PM.MOQ, Data = 'Forecast rev03', Total = ISNULL(FD0.ForecastQty1,0) " & vbCrLf & _
                          " ,F1=ISNULL(FD0.F1,0),F2=ISNULL(FD0.F2,0),F3=ISNULL(FD0.F3,0),F4=ISNULL(FD0.F4,0),F5=ISNULL(FD0.F5,0),F6=ISNULL(FD0.F6,0),F7=ISNULL(FD0.F7,0),F8=ISNULL(FD0.F8,0),F9=ISNULL(FD0.F9,0),F10=ISNULL(FD0.F10,0) " & vbCrLf & _
                          " ,F11=ISNULL(FD0.F11,0),F12=ISNULL(FD0.F12,0),F13=ISNULL(FD0.F13,0),F14=ISNULL(FD0.F14,0),F15=ISNULL(FD0.F15,0),F16=ISNULL(FD0.F16,0),F17=ISNULL(FD0.F17,0),F18=ISNULL(FD0.F18,0),F19=ISNULL(FD0.F19,0),F20=ISNULL(FD0.F20,0) " & vbCrLf & _
                          " ,F21=ISNULL(FD0.F21,0),F22=ISNULL(FD0.F22,0),F23=ISNULL(FD0.F23,0),F24=ISNULL(FD0.F24,0),F25=ISNULL(FD0.F25,0),F26=ISNULL(FD0.F26,0),F27=ISNULL(FD0.F27,0),F28=ISNULL(FD0.F28,0),F29=ISNULL(FD0.F29,0),F30=ISNULL(FD0.F30,0),F31=ISNULL(FD0.F31,0),Seq=3 " & vbCrLf & _
                          " From (Select PartNo, AffiliateID  " & vbCrLf & _
                          " 	  From ForecastDaily  " & vbCrLf & _
                          " 	  Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

            ls_SQL = ls_SQL + " 	  Group By PartNo, AffiliateID " & vbCrLf & _
                              " 	  ) FD " & vbCrLf & _
                              " Left Join MS_PartMapping PM ON FD.PartNo = PM.PartNo And FD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                              " Left Join  " & vbCrLf

        ls_SQL = ls_SQL + " 	( Select Period,Rev,PartNo, AffiliateID,ForecastQty1,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31 " & vbCrLf & _
                              " 	  From ForecastDaily " & vbCrLf & _
                              " 	  Where Rev = '3' " & vbCrLf & _
                              " 		And Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + " 	) FD0 ON FD.PartNo = FD0.PartNo And FD.AffiliateID = FD0.AffiliateID " & vbCrLf & _
                          " UNION ALL " & vbCrLf & _
                          " --PO Affiliate Kanban " & vbCrLf & _
                          " Select Period,Rev,PartNo,SupplierID,AffiliateID,MPQ,Data = ISNULL(Data,'PO Affiliate'),Total, " & vbCrLf & _
                          " [1] = ISNULL([1],0),[2] = ISNULL([2],0),[3] = ISNULL([3],0),[4] = ISNULL([4],0),[5] = ISNULL([5],0),[6] = ISNULL([6],0),[7] = ISNULL([7],0),[8] = ISNULL([8],0),[9] = ISNULL([9],0),[10] = ISNULL([10],0), " & vbCrLf & _
                          " [11] = ISNULL([11],0),[12] = ISNULL([12],0),[13] = ISNULL([13],0),[14] = ISNULL([14],0),[15] = ISNULL([15],0),[16] = ISNULL([16],0),[17] = ISNULL([17],0),[18] = ISNULL([18],0),[19] = ISNULL([19],0),[20] = ISNULL([20],0), " & vbCrLf & _
                          " [21] = ISNULL([21],0),[22] = ISNULL([22],0),[23] = ISNULL([23],0),[24] = ISNULL([24],0),[25] = ISNULL([25],0),[26] = ISNULL([26],0),[27] = ISNULL([27],0),[28] = ISNULL([28],0),[29] = ISNULL([29],0),[30] = ISNULL([30],0),[31] = ISNULL([31],0),Seq=4 " & vbCrLf & _
                          " FROM ( " & vbCrLf & _
                          " 	Select Period='',Rev='',FD.PartNo,PO.SupplierID,FD.AffiliateID,PO.MPQ,PO.Data,PO.Total,PO.KanbanDate,PO.KanbanQty " & vbCrLf & _
                          " 	From  " & vbCrLf

        ls_SQL = ls_SQL + " 			(Select PartNo, AffiliateID  " & vbCrLf & _
                          " 			  From ForecastDaily  " & vbCrLf & _
                          " 			  Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf

        ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + " 			  Group By PartNo, AffiliateID " & vbCrLf & _
                          " 			 ) FD " & vbCrLf & _
                          " 		Left Join ( " & vbCrLf & _
                          " 		Select KD.PartNo, PM.SupplierID, KD.AffiliateID, MPQ = PM.MOQ, Data = 'PO Affiliate', Total = '0', KanbanDate = datename(day,KM.KanbanDate), KD.KanbanQty " & vbCrLf & _
                          " 		From Kanban_Master KM " & vbCrLf & _
                          " 		Left Join Kanban_Detail KD ON KM.KanbanNo = KD.KanbanNo And KM.AffiliateID = KD.AffiliateID And KM.SupplierID = KD.SupplierID	 " & vbCrLf & _
                          " 		Left Join MS_PartMapping PM ON KD.PartNo = PM.PartNo And KD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 		Where Year(KM.KanbanDate) = '" & Format(dtPOPeriod.Value, "yyyy") & "' And Month(KM.KanbanDate) = '" & Format(dtPOPeriod.Value, "MM") & "' " & vbCrLf

        ls_SQL = ls_SQL + " 	) PO ON FD.PartNo = PO.PartNo And FD.AffiliateID = PO.AffiliateID " & vbCrLf & _
                          " ) As s " & vbCrLf & _
                          " PIVOT " & vbCrLf & _
                          " ( " & vbCrLf & _
                          " 	SUM(KanbanQty) " & vbCrLf & _
                          " 	FOR KanbanDate In ([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26],[27],[28],[29],[30],[31]) " & vbCrLf & _
                          " )AS pvt " & vbCrLf & _
                          "  " & vbCrLf & _
                          "  " & vbCrLf

        ls_SQL = ls_SQL + " UNION ALL " & vbCrLf & _
                          "  --Forecast Cumm  " & vbCrLf & _
                          "  Select FD0.Period,FD0.Rev,FD.PartNo, PM.SupplierID, FD.AffiliateID, MPQ = PM.MOQ, Data = 'Cumm FC vs PO Affiliate', Total = FD0.Total  " & vbCrLf & _
                          "  ,FD0.F1,FD0.F2,FD0.F3,FD0.F4,FD0.F5,FD0.F6,FD0.F7,FD0.F8,FD0.F9,FD0.F10,FD0.F11,FD0.F12,FD0.F13,FD0.F14,FD0.F15  " & vbCrLf & _
                          "  ,FD0.F16,FD0.F17,FD0.F18,FD0.F19,FD0.F20,FD0.F21,FD0.F22,FD0.F23,FD0.F24,FD0.F25,FD0.F26,FD0.F27,FD0.F28,FD0.F29,FD0.F30,FD0.F31,Seq=6 " & vbCrLf & _
                          "  From (Select PartNo, AffiliateID   " & vbCrLf & _
                          "  	  From ForecastDaily   " & vbCrLf & _
                          "  	  Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01'  " & vbCrLf

        ls_SQL = ls_SQL + ls_filter & vbCrLf

        ls_SQL = ls_SQL + "  	  Group By PartNo, AffiliateID  " & vbCrLf & _
                          "  	  ) FD  " & vbCrLf

        ls_SQL = ls_SQL + "  Left Join MS_PartMapping PM ON FD.PartNo = PM.PartNo And FD.AffiliateID = PM.AffiliateID  " & vbCrLf & _
                          "  Left Join   " & vbCrLf & _
                          "  	( Select Period,Rev='',PartNo, AffiliateID,Total,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31  " & vbCrLf & _
                          "  	  From ForecastDailyDiff " & vbCrLf & _
                          "  	  Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01'  " & vbCrLf & _
                          "  " & vbCrLf & _
                          "  	) FD0 ON FD.PartNo = FD0.PartNo And FD.AffiliateID = FD0.AffiliateID   " & vbCrLf & _
                          "  )  " & vbCrLf & _
                          "  " & vbCrLf


        ls_SQL = ls_SQL + " Select * From IBD " & vbCrLf & _
                          " UNION ALL " & vbCrLf & _
                          " Select Period='',Rev='',a.PartNo,a.SupplierID,a.AffiliateID,a.MPQ,Data='Diff FC vs PO Affiliate',Total='0' " & vbCrLf & _
                          "  ,F1 = Case When ISNULL(e.F1,0) <> 0 Then case when ISNULL(b.F1,0) = 0 then -100 else ((b.F1-e.F1)/e.F1)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F1,0) <> 0 Then case when ISNULL(b.F1,0) = 0 then -100 else ((b.F1-d.F1)/d.F1)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F1,0) <> 0 Then case when ISNULL(b.F1,0) = 0 then -100 else ((b.F1-c.F1)/c.F1)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F1,0) <> 0 Then case when ISNULL(b.F1,0) = 0 then -100 else ((b.F1-a.F1)/a.F1)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F1,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END  " & vbCrLf & _
                          "  ,F2 = Case When ISNULL(e.F2,0) <> 0 Then case when ISNULL(b.F2,0) = 0 then -100 else ((b.F2-e.F2)/e.F2)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F2,0) <> 0 Then case when ISNULL(b.F2,0) = 0 then -100 else ((b.F2-d.F2)/d.F2)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F2,0) <> 0 Then case when ISNULL(b.F2,0) = 0 then -100 else ((b.F2-c.F2)/c.F2)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F2,0) <> 0 Then case when ISNULL(b.F2,0) = 0 then -100 else ((b.F2-a.F2)/a.F2)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F2,0) = 0 then 0 Else 100 end " & vbCrLf

        ls_SQL = ls_SQL + "  	  END " & vbCrLf & _
                          "  ,F3 = Case When ISNULL(e.F3,0) <> 0 Then case when ISNULL(b.F3,0) = 0 then -100 else ((b.F3-e.F3)/e.F3)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F3,0) <> 0 Then case when ISNULL(b.F3,0) = 0 then -100 else ((b.F3-d.F3)/d.F3)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F3,0) <> 0 Then case when ISNULL(b.F3,0) = 0 then -100 else ((b.F3-c.F3)/c.F3)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F3,0) <> 0 Then case when ISNULL(b.F3,0) = 0 then -100 else ((b.F3-a.F3)/a.F3)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F3,0) = 0 then 0 Else 100 end  " & vbCrLf & _
                          "  	  END  " & vbCrLf & _
                          "  ,F4 = Case When ISNULL(e.F4,0) <> 0 Then case when ISNULL(b.F4,0) = 0 then -100 else ((b.F4-e.F4)/e.F4)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F4,0) <> 0 Then case when ISNULL(b.F4,0) = 0 then -100 else ((b.F4-d.F4)/d.F4)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F4,0) <> 0 Then case when ISNULL(b.F4,0) = 0 then -100 else ((b.F4-c.F4)/c.F4)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F4,0) <> 0 Then case when ISNULL(b.F4,0) = 0 then -100 else ((b.F4-a.F4)/a.F4)*100 end  " & vbCrLf

        ls_SQL = ls_SQL + "  		    ELSE case when ISNULL(b.F4,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END  " & vbCrLf & _
                          "  ,F5 = Case When ISNULL(e.F5,0) <> 0 Then case when ISNULL(b.F5,0) = 0 then -100 else ((b.F5-e.F5)/e.F5)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F5,0) <> 0 Then case when ISNULL(b.F5,0) = 0 then -100 else ((b.F5-d.F5)/d.F5)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F5,0) <> 0 Then case when ISNULL(b.F5,0) = 0 then -100 else ((b.F5-c.F5)/c.F5)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F5,0) <> 0 Then case when ISNULL(b.F5,0) = 0 then -100 else ((b.F5-a.F5)/a.F5)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F5,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END  " & vbCrLf & _
                          "  ,F6 = Case When ISNULL(e.F6,0) <> 0 Then case when ISNULL(b.F6,0) = 0 then -100 else ((b.F6-e.F6)/e.F6)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F6,0) <> 0 Then case when ISNULL(b.F6,0) = 0 then -100 else ((b.F6-d.F6)/d.F6)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F6,0) <> 0 Then case when ISNULL(b.F6,0) = 0 then -100 else ((b.F6-c.F6)/c.F6)*100 end   " & vbCrLf

        ls_SQL = ls_SQL + "  		    When ISNULL(a.F6,0) <> 0 Then case when ISNULL(b.F6,0) = 0 then -100 else ((b.F6-a.F6)/a.F6)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F6,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END  " & vbCrLf & _
                          "  ,F7 = Case When ISNULL(e.F7,0) <> 0 Then case when ISNULL(b.F7,0) = 0 then -100 else ((b.F7-e.F7)/e.F7)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F7,0) <> 0 Then case when ISNULL(b.F7,0) = 0 then -100 else ((b.F7-d.F7)/d.F7)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F7,0) <> 0 Then case when ISNULL(b.F7,0) = 0 then -100 else ((b.F7-c.F7)/c.F7)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F7,0) <> 0 Then case when ISNULL(b.F7,0) = 0 then -100 else ((b.F7-a.F7)/a.F7)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F7,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END  " & vbCrLf & _
                          "  ,F8 = Case When ISNULL(e.F8,0) <> 0 Then case when ISNULL(b.F8,0) = 0 then -100 else ((b.F8-e.F8)/e.F8)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F8,0) <> 0 Then case when ISNULL(b.F8,0) = 0 then -100 else ((b.F8-d.F8)/d.F8)*100 end   " & vbCrLf

        ls_SQL = ls_SQL + "  		    When ISNULL(c.F8,0) <> 0 Then case when ISNULL(b.F8,0) = 0 then -100 else ((b.F8-c.F8)/c.F8)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F8,0) <> 0 Then case when ISNULL(b.F8,0) = 0 then -100 else ((b.F8-a.F8)/a.F8)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F8,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          " 	  END " & vbCrLf & _
                          "  ,F9 = Case When ISNULL(e.F9,0) <> 0 Then case when ISNULL(b.F9,0) = 0 then -100 else ((b.F9-e.F9)/e.F9)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F9,0) <> 0 Then case when ISNULL(b.F9,0) = 0 then -100 else ((b.F9-d.F9)/d.F9)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F9,0) <> 0 Then case when ISNULL(b.F9,0) = 0 then -100 else ((b.F9-c.F9)/c.F9)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F9,0) <> 0 Then case when ISNULL(b.F9,0) = 0 then -100 else ((b.F9-a.F9)/a.F9)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F9,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F10 = Case When ISNULL(e.F10,0) <> 0 Then case when ISNULL(b.F10,0) = 0 then -100 else ((b.F10-e.F10)/e.F10)*100 end   " & vbCrLf

        ls_SQL = ls_SQL + "  		    When ISNULL(d.F10,0) <> 0 Then case when ISNULL(b.F10,0) = 0 then -100 else ((b.F10-d.F10)/d.F10)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F10,0) <> 0 Then case when ISNULL(b.F10,0) = 0 then -100 else ((b.F10-c.F10)/c.F10)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F10,0) <> 0 Then case when ISNULL(b.F10,0) = 0 then -100 else ((b.F10-a.F10)/a.F10)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F10,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F11 = Case When ISNULL(e.F11,0) <> 0 Then case when ISNULL(b.F11,0) = 0 then -100 else ((b.F11-e.F11)/e.F11)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F11,0) <> 0 Then case when ISNULL(b.F11,0) = 0 then -100 else ((b.F11-d.F11)/d.F11)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F11,0) <> 0 Then case when ISNULL(b.F11,0) = 0 then -100 else ((b.F11-c.F11)/c.F11)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F11,0) <> 0 Then case when ISNULL(b.F11,0) = 0 then -100 else ((b.F11-a.F11)/a.F11)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F11,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf

        ls_SQL = ls_SQL + "  ,F12 = Case When ISNULL(e.F12,0) <> 0 Then case when ISNULL(b.F12,0) = 0 then -100 else ((b.F12-e.F12)/e.F12)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F12,0) <> 0 Then case when ISNULL(b.F12,0) = 0 then -100 else ((b.F12-d.F12)/d.F12)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F12,0) <> 0 Then case when ISNULL(b.F12,0) = 0 then -100 else ((b.F12-c.F12)/c.F12)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F12,0) <> 0 Then case when ISNULL(b.F12,0) = 0 then -100 else ((b.F12-a.F12)/a.F12)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F12,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F13 = Case When ISNULL(e.F13,0) <> 0 Then case when ISNULL(b.F13,0) = 0 then -100 else ((b.F13-e.F13)/e.F13)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F13,0) <> 0 Then case when ISNULL(b.F13,0) = 0 then -100 else ((b.F13-d.F13)/d.F13)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F13,0) <> 0 Then case when ISNULL(b.F13,0) = 0 then -100 else ((b.F13-c.F13)/c.F13)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F13,0) <> 0 Then case when ISNULL(b.F13,0) = 0 then -100 else ((b.F13-a.F13)/a.F13)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F13,0) = 0 then 0 Else 100 end " & vbCrLf

        ls_SQL = ls_SQL + "  	  END " & vbCrLf & _
                          "  ,F14 = Case When ISNULL(e.F14,0) <> 0 Then case when ISNULL(b.F14,0) = 0 then -100 else ((b.F14-e.F14)/e.F14)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F14,0) <> 0 Then case when ISNULL(b.F14,0) = 0 then -100 else ((b.F14-d.F14)/d.F14)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F14,0) <> 0 Then case when ISNULL(b.F14,0) = 0 then -100 else ((b.F14-c.F14)/c.F14)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F14,0) <> 0 Then case when ISNULL(b.F14,0) = 0 then -100 else ((b.F14-a.F14)/a.F14)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F14,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F15 = Case When ISNULL(e.F15,0) <> 0 Then case when ISNULL(b.F15,0) = 0 then -100 else ((b.F15-e.F15)/e.F15)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F15,0) <> 0 Then case when ISNULL(b.F15,0) = 0 then -100 else ((b.F15-d.F15)/d.F15)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F15,0) <> 0 Then case when ISNULL(b.F15,0) = 0 then -100 else ((b.F15-c.F15)/c.F15)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F15,0) <> 0 Then case when ISNULL(b.F15,0) = 0 then -100 else ((b.F15-a.F15)/a.F15)*100 end  " & vbCrLf

        ls_SQL = ls_SQL + "  		    ELSE case when ISNULL(b.F15,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F16 = Case When ISNULL(e.F16,0) <> 0 Then case when ISNULL(b.F16,0) = 0 then -100 else ((b.F16-e.F16)/e.F16)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F16,0) <> 0 Then case when ISNULL(b.F16,0) = 0 then -100 else ((b.F16-d.F16)/d.F16)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F16,0) <> 0 Then case when ISNULL(b.F16,0) = 0 then -100 else ((b.F16-c.F16)/c.F16)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F16,0) <> 0 Then case when ISNULL(b.F16,0) = 0 then -100 else ((b.F16-a.F16)/a.F16)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F16,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F17 = Case When ISNULL(e.F17,0) <> 0 Then case when ISNULL(b.F17,0) = 0 then -100 else ((b.F17-e.F17)/e.F17)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F17,0) <> 0 Then case when ISNULL(b.F17,0) = 0 then -100 else ((b.F17-d.F17)/d.F17)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F17,0) <> 0 Then case when ISNULL(b.F17,0) = 0 then -100 else ((b.F17-c.F17)/c.F17)*100 end   " & vbCrLf

        ls_SQL = ls_SQL + "  		    When ISNULL(a.F17,0) <> 0 Then case when ISNULL(b.F17,0) = 0 then -100 else ((b.F17-a.F17)/a.F17)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F17,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F18 = Case When ISNULL(e.F18,0) <> 0 Then case when ISNULL(b.F18,0) = 0 then -100 else ((b.F18-e.F18)/e.F18)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F18,0) <> 0 Then case when ISNULL(b.F18,0) = 0 then -100 else ((b.F18-d.F18)/d.F18)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F18,0) <> 0 Then case when ISNULL(b.F18,0) = 0 then -100 else ((b.F18-c.F18)/c.F18)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F18,0) <> 0 Then case when ISNULL(b.F18,0) = 0 then -100 else ((b.F18-a.F18)/a.F18)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F18,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F19 = Case When ISNULL(e.F19,0) <> 0 Then case when ISNULL(b.F19,0) = 0 then -100 else ((b.F19-e.F19)/e.F19)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F19,0) <> 0 Then case when ISNULL(b.F19,0) = 0 then -100 else ((b.F19-d.F19)/d.F19)*100 end   " & vbCrLf

        ls_SQL = ls_SQL + "  		    When ISNULL(c.F19,0) <> 0 Then case when ISNULL(b.F19,0) = 0 then -100 else ((b.F19-c.F19)/c.F19)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F19,0) <> 0 Then case when ISNULL(b.F19,0) = 0 then -100 else ((b.F19-a.F19)/a.F19)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F19,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F20 = Case When ISNULL(e.F20,0) <> 0 Then case when ISNULL(b.F20,0) = 0 then -100 else ((b.F20-e.F20)/e.F20)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F20,0) <> 0 Then case when ISNULL(b.F20,0) = 0 then -100 else ((b.F20-d.F20)/d.F20)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F20,0) <> 0 Then case when ISNULL(b.F20,0) = 0 then -100 else ((b.F20-c.F20)/c.F20)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F20,0) <> 0 Then case when ISNULL(b.F20,0) = 0 then -100 else ((b.F20-a.F20)/a.F20)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F20,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F21 = Case When ISNULL(e.F21,0) <> 0 Then case when ISNULL(b.F21,0) = 0 then -100 else ((b.F21-e.F21)/e.F21)*100 end   " & vbCrLf

        ls_SQL = ls_SQL + "  		    When ISNULL(d.F21,0) <> 0 Then case when ISNULL(b.F21,0) = 0 then -100 else ((b.F21-d.F21)/d.F21)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F21,0) <> 0 Then case when ISNULL(b.F21,0) = 0 then -100 else ((b.F21-c.F21)/c.F21)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F21,0) <> 0 Then case when ISNULL(b.F21,0) = 0 then -100 else ((b.F21-a.F21)/a.F21)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F21,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END  " & vbCrLf & _
                          "  ,F22 = Case When ISNULL(e.F22,0) <> 0 Then case when ISNULL(b.F22,0) = 0 then -100 else ((b.F22-e.F22)/e.F22)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F22,0) <> 0 Then case when ISNULL(b.F22,0) = 0 then -100 else ((b.F22-d.F22)/d.F22)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F22,0) <> 0 Then case when ISNULL(b.F22,0) = 0 then -100 else ((b.F22-c.F22)/c.F22)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F22,0) <> 0 Then case when ISNULL(b.F22,0) = 0 then -100 else ((b.F22-a.F22)/a.F22)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F22,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf

        ls_SQL = ls_SQL + "  ,F23 = Case When ISNULL(e.F23,0) <> 0 Then case when ISNULL(b.F23,0) = 0 then -100 else ((b.F23-e.F23)/e.F23)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F23,0) <> 0 Then case when ISNULL(b.F23,0) = 0 then -100 else ((b.F23-d.F23)/d.F23)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F23,0) <> 0 Then case when ISNULL(b.F23,0) = 0 then -100 else ((b.F23-c.F23)/c.F23)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F23,0) <> 0 Then case when ISNULL(b.F23,0) = 0 then -100 else ((b.F23-a.F23)/a.F23)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F23,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F24 = Case When ISNULL(e.F24,0) <> 0 Then case when ISNULL(b.F24,0) = 0 then -100 else ((b.F24-e.F24)/e.F24)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F24,0) <> 0 Then case when ISNULL(b.F24,0) = 0 then -100 else ((b.F24-d.F24)/d.F24)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F24,0) <> 0 Then case when ISNULL(b.F24,0) = 0 then -100 else ((b.F24-c.F24)/c.F24)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F24,0) <> 0 Then case when ISNULL(b.F24,0) = 0 then -100 else ((b.F24-a.F24)/a.F24)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F24,0) = 0 then 0 Else 100 end " & vbCrLf

        ls_SQL = ls_SQL + "  	  END " & vbCrLf & _
                          "  ,F25 = Case When ISNULL(e.F25,0) <> 0 Then case when ISNULL(b.F25,0) = 0 then -100 else ((b.F25-e.F25)/e.F25)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F25,0) <> 0 Then case when ISNULL(b.F25,0) = 0 then -100 else ((b.F25-d.F25)/d.F25)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F25,0) <> 0 Then case when ISNULL(b.F25,0) = 0 then -100 else ((b.F25-c.F25)/c.F25)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F25,0) <> 0 Then case when ISNULL(b.F25,0) = 0 then -100 else ((b.F25-a.F25)/a.F25)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F25,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F26 = Case When ISNULL(e.F26,0) <> 0 Then case when ISNULL(b.F26,0) = 0 then -100 else ((b.F26-e.F26)/e.F26)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F26,0) <> 0 Then case when ISNULL(b.F26,0) = 0 then -100 else ((b.F26-d.F26)/d.F26)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F26,0) <> 0 Then case when ISNULL(b.F26,0) = 0 then -100 else ((b.F26-c.F26)/c.F26)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F26,0) <> 0 Then case when ISNULL(b.F26,0) = 0 then -100 else ((b.F26-a.F26)/a.F26)*100 end  " & vbCrLf

        ls_SQL = ls_SQL + "  		    ELSE case when ISNULL(b.F26,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F27 = Case When ISNULL(e.F27,0) <> 0 Then case when ISNULL(b.F27,0) = 0 then -100 else ((b.F27-e.F27)/e.F27)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F27,0) <> 0 Then case when ISNULL(b.F27,0) = 0 then -100 else ((b.F27-d.F27)/d.F27)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F27,0) <> 0 Then case when ISNULL(b.F27,0) = 0 then -100 else ((b.F27-c.F27)/c.F27)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F27,0) <> 0 Then case when ISNULL(b.F27,0) = 0 then -100 else ((b.F27-a.F27)/a.F27)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F27,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F28 = Case When ISNULL(e.F28,0) <> 0 Then case when ISNULL(b.F28,0) = 0 then -100 else ((b.F28-e.F28)/e.F28)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F28,0) <> 0 Then case when ISNULL(b.F28,0) = 0 then -100 else ((b.F28-d.F28)/d.F28)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F28,0) <> 0 Then case when ISNULL(b.F28,0) = 0 then -100 else ((b.F28-c.F28)/c.F28)*100 end   " & vbCrLf

        ls_SQL = ls_SQL + "  		    When ISNULL(a.F28,0) <> 0 Then case when ISNULL(b.F28,0) = 0 then -100 else ((b.F28-a.F28)/a.F28)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F28,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F29 = Case When ISNULL(e.F29,0) <> 0 Then case when ISNULL(b.F29,0) = 0 then -100 else ((b.F29-e.F29)/e.F29)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F29,0) <> 0 Then case when ISNULL(b.F29,0) = 0 then -100 else ((b.F29-d.F29)/d.F29)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F29,0) <> 0 Then case when ISNULL(b.F29,0) = 0 then -100 else ((b.F29-c.F29)/c.F29)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F29,0) <> 0 Then case when ISNULL(b.F29,0) = 0 then -100 else ((b.F29-a.F29)/a.F29)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F29,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F30 = Case When ISNULL(e.F30,0) <> 0 Then case when ISNULL(b.F30,0) = 0 then -100 else ((b.F30-e.F30)/e.F30)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F30,0) <> 0 Then case when ISNULL(b.F30,0) = 0 then -100 else ((b.F30-d.F30)/d.F30)*100 end   " & vbCrLf

        ls_SQL = ls_SQL + "  		    When ISNULL(c.F30,0) <> 0 Then case when ISNULL(b.F30,0) = 0 then -100 else ((b.F30-c.F30)/c.F30)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F30,0) <> 0 Then case when ISNULL(b.F30,0) = 0 then -100 else ((b.F30-b.F30)/a.F30)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F30,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  ,F31 = Case When ISNULL(e.F31,0) <> 0 Then case when ISNULL(b.F31,0) = 0 then -100 else ((b.F31-e.F31)/e.F31)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(d.F31,0) <> 0 Then case when ISNULL(b.F31,0) = 0 then -100 else ((b.F31-d.F31)/d.F31)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(c.F31,0) <> 0 Then case when ISNULL(b.F31,0) = 0 then -100 else ((b.F31-c.F31)/c.F31)*100 end   " & vbCrLf & _
                          "  		    When ISNULL(a.F31,0) <> 0 Then case when ISNULL(b.F31,0) = 0 then -100 else ((b.F31-a.F31)/a.F31)*100 end  " & vbCrLf & _
                          "  		    ELSE case when ISNULL(b.F31,0) = 0 then 0 Else 100 end " & vbCrLf & _
                          "  	  END " & vbCrLf & _
                          "  "


        ls_SQL = ls_SQL + "  " & vbCrLf & _
                          "  " & vbCrLf & _
                          " ,Seq=5 " & vbCrLf & _
                          "  " & vbCrLf & _
                          " From IBD a " & vbCrLf & _
                          " Left Join IBD b On a.PartNo = b.PartNo And a.AffiliateID = b.AffiliateID And b.Data = 'PO Affiliate' " & vbCrLf & _
                          " Left Join IBD c On a.PartNo = c.PartNo And a.AffiliateID = c.AffiliateID And c.Data = 'Forecast rev01' " & vbCrLf & _
                          " Left Join IBD d On a.PartNo = d.PartNo And a.AffiliateID = d.AffiliateID And d.Data = 'Forecast rev02' " & vbCrLf & _
                          " Left Join IBD e On a.PartNo = e.PartNo And a.AffiliateID = e.AffiliateID And e.Data = 'Forecast rev03' " & vbCrLf & _
                          " Where a.Data = 'Forecast rev00' " & vbCrLf & _
                          " --UNION ALL " & vbCrLf

        ls_SQL = ls_SQL + " --Select Period='',Rev='',PartNo, SupplierID, AffiliateID, MPQ, Data = 'Cumm FC vs PO Affiliate', Total = 0 " & vbCrLf & _
                          " --,F1=0,F2=0,F3=0,F4=0,F5=0,F6=0,F7=0,F8=0,F9=0,F10=0,F11=0,F12=0,F13=0,F14=0,F15=0 " & vbCrLf & _
                          " --,F16=0,F17=0,F18=0,F19=0,F20=0,F21=0,F22=0,F23=0,F24=0,F25=0,F26=0,F27=0,F28=0,F29=0,F30=0,F31=0,Seq=6 " & vbCrLf & _
                          " --From IBD Where Data = 'Forecast rev00' " & vbCrLf & _
                          "  " & vbCrLf & _
                          " Order By PartNo,AffiliateID,Seq " & vbCrLf

        Return ls_SQL

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
                    FileName = "TemplateForecastReportRegular.xlsx"
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

    Private Sub grid_CustomColumnDisplayText(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewColumnDisplayTextEventArgs) Handles grid.CustomColumnDisplayText
        With e.Column
            For i = 1 To 31
                If .FieldName = "F" & i Then
                    If e.GetFieldValue("Data") = "Cumm FC vs PO Affiliate" Then
                        e.DisplayText = CInt(e.GetFieldValue("F" & i)) & " %"
                    End If
                    If e.GetFieldValue("Data") = "Diff FC vs PO Affiliate" Then
                        e.DisplayText = CInt(e.GetFieldValue("F" & i)) & " %"
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

End Class