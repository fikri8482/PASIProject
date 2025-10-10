Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports OfficeOpenXml
Imports System.IO
'Menu ID : C01
'Menu Desc : PO LIST
Public Class POList
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "PO LIST"
            If Session("C01Url") <> "" Then
                Call bindData()
                Session.Remove("C01Url")
            End If
            dtPeriodFrom.Value = Now
            dtPeriodTo.Value = Now
            rdrAff1.Checked = True
            rdrCom1.Checked = True
            lblInfo.Text = ""
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("C01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub btnADD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnADD.Click
        Session("C01Url") = "~/PurchaseOrder/POList.aspx"
        Response.Redirect("~/PurchaseOrder/POEntry.aspx")
    End Sub

    Private Sub btnEDI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEDI.Click
        Response.Redirect("~/PurchaseOrder/POEntryEDI.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            grid.JSProperties("cpMessage") = Session("C01Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                Case "downloadSummary"
                    If txtPONo.Text = "" Then
                        Call clsMsg.DisplayMessage(lblInfo, "8001", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Exit Select
                    End If
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = clsPO.GetTable(txtPONo.Text, Session("AffiliateID"))
                    FileName = "TemplatePO.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:13", psERR)
                    End If
            End Select

EndProcedure:
            Session("C01Msg") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim pDateFrom As Date
        Dim pDateTo As Date

        pDateFrom = Format(dtPeriodFrom.Value, "yyyy-MM-01")
        pDateTo = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPeriodTo.Value, "yyyy-MM-01"))))

        If txtPONo.Text.Trim <> "" Then
            pWhere = pWhere + " and a.PONo like '%" & txtPONo.Text.Trim & "%' "
        End If

        If rdrAff2.Checked = True Then
            pWhere = pWhere + " and a.AffiliateApproveDate is not null "
        End If

        If rdrAff3.Checked = True Then
            pWhere = pWhere + " and a.AffiliateApproveDate is null "
        End If

        If rdrCom2.Checked = True Then
            pWhere = pWhere + " and a.CommercialCls = '1' "
        End If

        If rdrCom3.Checked = True Then
            pWhere = pWhere + " and a.CommercialCls = '0' "
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  distinct " & vbCrLf & _
                  " 	'DETAIL' DetailPage, " & vbCrLf & _
                  " 	'REVISE' RevisePage, 0 AllowAccess, " & vbCrLf & _
                  " 	Period, " & vbCrLf & _
                  " 	RTRIM(a.PONo) PONo, " & vbCrLf & _
                  " 	CASE WHEN b.KanbanCls = 0 then RTRIM(a.PONo) + '-' + RTRIM(a.SupplierID) ELSE a.PONo END POMarking, " & vbCrLf & _
                  " 	case CommercialCls when '0' then 'NO' else 'YES' end CommercialCls, " & vbCrLf & _
                  " 	RTRIM(ShipCls) ShipCls, " & vbCrLf & _
                  " 	a.EntryDate,  " & vbCrLf & _
                  " 	a.EntryUser, "

            ls_SQL = ls_SQL + " 	case ISNULL(a.EntryDate,0) when 0 then 0 else 1 end POStatus1, " & vbCrLf & _
                              " 	case ISNULL(AffiliateApproveDate,0) when 0 then 0 else 1 end POStatus2, " & vbCrLf & _
                              " 	case ISNULL(PASISendAffiliateDate,0) when 0 then 0 else 1 end POStatus3, " & vbCrLf & _
                              " 	case ISNULL(SupplierApproveDate,0) when 0 then 0 else 1 end POStatus4, " & vbCrLf & _
                              " 	case ISNULL(SupplierApprovePendingDate,0) when 0 then 0 else 1 end POStatus5, " & vbCrLf & _
                              " 	case ISNULL(SupplierUnApproveDate,0) when 0 then 0 else 1 end POStatus6, " & vbCrLf & _
                              " 	case ISNULL(PASIApproveDate,0) when 0 then 0 else 1 end POStatus7, " & vbCrLf & _
                              " 	case ISNULL(FinalApproveDate,0) when 0 then 0 else 1 end POStatus8, RTRIM(a.SupplierID) SupplierID " & vbCrLf & _
                              " from po_master a " & vbCrLf & _
                              " inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID" & vbCrLf

            ls_SQL = ls_SQL + " where a.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                              " and Period between '" & Format(pDateFrom, "yyyy-MM-dd") & "' and '" & Format(pDateTo, "yyyy-MM-dd") & "' " & pWhere & ""


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' DetailPage, '' AllowAccess,'' as RevisePage, '' Period, ''PONo, ''POMarking, ''CommercialCls, ''ShipCls, ''CurrAff, ''AmountAff, '' EntryDate, ''EntryUser, '' POStatus1, ''POStatus2, ''POStatus3, ''POStatus4, ''POStatus5, ''POStatus6, ''POStatus7, ''POStatus8"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()

            End With

            sqlConn.Close()
        End Using
    End Sub

    Protected Function GetRowValue(ByVal container As GridViewDataItemTemplateContainer) As String
        If Not IsNothing(container.KeyValue) Then
            Return container.Grid.GetRowValuesByKeyValue(container.KeyValue, "PONo").ToString()
        End If
    End Function

    Protected Function GetAffiliateID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateID = container.Grid.GetRowValues(container.ItemIndex, "ShipCls")
    End Function

    Protected Function GetAffiliateName(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateName = container.Grid.GetRowValues(container.ItemIndex, "CommercialCls")
    End Function

    Protected Function GetPeriod(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPeriod = container.Grid.GetRowValues(container.ItemIndex, "Period")
    End Function

    Protected Function GetSupplierID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetSupplierID = container.Grid.GetRowValues(container.ItemIndex, "SupplierID")
    End Function

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "TemplatePO " & pData.Rows(0)(1).ToString.Trim & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\PurchaseOrder\Import\" & tempFile & "")
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
                .Cells(3, 3).Value = Format(pData.Rows(0)(0), "MMM-yy")
                .Cells(4, 3).Value = pData.Rows(0)(2)
                .Cells(5, 3).Value = pData.Rows(0)(3)
                .Cells(6, 3).Value = pData.Rows(0)(1)
                .Cells(7, 3).Value = pData.Rows(0)(4)
                .Cells(8, 3).Value = "Approved By PASI"

                For irow = 0 To pData.Rows.Count - 1
                    For icol = 1 To pData.Columns.Count - 5
                        If icol = 1 Then
                            .Cells(irow + rowstart + 1, icol).Value = irow + 1
                        End If
                        If icol + 1 < 13 Then
                            .Cells(irow + rowstart + 1, icol + 1).Value = pData.Rows(irow)(icol + 4)
                        Else
                            .Cells(irow + rowstart + 1, icol + 2).Value = pData.Rows(irow)(icol + 4)
                        End If

                    Next
                Next

                ''ALIGNMENT
                ''.Cells(rowstart + 1, icol, irow, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                .Cells(14, 3, irow + rowstart, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colSupplierName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(iRow + space, colKanbanSeqNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colSupplierDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPartName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIInvQty).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIInvCurr).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left

                'Dim rgAll As ExcelRange = .Cells('.Cells(Space() - 2, colNo, grid.VisibleRowCount + (Space() - 1), colCount - 1)
                Dim rgAll As ExcelRange = .Cells(14, 1, 13 + irow, 44)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\PurchaseOrder\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

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
#End Region
End Class