Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions

Public Class PORevHistory
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "D05"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String, pub_Remarks As String, pub_Revision As String
    Dim pub_FinalApproval As String, pub_DeliveyBy As String
    Dim pub_Period As Date
    Dim pub_HeijunkaSttus As Boolean
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then

            Session("Mode") = "New"
            dtPeriodFrom.Value = Now
            up_FillCombo(dtPeriodFrom.Value)
            up_FillComboPart()
            up_FillComboAffiliateID()
            dtPeriodFrom.Focus()

            lblInfo.Text = ""

        ElseIf IsCallback Then
            If grid.VisibleRowCount = 0 Then Exit Sub
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        grid.JSProperties("cpMessage") = ""
        Call bindData()
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()
                    'Call bindPOStatus()

                    grid.JSProperties("cpSearch") = "search"

                    'Dim TempASPxGridViewCellMerger As ASPxGridViewCellMerger = New ASPxGridViewCellMerger(grid, "NoUrut,PartNo,PartName,KanbanCls,UnitDesc,MOQ,QtyBox,Maker")
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                    grid.JSProperties("cpSearch") = ""
            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        grid.CollapseAll()

        cboPONo.Text = ""
        cboPartNo.Text = ""

        dtPeriodFrom.Value = Now

        up_FillCombo(dtPeriodFrom.Value)

        cboPartNo.Items.Clear()

        lblInfo.Text = ""

    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        If x > grid.VisibleRowCount Then Exit Sub

        With grid
            If .VisibleRowCount > 0 Then
                If e.GetValue("Header") = "1" Then
                    e.Cell.BackColor = Color.Aquamarine
                    If e.DataColumn.FieldName = "POQty" _
                        Or e.DataColumn.FieldName = "DeliveryD1" Or e.DataColumn.FieldName = "DeliveryD2" Or e.DataColumn.FieldName = "DeliveryD3" _
                        Or e.DataColumn.FieldName = "DeliveryD4" Or e.DataColumn.FieldName = "DeliveryD5" Or e.DataColumn.FieldName = "DeliveryD6" _
                        Or e.DataColumn.FieldName = "DeliveryD7" Or e.DataColumn.FieldName = "DeliveryD8" Or e.DataColumn.FieldName = "DeliveryD9" _
                        Or e.DataColumn.FieldName = "DeliveryD10" Or e.DataColumn.FieldName = "DeliveryD11" Or e.DataColumn.FieldName = "DeliveryD12" _
                        Or e.DataColumn.FieldName = "DeliveryD13" Or e.DataColumn.FieldName = "DeliveryD14" Or e.DataColumn.FieldName = "DeliveryD15" _
                        Or e.DataColumn.FieldName = "DeliveryD16" Or e.DataColumn.FieldName = "DeliveryD17" Or e.DataColumn.FieldName = "DeliveryD18" _
                        Or e.DataColumn.FieldName = "DeliveryD19" Or e.DataColumn.FieldName = "DeliveryD20" Or e.DataColumn.FieldName = "DeliveryD21" _
                        Or e.DataColumn.FieldName = "DeliveryD22" Or e.DataColumn.FieldName = "DeliveryD23" Or e.DataColumn.FieldName = "DeliveryD24" _
                        Or e.DataColumn.FieldName = "DeliveryD25" Or e.DataColumn.FieldName = "DeliveryD26" Or e.DataColumn.FieldName = "DeliveryD27" _
                        Or e.DataColumn.FieldName = "DeliveryD28" Or e.DataColumn.FieldName = "DeliveryD29" Or e.DataColumn.FieldName = "DeliveryD30" _
                        Or e.DataColumn.FieldName = "DeliveryD31" Then
                        e.Cell.Text = ""
                    End If
                End If
                If e.GetValue("Header") = "2" Then
                    If e.DataColumn.FieldName = "NoUrut" Or e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" Then
                        e.Cell.Text = ""
                    End If
                End If
                If e.GetValue("AffiliateName") = "DIFFERENCE" Then
                    If e.DataColumn.FieldName = "AffiliateName" Or e.DataColumn.FieldName = "PONo" _
                        Or e.DataColumn.FieldName = "POQty" _
                        Or e.DataColumn.FieldName = "DeliveryD1" Or e.DataColumn.FieldName = "DeliveryD2" Or e.DataColumn.FieldName = "DeliveryD3" _
                        Or e.DataColumn.FieldName = "DeliveryD4" Or e.DataColumn.FieldName = "DeliveryD5" Or e.DataColumn.FieldName = "DeliveryD6" _
                        Or e.DataColumn.FieldName = "DeliveryD7" Or e.DataColumn.FieldName = "DeliveryD8" Or e.DataColumn.FieldName = "DeliveryD9" _
                        Or e.DataColumn.FieldName = "DeliveryD10" Or e.DataColumn.FieldName = "DeliveryD11" Or e.DataColumn.FieldName = "DeliveryD12" _
                        Or e.DataColumn.FieldName = "DeliveryD13" Or e.DataColumn.FieldName = "DeliveryD14" Or e.DataColumn.FieldName = "DeliveryD15" _
                        Or e.DataColumn.FieldName = "DeliveryD16" Or e.DataColumn.FieldName = "DeliveryD17" Or e.DataColumn.FieldName = "DeliveryD18" _
                        Or e.DataColumn.FieldName = "DeliveryD19" Or e.DataColumn.FieldName = "DeliveryD20" Or e.DataColumn.FieldName = "DeliveryD21" _
                        Or e.DataColumn.FieldName = "DeliveryD22" Or e.DataColumn.FieldName = "DeliveryD23" Or e.DataColumn.FieldName = "DeliveryD24" _
                        Or e.DataColumn.FieldName = "DeliveryD25" Or e.DataColumn.FieldName = "DeliveryD26" Or e.DataColumn.FieldName = "DeliveryD27" _
                        Or e.DataColumn.FieldName = "DeliveryD28" Or e.DataColumn.FieldName = "DeliveryD29" Or e.DataColumn.FieldName = "DeliveryD30" _
                        Or e.DataColumn.FieldName = "DeliveryD31" Then
                        'e.Cell.Text = ""
                        e.Cell.BackColor = Color.Gray
                    End If
                    'If e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" _
                    '    Or e.DataColumn.FieldName = "UnitDesc" Or e.DataColumn.FieldName = "PartNo" _
                    '    Or e.DataColumn.FieldName = "PartName" Or e.DataColumn.FieldName = "Maker" Or e.DataColumn.FieldName = "NoUrut" Then
                    '    e.Cell.Text = ""
                    'End If
                End If

                'If e.GetValue("AffiliateName") <> "BY AFFILIATE" And e.GetValue("AffiliateName") <> "DIFFERENCE" Then
                '    If e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" _
                '        Or e.DataColumn.FieldName = "UnitDesc" Or e.DataColumn.FieldName = "PartNo" _
                '        Or e.DataColumn.FieldName = "PartName" Or e.DataColumn.FieldName = "Maker" Or e.DataColumn.FieldName = "NoUrut" Then
                '        e.Cell.Text = ""
                '    End If
                '    If CDbl(e.GetValue("POQty")) <> CDbl(e.GetValue("POQtyOld")) Then
                '        If e.DataColumn.FieldName = "POQty" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD1")) <> CDbl(e.GetValue("DeliveryD1Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD1" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD2")) <> CDbl(e.GetValue("DeliveryD2Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD2" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD3")) <> CDbl(e.GetValue("DeliveryD3Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD3" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD4")) <> CDbl(e.GetValue("DeliveryD4Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD4" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD5")) <> CDbl(e.GetValue("DeliveryD5Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD5" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD6")) <> CDbl(e.GetValue("DeliveryD6Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD6" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD7")) <> CDbl(e.GetValue("DeliveryD7Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD7" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD8")) <> CDbl(e.GetValue("DeliveryD8Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD8" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD9")) <> CDbl(e.GetValue("DeliveryD9Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD9" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD10")) <> CDbl(e.GetValue("DeliveryD10Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD10" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD11")) <> CDbl(e.GetValue("DeliveryD11Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD11" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD12")) <> CDbl(e.GetValue("DeliveryD12Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD12" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD13")) <> CDbl(e.GetValue("DeliveryD13Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD13" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD14")) <> CDbl(e.GetValue("DeliveryD14Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD14" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD15")) <> CDbl(e.GetValue("DeliveryD15Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD15" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD16")) <> CDbl(e.GetValue("DeliveryD16Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD16" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD17")) <> CDbl(e.GetValue("DeliveryD17Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD17" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD18")) <> CDbl(e.GetValue("DeliveryD18Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD18" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD19")) <> CDbl(e.GetValue("DeliveryD19Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD19" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD20")) <> CDbl(e.GetValue("DeliveryD20Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD20" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD21")) <> CDbl(e.GetValue("DeliveryD21Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD21" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD22")) <> CDbl(e.GetValue("DeliveryD22Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD22" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD23")) <> CDbl(e.GetValue("DeliveryD23Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD23" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD24")) <> CDbl(e.GetValue("DeliveryD24Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD24" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD25")) <> CDbl(e.GetValue("DeliveryD25Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD25" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD26")) <> CDbl(e.GetValue("DeliveryD26Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD26" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD27")) <> CDbl(e.GetValue("DeliveryD27Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD27" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD28")) <> CDbl(e.GetValue("DeliveryD28Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD28" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD29")) <> CDbl(e.GetValue("DeliveryD29Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD29" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD30")) <> CDbl(e.GetValue("DeliveryD30Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD30" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                '    If CDbl(e.GetValue("DeliveryD31")) <> CDbl(e.GetValue("DeliveryD31Old")) Then
                '        If e.DataColumn.FieldName = "DeliveryD31" Then
                '            e.Cell.BackColor = Color.Yellow
                '        End If
                '    End If
                'End If
            End If
        End With
    End Sub

    Private Sub cboPONo_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboPONo.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Dim pPeriod As String = Mid(pAction, 12, 4) + "-" + clsGlobal.uf_GetShortMonth(Mid(pAction, 5, 3)) + "-" + "01"
        up_FillCombo(pPeriod)
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If cboPartNo.Text <> clsGlobal.gs_All Then
            pWhere = pWhere + " and PartNo2 = '" & cboPartNo.Text & "'"
        End If

        If cboPONo.Text <> clsGlobal.gs_All Then
            pWhere = pWhere + " and PONo = '" & cboPONo.Text & "'"
        End If

        If cboAffiliateID.Text <> clsGlobal.gs_All Then
            pWhere = pWhere + " and AffiliateID = '" & cboAffiliateID.Text & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "select * from (" & vbCrLf & _
                  " select '1' Header ,'1' NoUrutDesc, row_number() over (order by PartNo asc) as NoUrut, *, " & vbCrLf & _
                  " 			'' AffiliateName , '' PONo,  " & vbCrLf & _
                  "   			0 POQty, 0 POQtyOld, '' CurrDesc, 0 Price, 0 Amount,  " & vbCrLf & _
                  "   			0 DeliveryD1, 0 DeliveryD2, 0 DeliveryD3, 0 DeliveryD4, 0 DeliveryD5,    " & vbCrLf & _
                  "   			0 DeliveryD6, 0 DeliveryD7, 0 DeliveryD8, 0 DeliveryD9, 0 DeliveryD10,    " & vbCrLf & _
                  "   			0 DeliveryD11, 0 DeliveryD12, 0 DeliveryD13, 0 DeliveryD14, 0 DeliveryD15,    " & vbCrLf & _
                  "   			0 DeliveryD16, 0 DeliveryD17,0 DeliveryD18, 0 DeliveryD19, 0 DeliveryD20,    " & vbCrLf & _
                  "   			0 DeliveryD21, 0 DeliveryD22, 0 DeliveryD23, 0 DeliveryD24, 0 DeliveryD25,    " & vbCrLf & _
                  "   			0 DeliveryD26, 0 DeliveryD27, 0 DeliveryD28, 0 DeliveryD29, 0 DeliveryD30,    " & vbCrLf & _
                  "   			0 DeliveryD31,   " & vbCrLf & _
                  "   			0 DeliveryD1Old, 0 DeliveryD2Old, 0 DeliveryD3Old, 0 DeliveryD4Old, 0 DeliveryD5Old,    " & vbCrLf

            ls_SQL = ls_SQL + "   			0 DeliveryD6Old, 0 DeliveryD7Old, 0 DeliveryD8Old, 0 DeliveryD9Old, 0 DeliveryD10Old,   	 " & vbCrLf & _
                              " 			0 DeliveryD11Old, 0 DeliveryD12Old, 0 DeliveryD13Old, 0 DeliveryD14Old, 0 DeliveryD15Old,    " & vbCrLf & _
                              "   			0 DeliveryD16Old, 0 DeliveryD17Old, 0 DeliveryD18Old, 0 DeliveryD19Old, 0 DeliveryD20Old,    " & vbCrLf & _
                              "   			0 DeliveryD21Old, 0 DeliveryD22Old, 0 DeliveryD23Old, 0 DeliveryD24Old, 0 DeliveryD25Old,    " & vbCrLf & _
                              "   			0 DeliveryD26Old, 0 DeliveryD27Old, 0 DeliveryD28Old, 0 DeliveryD29Old, 0 DeliveryD30Old,    " & vbCrLf & _
                              "   			0 DeliveryD31Old " & vbCrLf & _
                              "   			from  " & vbCrLf & _
                              "   		(  " & vbCrLf & _
                              "   			select  distinct 			  " & vbCrLf & _
                              "   				b.PartNo, b.PartNo PartNo2 ,c.PartName, b.AffiliateID, b.AffiliateID AffiliateID2, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker " & vbCrLf & _
                              "   			from PO_Master a   " & vbCrLf

            ls_SQL = ls_SQL + "   			inner join PO_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                              "   			inner join MS_Parts c on b.PartNo = c.PartNo   " & vbCrLf & _
                              "   			inner join MS_UnitCls d on d.UnitCls = c.UnitCls   " & vbCrLf & _
                              "   			where YEAR(a.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(a.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "   		)x " & vbCrLf & _
                              "  UNION ALL" & vbCrLf

            ls_SQL = ls_SQL + "  select '2' Header, tbl1.NoUrutDesc, '' NoUrut, ''PartNo, tbl1.PartNo PartNo2 , ''PartName, '' AffiliateID, tbl1.AffiliateID2, " & vbCrLf & _
                  " 	''UnitDesc, 0 MOQ, 0 QtyBox,''Maker, tbl1.AffiliateName ,isnull(tbl2.PORevNo,tbl2.PONo)PONo, " & vbCrLf & _
                  "  	ISNULL(POQty,0)POQty, ISNULL(POQtyOld,0)POQtyOld, CurrDesc, Price, Amount, " & vbCrLf & _
                  "  	ISNULL(DeliveryD1,0)DeliveryD1, ISNULL(DeliveryD2,0)DeliveryD2, ISNULL(DeliveryD3,0)DeliveryD3, ISNULL(DeliveryD4,0)DeliveryD4, ISNULL(DeliveryD5,0)DeliveryD5,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD6,0)DeliveryD6, ISNULL(DeliveryD7,0)DeliveryD7, ISNULL(DeliveryD8,0)DeliveryD8, ISNULL(DeliveryD9,0)DeliveryD9, ISNULL(DeliveryD10,0)DeliveryD10,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD11,0)DeliveryD11, ISNULL(DeliveryD12,0)DeliveryD12, ISNULL(DeliveryD13,0)DeliveryD13, ISNULL(DeliveryD14,0)DeliveryD14, ISNULL(DeliveryD15,0)DeliveryD15,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD16,0)DeliveryD16, ISNULL(DeliveryD17,0)DeliveryD17, ISNULL(DeliveryD18,0)DeliveryD18, ISNULL(DeliveryD19,0)DeliveryD19, ISNULL(DeliveryD20,0)DeliveryD20,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD21,0)DeliveryD21, ISNULL(DeliveryD22,0)DeliveryD22, ISNULL(DeliveryD23,0)DeliveryD23, ISNULL(DeliveryD24,0)DeliveryD24, ISNULL(DeliveryD25,0)DeliveryD25,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD26,0)DeliveryD26, ISNULL(DeliveryD27,0)DeliveryD27, ISNULL(DeliveryD28,0)DeliveryD28, ISNULL(DeliveryD29,0)DeliveryD29, ISNULL(DeliveryD30,0)DeliveryD30,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD31,0)DeliveryD31,  " & vbCrLf & _
                  "  	ISNULL(DeliveryD1Old,0)DeliveryD1Old, ISNULL(DeliveryD2Old,0)DeliveryD2Old, ISNULL(DeliveryD3Old,0)DeliveryD3Old, ISNULL(DeliveryD4Old,0)DeliveryD4Old, ISNULL(DeliveryD5Old,0)DeliveryD5Old,   " & vbCrLf

            ls_SQL = ls_SQL + "  	ISNULL(DeliveryD6Old,0)DeliveryD6Old, ISNULL(DeliveryD7Old,0)DeliveryD7Old, ISNULL(DeliveryD8Old,0)DeliveryD8Old, ISNULL(DeliveryD9Old,0)DeliveryD9Old, ISNULL(DeliveryD10Old,0)DeliveryD10Old,   	ISNULL(DeliveryD11Old,0)DeliveryD11Old, ISNULL(DeliveryD12Old,0)DeliveryD12Old, ISNULL(DeliveryD13Old,0)DeliveryD13Old, ISNULL(DeliveryD14Old,0)DeliveryD14Old, ISNULL(DeliveryD15Old,0)DeliveryD15Old,   " & vbCrLf & _
                              "  	ISNULL(DeliveryD16Old,0)DeliveryD16Old, ISNULL(DeliveryD17Old,0)DeliveryD17Old, ISNULL(DeliveryD18Old,0)DeliveryD18Old, ISNULL(DeliveryD19Old,0)DeliveryD19Old, ISNULL(DeliveryD20Old,0)DeliveryD20Old,   " & vbCrLf & _
                              "  	ISNULL(DeliveryD21Old,0)DeliveryD21Old, ISNULL(DeliveryD22Old,0)DeliveryD22Old, ISNULL(DeliveryD23Old,0)DeliveryD23Old, ISNULL(DeliveryD24Old,0)DeliveryD24Old, ISNULL(DeliveryD25Old,0)DeliveryD25Old,   " & vbCrLf & _
                              "  	ISNULL(DeliveryD26Old,0)DeliveryD26Old, ISNULL(DeliveryD27Old,0)DeliveryD27Old, ISNULL(DeliveryD28Old,0)DeliveryD28Old, ISNULL(DeliveryD29Old,0)DeliveryD29Old, ISNULL(DeliveryD30Old,0)DeliveryD30Old,   " & vbCrLf & _
                              "  	ISNULL(DeliveryD31Old,0)DeliveryD31Old  " & vbCrLf & _
                              "  from   " & vbCrLf & _
                              "  (  " & vbCrLf & _
                              "  	select * from  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select '1' NoUrutDesc, 'BY AFFILIATE' AffiliateName  " & vbCrLf & _
                              "  		union all  		 " & vbCrLf

            ls_SQL = ls_SQL + "  		select '2' NoUrutDesc, 'BY PASI' AffiliateName  " & vbCrLf & _
                              "  		union all  " & vbCrLf & _
                              "  		select '3' NoUrutDesc, 'BY SUPPLIER' AffiliateName  " & vbCrLf & _
                              "  		union all  " & vbCrLf & _
                              "  		select '4' NoUrutDesc, 'PO REVISION' AffiliateName  " & vbCrLf & _
                              "  		union all  " & vbCrLf & _
                              "  		select '5' NoUrutDesc, 'DIFFERENCE' AffiliateName  " & vbCrLf & _
                              "  	)tbla  " & vbCrLf & _
                              "  	cross join  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select row_number() over (order by PartNo asc) as NoUrut, " & vbCrLf

            ls_SQL = ls_SQL + "  			PartNo, AffiliateID, AffiliateID2 from " & vbCrLf & _
                              "  		( " & vbCrLf & _
                              "  			select  distinct 			 " & vbCrLf & _
                              "  				b.PartNo, b.PartNo PartNo1, b.AffiliateID, b.AffiliateID AffiliateID2  " & vbCrLf & _
                              "  			from PO_Master a  " & vbCrLf & _
                              "  			inner join PO_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "  			inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  			inner join MS_UnitCls d on d.UnitCls = c.UnitCls  " & vbCrLf & _
                              "  			where YEAR(a.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(a.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "  		)x " & vbCrLf & _
                              "  	)tb1b  " & vbCrLf

            ls_SQL = ls_SQL + "  )tbl1  " & vbCrLf & _
                              "  left join  " & vbCrLf & _
                              "  (  " & vbCrLf & _
                              "  	select   " & vbCrLf & _
                              "   		'BY AFFILIATE' AffiliateName, '1' NoUrutDesc, " & vbCrLf & _
                              "   		b.PartNo, c.PartName, b.AffiliateID, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker, b.PONo, NULL PORevNo, " & vbCrLf & _
                              "   		b.POQty, POQty POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              "   		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf & _
                              "   		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf

            ls_SQL = ls_SQL + "   		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   		b.DeliveryD31,  " & vbCrLf & _
                              "   		0 DeliveryD1Old, 0 DeliveryD2Old, 0 DeliveryD3Old, 0 DeliveryD4Old, 0 DeliveryD5Old,   		 " & vbCrLf & _
                              "   		0 DeliveryD6Old, 0 DeliveryD7Old, 0 DeliveryD8Old, 0 DeliveryD9Old, 0 DeliveryD10Old,   " & vbCrLf & _
                              "   		0 DeliveryD11Old, 0 DeliveryD12Old, 0 DeliveryD13Old, 0 DeliveryD14Old, 0 DeliveryD15Old,   " & vbCrLf & _
                              "   		0 DeliveryD16Old, 0 DeliveryD17Old, 0 DeliveryD18Old, 0 DeliveryD19Old, 0 DeliveryD20Old,   " & vbCrLf & _
                              "   		0 DeliveryD21Old, 0 DeliveryD22Old, 0 DeliveryD23Old, 0 DeliveryD24Old, 0 DeliveryD25Old,   " & vbCrLf & _
                              "   		0 DeliveryD26Old, 0 DeliveryD27Old, 0 DeliveryD28Old, 0 DeliveryD29Old, 0 DeliveryD30Old,   " & vbCrLf & _
                              "   		0 DeliveryD31Old " & vbCrLf & _
                              "  	from PO_Master a  " & vbCrLf

            ls_SQL = ls_SQL + "  	inner join PO_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "  	inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  	inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  	left join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf & _
                              "  	where YEAR(a.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(a.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "  	UNION ALL " & vbCrLf & _
                              "  	select   " & vbCrLf & _
                              "   		'BY PASI' AffiliateName, '2' NoUrutDesc, " & vbCrLf & _
                              "   		b.PartNo, c.PartName, b.AffiliateID, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker, b.PONo, NULL PORevNo, " & vbCrLf & _
                              "   		b.POQty, POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              "   		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf

            ls_SQL = ls_SQL + "   		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   		b.DeliveryD31,  " & vbCrLf & _
                              "   		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old,  " & vbCrLf & _
                              "   		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,   " & vbCrLf & _
                              "   		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,   " & vbCrLf & _
                              "   		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,   " & vbCrLf & _
                              "   		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,   " & vbCrLf

            ls_SQL = ls_SQL + "   		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,   " & vbCrLf & _
                              "   		b.DeliveryD31Old " & vbCrLf & _
                              "  	from Affiliate_Master a  " & vbCrLf & _
                              "  	inner join Affiliate_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "  	inner join PO_Master f on f.PONo = a.PONo and a.AffiliateID = f.AffiliateID and f.SupplierID = a.SupplierID " & vbCrLf & _
                              "  	inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  	inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  	left join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf & _
                              "  	where YEAR(f.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(f.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "  	UNION ALL " & vbCrLf & _
                              "  	select   " & vbCrLf

            ls_SQL = ls_SQL + "   		'BY SUPPLIER' AffiliateName, '3' NoUrutDesc, " & vbCrLf & _
                              "   		b.PartNo, c.PartName, b.AffiliateID, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker, b.PONo, NULL PORevNo, " & vbCrLf & _
                              "   		b.POQty, POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              "   		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf & _
                              "   		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   		b.DeliveryD31,  " & vbCrLf & _
                              "   		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old,  " & vbCrLf

            ls_SQL = ls_SQL + "   		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,   " & vbCrLf & _
                              "   		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,   " & vbCrLf & _
                              "   		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,   " & vbCrLf & _
                              "   		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,   " & vbCrLf & _
                              "   		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,   " & vbCrLf & _
                              "   		b.DeliveryD31Old " & vbCrLf & _
                              "  	from PO_MasterUpload a  " & vbCrLf & _
                              "  	inner join PO_DetailUpload b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "  	inner join PO_Master f on f.PONo = a.PONo and a.AffiliateID = f.AffiliateID and f.SupplierID = a.SupplierID " & vbCrLf & _
                              "  	inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  	inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf

            ls_SQL = ls_SQL + "  	left join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf & _
                              "  	where YEAR(f.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(f.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "  	UNION ALL " & vbCrLf & _
                              "  	select   " & vbCrLf & _
                              "   		'PO REVISION' AffiliateName, '4' NoUrutDesc, " & vbCrLf & _
                              "   		b.PartNo, c.PartName, b.AffiliateID, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker, b.PONo, b.PORevNo, " & vbCrLf & _
                              "   		b.POQty, POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              "   		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf & _
                              "   		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf

            ls_SQL = ls_SQL + "   		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   		b.DeliveryD31,  " & vbCrLf & _
                              "   		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old,  " & vbCrLf & _
                              "   		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,   " & vbCrLf & _
                              "   		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,   " & vbCrLf & _
                              "   		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,   " & vbCrLf & _
                              "   		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,   " & vbCrLf & _
                              "   		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,   " & vbCrLf & _
                              "   		b.DeliveryD31Old " & vbCrLf & _
                              "  	from PORev_Master a  " & vbCrLf

            ls_SQL = ls_SQL + "  	inner join PORev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf & _
                              "  	inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  	inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  	left join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf & _
                              "  	where YEAR(a.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(a.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "  	UNION ALL " & vbCrLf & _
                              "  	select  " & vbCrLf & _
                              " 		'DIFFERENCE' AffiliateName, '5' NoUrutDesc, " & vbCrLf & _
                              " 		PartNo, PartName, AffiliateID, UnitDesc, MOQ, QtyBox, Maker, PONo, PORevNo, " & vbCrLf & _
                              " 		POQty = POQtyOld - POQty, POQtyOld, CurrDesc, Price, Amount, " & vbCrLf & _
                              " 		DeliveryD1 = DeliveryD1Old-DeliveryD1, DeliveryD2 = DeliveryD2Old-DeliveryD2, DeliveryD3 = DeliveryD3Old-DeliveryD3,  " & vbCrLf

            ls_SQL = ls_SQL + " 		DeliveryD4 = DeliveryD4Old-DeliveryD4, DeliveryD5 = DeliveryD5Old-DeliveryD5, DeliveryD6 = DeliveryD6Old-DeliveryD6, " & vbCrLf & _
                              " 		DeliveryD7 = DeliveryD7Old-DeliveryD7, DeliveryD8 = DeliveryD8Old-DeliveryD8, DeliveryD9 = DeliveryD9Old-DeliveryD9,  " & vbCrLf & _
                              " 		DeliveryD10 = DeliveryD10Old-DeliveryD10, DeliveryD11 = DeliveryD11Old-DeliveryD11, DeliveryD12 = DeliveryD12Old-DeliveryD12,  " & vbCrLf & _
                              " 		DeliveryD13 = DeliveryD13Old-DeliveryD13, DeliveryD14 = DeliveryD14Old-DeliveryD14, DeliveryD15 = DeliveryD15Old-DeliveryD15,   " & vbCrLf & _
                              " 		DeliveryD16 = DeliveryD16Old-DeliveryD16, DeliveryD17 = DeliveryD17Old-DeliveryD17, DeliveryD18 = DeliveryD18Old-DeliveryD18,  " & vbCrLf & _
                              " 		DeliveryD19 = DeliveryD19Old-DeliveryD19, DeliveryD20 = DeliveryD20Old-DeliveryD20, DeliveryD21 = DeliveryD21Old-DeliveryD21,  " & vbCrLf & _
                              " 		DeliveryD22 = DeliveryD22Old-DeliveryD22, DeliveryD23 = DeliveryD23Old-DeliveryD23, DeliveryD24 = DeliveryD24Old-DeliveryD24,  " & vbCrLf & _
                              " 		DeliveryD25 = DeliveryD25Old-DeliveryD25, DeliveryD26 = DeliveryD26Old-DeliveryD26, DeliveryD27 = DeliveryD27Old-DeliveryD27,  " & vbCrLf & _
                              " 		DeliveryD28 = DeliveryD28Old-DeliveryD28, DeliveryD29 = DeliveryD29Old-DeliveryD29, DeliveryD30 = DeliveryD30Old-DeliveryD30,   " & vbCrLf & _
                              " 		DeliveryD31 = DeliveryD31Old-DeliveryD31,  " & vbCrLf & _
                              " 		DeliveryD1Old, DeliveryD2Old, DeliveryD3Old, DeliveryD4Old, DeliveryD5Old,  " & vbCrLf

            ls_SQL = ls_SQL + " 		DeliveryD6Old, DeliveryD7Old, DeliveryD8Old, DeliveryD9Old, DeliveryD10Old,   " & vbCrLf & _
                              " 		DeliveryD11Old, DeliveryD12Old, DeliveryD13Old, DeliveryD14Old, DeliveryD15Old,   " & vbCrLf & _
                              " 		DeliveryD16Old, DeliveryD17Old, DeliveryD18Old, DeliveryD19Old, DeliveryD20Old,   " & vbCrLf & _
                              " 		DeliveryD21Old, DeliveryD22Old, DeliveryD23Old, DeliveryD24Old, DeliveryD25Old,   " & vbCrLf & _
                              " 		DeliveryD26Old, DeliveryD27Old, DeliveryD28Old, DeliveryD29Old, DeliveryD30Old,   " & vbCrLf & _
                              " 		DeliveryD31Old  " & vbCrLf & _
                              " 	from  " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 				tbl99.NoUrutDesc, " & vbCrLf & _
                              "   				tbl99.PartNo, tbl99.PartName, tbl99.AffiliateID, tbl99.UnitDesc, tbl99.MOQ, tbl99.QtyBox, tbl99.Maker, tbl99.PONo, tbl99.PORevNo, " & vbCrLf

            ls_SQL = ls_SQL + "   				tbl99.POQty, tbl99.POQtyOld, tbl99.CurrDesc, tbl99.Price, tbl99.Amount, " & vbCrLf & _
                              "   				tbl99.DeliveryD1, tbl99.DeliveryD2, tbl99.DeliveryD3, tbl99.DeliveryD4, tbl99.DeliveryD5,  " & vbCrLf & _
                              "   				tbl99.DeliveryD6, tbl99.DeliveryD7, tbl99.DeliveryD8, tbl99.DeliveryD9, tbl99.DeliveryD10,   " & vbCrLf & _
                              "   				tbl99.DeliveryD11, tbl99.DeliveryD12, tbl99.DeliveryD13, tbl99.DeliveryD14, tbl99.DeliveryD15,   " & vbCrLf & _
                              "   				tbl99.DeliveryD16, tbl99.DeliveryD17, tbl99.DeliveryD18, tbl99.DeliveryD19, tbl99.DeliveryD20,   " & vbCrLf & _
                              "   				tbl99.DeliveryD21, tbl99.DeliveryD22, tbl99.DeliveryD23, tbl99.DeliveryD24, tbl99.DeliveryD25,   " & vbCrLf & _
                              "   				tbl99.DeliveryD26, tbl99.DeliveryD27, tbl99.DeliveryD28, tbl99.DeliveryD29, tbl99.DeliveryD30,   " & vbCrLf & _
                              "   				tbl99.DeliveryD31,  " & vbCrLf & _
                              "   				tbl99.DeliveryD1Old, tbl99.DeliveryD2Old, tbl99.DeliveryD3Old, tbl99.DeliveryD4Old, tbl99.DeliveryD5Old,   		 " & vbCrLf & _
                              "   				tbl99.DeliveryD6Old, tbl99.DeliveryD7Old, tbl99.DeliveryD8Old, tbl99.DeliveryD9Old, tbl99.DeliveryD10Old,   " & vbCrLf & _
                              "   				tbl99.DeliveryD11Old, tbl99.DeliveryD12Old, tbl99.DeliveryD13Old, tbl99.DeliveryD14Old, tbl99.DeliveryD15Old,   " & vbCrLf

            ls_SQL = ls_SQL + "   				tbl99.DeliveryD16Old, tbl99.DeliveryD17Old, tbl99.DeliveryD18Old, tbl99.DeliveryD19Old, tbl99.DeliveryD20Old,   " & vbCrLf & _
                              "   				tbl99.DeliveryD21Old, tbl99.DeliveryD22Old, tbl99.DeliveryD23Old, tbl99.DeliveryD24Old, tbl99.DeliveryD25Old,   " & vbCrLf & _
                              "   				tbl99.DeliveryD26Old, tbl99.DeliveryD27Old, tbl99.DeliveryD28Old, tbl99.DeliveryD29Old, tbl99.DeliveryD30Old,   " & vbCrLf & _
                              "   				tbl99.DeliveryD31Old " & vbCrLf & _
                              " 		from " & vbCrLf & _
                              " 		( " & vbCrLf & _
                              " 			select   " & vbCrLf & _
                              "   				'BY AFFILIATE' AffiliateName, '1' NoUrutDesc, " & vbCrLf & _
                              "   				b.PartNo, c.PartName, b.AffiliateID, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker, b.PONo, NULL PORevNo, " & vbCrLf & _
                              "   				b.POQty, b.POQty POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              "   				b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf

            ls_SQL = ls_SQL + "   				b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   				b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   				b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   				b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   				b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   				b.DeliveryD31,  " & vbCrLf & _
                              "   				DeliveryD1 DeliveryD1Old, DeliveryD2 DeliveryD2Old, DeliveryD3 DeliveryD3Old, DeliveryD4 DeliveryD4Old, DeliveryD5 DeliveryD5Old,   		 " & vbCrLf & _
                              "   				DeliveryD6 DeliveryD6Old, DeliveryD7 DeliveryD7Old, DeliveryD8 DeliveryD8Old, DeliveryD9 DeliveryD9Old, DeliveryD10 DeliveryD10Old,   " & vbCrLf & _
                              "   				DeliveryD11 DeliveryD11Old, DeliveryD12 DeliveryD12Old, DeliveryD13 DeliveryD13Old, DeliveryD14 DeliveryD14Old, DeliveryD15 DeliveryD15Old,   " & vbCrLf & _
                              "   				DeliveryD16 DeliveryD16Old, DeliveryD17 DeliveryD17Old, DeliveryD18 DeliveryD18Old, DeliveryD19 DeliveryD19Old, DeliveryD20 DeliveryD20Old,   " & vbCrLf & _
                              "   				DeliveryD21 DeliveryD21Old, DeliveryD22 DeliveryD22Old, DeliveryD23 DeliveryD23Old, DeliveryD24 DeliveryD24Old, DeliveryD25 DeliveryD25Old,   " & vbCrLf

            ls_SQL = ls_SQL + "   				DeliveryD26 DeliveryD26Old, DeliveryD27 DeliveryD27Old, DeliveryD28 DeliveryD28Old, DeliveryD29 DeliveryD29Old, DeliveryD30 DeliveryD30Old,   " & vbCrLf & _
                              "   				DeliveryD31 DeliveryD31Old " & vbCrLf & _
                              "  			from PO_Master a  " & vbCrLf & _
                              "  			inner join PO_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "  			inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  			inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  			left join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf & _
                              "  			where YEAR(a.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(a.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "  			UNION ALL " & vbCrLf & _
                              "  			select   " & vbCrLf & _
                              "   				'BY PASI' AffiliateName, '2' NoUrutDesc, " & vbCrLf

            ls_SQL = ls_SQL + "   				b.PartNo, c.PartName, b.AffiliateID, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker, b.PONo, NULL PORevNo, " & vbCrLf & _
                              "   				b.POQty, POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              "   				b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf & _
                              "   				b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   				b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   				b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   				b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   				b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   				b.DeliveryD31,  " & vbCrLf & _
                              "   				b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old,  " & vbCrLf & _
                              "   				b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,   " & vbCrLf

            ls_SQL = ls_SQL + "   				b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,   " & vbCrLf & _
                              "   				b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,   " & vbCrLf & _
                              "   				b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,   " & vbCrLf & _
                              "   				b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,   " & vbCrLf & _
                              "   				b.DeliveryD31Old " & vbCrLf & _
                              "  			from Affiliate_Master a  " & vbCrLf & _
                              "  			inner join Affiliate_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "  			inner join PO_Master f on f.PONo = a.PONo and a.AffiliateID = f.AffiliateID and f.SupplierID = a.SupplierID " & vbCrLf & _
                              "  			inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  			inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  			left join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf

            ls_SQL = ls_SQL + "  			where YEAR(f.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(f.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "  			UNION ALL " & vbCrLf & _
                              "  			select   " & vbCrLf & _
                              "   				'BY SUPPLIER' AffiliateName, '3' NoUrutDesc, " & vbCrLf & _
                              "   				b.PartNo, c.PartName, b.AffiliateID, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker, b.PONo, NULL PORevNo, " & vbCrLf & _
                              "   				b.POQty, POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              "   				b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf & _
                              "   				b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   				b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   				b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   				b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf

            ls_SQL = ls_SQL + "   				b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   				b.DeliveryD31,  " & vbCrLf & _
                              "   				b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old,  " & vbCrLf & _
                              "   				b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,   " & vbCrLf & _
                              "   				b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,   " & vbCrLf & _
                              "   				b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,   " & vbCrLf & _
                              "   				b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,   " & vbCrLf & _
                              "   				b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,   " & vbCrLf & _
                              "   				b.DeliveryD31Old " & vbCrLf & _
                              "  			from PO_MasterUpload a  " & vbCrLf & _
                              "  			inner join PO_DetailUpload b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "  			inner join PO_Master f on f.PONo = a.PONo and a.AffiliateID = f.AffiliateID and f.SupplierID = a.SupplierID " & vbCrLf & _
                              "  			inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  			inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  			left join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf & _
                              "  			where YEAR(f.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(f.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              "  			UNION ALL " & vbCrLf & _
                              "  			select   " & vbCrLf & _
                              "   				'PO REVISION' AffiliateName, '4' NoUrutDesc, " & vbCrLf & _
                              "   				b.PartNo, c.PartName, b.AffiliateID, d.Description UnitDesc, c.MOQ, c.QtyBox, c.Maker, b.PONo, b.PORevNo, " & vbCrLf & _
                              "   				b.POQty, POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              "   				b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf

            ls_SQL = ls_SQL + "   				b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   				b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   				b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   				b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   				b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   				b.DeliveryD31,  " & vbCrLf & _
                              "   				b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old,  " & vbCrLf & _
                              "   				b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,   " & vbCrLf & _
                              "   				b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,   " & vbCrLf & _
                              "   				b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,   " & vbCrLf & _
                              "   				b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,   " & vbCrLf

            ls_SQL = ls_SQL + "   				b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,   " & vbCrLf & _
                              "   				b.DeliveryD31Old " & vbCrLf & _
                              "  			from PORev_Master a  " & vbCrLf & _
                              "  			inner join PORev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf & _
                              "  			inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  			inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  			left join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf & _
                              "  			where YEAR(a.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(a.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              " 		)tbl99 inner join " & vbCrLf & _
                              " 		(	 " & vbCrLf & _
                              " 			select  " & vbCrLf

            ls_SQL = ls_SQL + " 				MAX(NoUrutDesc)NoUrutDesc, max(PORevNo) PORevNo, PONo, PartNo, AffiliateID " & vbCrLf & _
                              " 			from " & vbCrLf & _
                              " 				( " & vbCrLf & _
                              " 				select   " & vbCrLf & _
                              " 					'BY AFFILIATE' AffiliateName, '1' NoUrutDesc, b.PONo, NULL PORevNo, PartNo, b.AffiliateID " & vbCrLf & _
                              " 				from PO_Master a  " & vbCrLf & _
                              " 				inner join PO_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " 				where YEAR(a.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(a.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              " 				UNION ALL " & vbCrLf & _
                              " 				select   " & vbCrLf & _
                              " 					'BY PASI' AffiliateName, '2' NoUrutDesc, b.PONo, NULL PORevNo, PartNo, b.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 				from Affiliate_Master a  " & vbCrLf & _
                              " 				inner join Affiliate_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " 				inner join PO_Master f on f.PONo = a.PONo and a.AffiliateID = f.AffiliateID and f.SupplierID = a.SupplierID	 " & vbCrLf & _
                              " 				where YEAR(f.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(f.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf & _
                              " 				UNION ALL " & vbCrLf & _
                              " 				select   " & vbCrLf & _
                              " 					'BY SUPPLIER' AffiliateName, '3' NoUrutDesc, b.PONo, NULL PORevNo, PartNo, b.AffiliateID " & vbCrLf & _
                              " 				from PO_MasterUpload a  " & vbCrLf & _
                              " 				inner join PO_DetailUpload b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " 				inner join PO_Master f on f.PONo = a.PONo and a.AffiliateID = f.AffiliateID and f.SupplierID = a.SupplierID	 " & vbCrLf & _
                              " 				where YEAR(f.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(f.Period) = " & Month(dtPeriodFrom.Value) & " " & vbCrLf

            ls_SQL = ls_SQL + " 				UNION ALL " & vbCrLf & _
                              " 				select   " & vbCrLf & _
                              " 					'PO REVISION' AffiliateName, '4' NoUrutDesc, b.PONo, b.PORevNo, PartNo, b.AffiliateID " & vbCrLf & _
                              " 				from PORev_Master a  " & vbCrLf & _
                              " 				inner join PORev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo  	 " & vbCrLf & _
                              " 				where YEAR(a.Period) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(a.Period) = " & Month(dtPeriodFrom.Value) & "  " & vbCrLf & _
                              " 			)tbl99 " & vbCrLf & _
                              " 			group by PONo, PartNo, AffiliateID " & vbCrLf & _
                              " 		)tbl88 on tbl88.NoUrutDesc = tbl99.NoUrutDesc  " & vbCrLf & _
                              " 		and tbl88.PONo = tbl99.PONo and isnull(tbl88.PORevNo,'') = isnull(tbl99.PORevNo,'') " & vbCrLf & _
                              " 		and tbl88.PartNo = tbl99.PartNo " & vbCrLf

            ls_SQL = ls_SQL + " 	)tblDesc " & vbCrLf & _
                              " 	)tbl2 on tbl2.AffiliateName = tbl1.AffiliateName and tbl1.PartNo = tbl2.PartNo and tbl1.NoUrutDesc = tbl2.NoUrutDesc and tbl1.AffiliateID = tbl2.AffiliateID " & vbCrLf & _
                              " where tbl2.PartNo is not null " & pWhere & " " & vbCrLf & _
                              " ) x " & vbCrLf & _
                              "  " & vbCrLf & _
                              " order by PartNo2, AffiliateID2, PONo ,Header, NoUrutDesc "

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

            ls_SQL = " select top 0 '' NoUrut, '' AffiliateName, '' PartNo, '' PartName, '' UnitDesc, '' MOQ, '' QtyBox, '' Maker, " & vbCrLf & _
                  " 0 POQty, 0 POQtyOld, '' CurrDesc, '' Price, '' Amount, PONo," & vbCrLf & _
                  " 0 DeliveryD1, 0 DeliveryD2, 0 DeliveryD3, 0 DeliveryD4, 0 DeliveryD5, " & vbCrLf & _
                  " 0 DeliveryD6, 0 DeliveryD7, 0 DeliveryD8, 0 DeliveryD9, 0 DeliveryD10, " & vbCrLf & _
                  " 0 DeliveryD11, 0 DeliveryD12, 0 DeliveryD13, 0 DeliveryD14, 0 DeliveryD15, " & vbCrLf & _
                  " 0 DeliveryD16, 0 DeliveryD17, 0 DeliveryD18, 0 DeliveryD19, 0 DeliveryD20, " & vbCrLf & _
                  " 0 DeliveryD21, 0 DeliveryD22, 0 DeliveryD23, 0 DeliveryD24, 0 DeliveryD25, " & vbCrLf & _
                  " 0 DeliveryD26, 0 DeliveryD27, 0 DeliveryD28, 0 DeliveryD29, 0 DeliveryD30, " & vbCrLf & _
                  " 0 DeliveryD31, " & vbCrLf & _
                  " 0 DeliveryD1Old, 0 DeliveryD2Old, 0 DeliveryD3Old, 0 DeliveryD4Old, 0 DeliveryD5Old, " & vbCrLf & _
                  " 0 DeliveryD6Old, 0 DeliveryD7Old, 0 DeliveryD8Old, 0 DeliveryD9Old, 0 DeliveryD10Old, " & vbCrLf & _
                  " 0 DeliveryD11Old, 0 DeliveryD12Old, 0 DeliveryD13Old, 0 DeliveryD14Old, 0 DeliveryD15Old, " & vbCrLf & _
                  " 0 DeliveryD16Old, 0 DeliveryD17Old, 0 DeliveryD18Old, 0 DeliveryD19Old, 0 DeliveryD20Old, " & vbCrLf & _
                  " 0 DeliveryD21Old, 0 DeliveryD22Old, 0 DeliveryD23Old, 0 DeliveryD24Old, 0 DeliveryD25Old, " & vbCrLf & _
                  " 0 DeliveryD26Old, 0 DeliveryD27Old, 0 DeliveryD28Old, 0 DeliveryD29Old, 0 DeliveryD30Old, " & vbCrLf & _
                  " 0 DeliveryD31Old "

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

    Private Sub up_FillCombo(ByVal pPeriod As String)
        Dim ls_SQL As String = ""

        ls_SQL = "select '" & clsGlobal.gs_All & "' PONo union all select RTRIM(PONo) PONo from PO_Master where Year(Period) = '" & Year(pPeriod) & "' and month(Period) = '" & Month(pPeriod) & "' order by PONo " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPONo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PONo")
                .Columns(0).Width = 180

                .TextField = "PONo"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_FillComboPart()
        Dim ls_SQL As String = ""

        ls_SQL = "select '" & clsGlobal.gs_All & "'PartNo, '" & clsGlobal.gs_All & "'PartName union all select RTRIM(a.PartNo)PartNo, RTRIM(b.PartName) PartName from MS_PartMapping a left join MS_Parts b on a.PartNo = b.PartNo order by PartNo " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 85
                .Columns.Add("PartName")
                .Columns(1).Width = 180

                .TextField = "PartNo"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_FillComboAffiliateID()
        Dim ls_SQL As String = ""

        ls_SQL = "select '" & clsGlobal.gs_All & "'AffiliateID, '" & clsGlobal.gs_All & "'AffiliateName union all select RTRIM(AffiliateID)AffiliateID, RTRIM(AffiliateName) AffiliateName from MS_Affiliate Order by AffiliateID " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateID
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 85
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 180

                .TextField = "AffiliateID"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
    End Sub

#End Region


End Class