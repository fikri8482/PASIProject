Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports OfficeOpenXml
Imports Microsoft.Office.Interop
Imports System.Net
Imports System.Net.Mail
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO

Public Class POExportFinalApprovalMonthly
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "B04"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_AffiliateID As String, pub_AffiliateName As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String, pub_SupplierName As String, pub_Remarks As String
    Dim pub_FinalApproval As String, pub_DeliveyBy As String
    Dim pub_Period As Date
    Dim pub_HeijunkaSttus As Boolean

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String

    Dim flag As Boolean = True

    Dim pStatus As Boolean

    Dim pPeriod As Date
    Dim pCommercial As String
    Dim pDeliveryCode As String
    Dim pDeliveryName As String
    Dim pPOEmergency As String
    Dim pShipBy As String
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pSupplierCode As String
    Dim pSupplierName As String
    Dim pPORevNo As String
    Dim pPO As String
    Dim pRemarks As String

    Dim pOrderNo As String
    Dim pETDVendorOrder As String
    Dim pETDVendorSupplier As String
    Dim pETDPort As String
    Dim pETAPort As String
    Dim pETAFactory As String

    Dim pFilter As String
    Dim pub_Param As String
    Dim pstatusInsert As String
#End Region

#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim param As String = ""
        Dim filterQty As String = ""


        Try
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                'If Not IsNothing(Request.QueryString("prm")) Then


                If Session("PORevExportList") <> "" Then
                    param = Session("PORevExportList").ToString()
                ElseIf Session("TampungDelivery") <> "" Then
                    param = Session("TampungDelivery").ToString()
                Else
                    param = Request.QueryString("prm").ToString
                End If

                If Split(param, "|")(3) = "E" Then
                    Session("MenuDesc") = "PO FROM SUPPLIER APPROVE BY PASI (EMERGENCY)"
                Else
                    Session("MenuDesc") = "PO FROM SUPPLIER APPROVE BY PASI (MONTHLY)"
                End If


                If param = "  'back'" Then
                    btnSubMenu.Text = "BACK"
                Else
                    If pStatus = False And Session("GOTOStatus") <> "3" Then
                        '2016-01-01|2016-01-01|1|M|B|HESTO|HESTO HARNESSES (PROPRIETARY) LIMITED|DHL|DHL Indonesia||YGP16011|2016-01-04|2016-01-04|2016-01-06|2016-01-07|2016-01-08|YGP16011
                        pPeriod = Split(param, "|")(0)
                        pAffiliateCode = Split(param, "|")(5)
                        pAffiliateName = Split(param, "|")(2)
                        pSupplierCode = Split(param, "|")(1)
                        pSupplierName = Split(param, "|")(4)
                        pDeliveryCode = Split(param, "|")(7)
                        pDeliveryName = Split(param, "|")(8)
                        pCommercial = Split(param, "|")(2)
                        pPOEmergency = Split(param, "|")(3)
                        pShipBy = Split(param, "|")(4)
                        pRemarks = Split(param, "|")(10)
                        pPO = Split(param, "|")(16)
                        pOrderNo = Split(param, "|")(10)

                        pStatus = True

                        Call bindDataHeader(pAffiliateCode, pPO, pOrderNo, pSupplierCode)
                        Call bindDataDetail(pAffiliateCode, pPO, pOrderNo, pSupplierCode)

                        Session("pFilter") = pFilter
                        Session.Remove("POList")

                    ElseIf Session("GOTOStatus") = "3" Then
                        lblInfo.Text = ""
                        pPeriod = Split(param, "|")(0)
                        pSupplierCode = Split(param, "|")(1)
                        pCommercial = Split(param, "|")(2)
                        pPOEmergency = Split(param, "|")(3)
                        pShipBy = Split(param, "|")(4)
                        pAffiliateCode = Split(param, "|")(5)
                        pAffiliateName = Split(param, "|")(6)
                        pDeliveryCode = Split(param, "|")(7)
                        pDeliveryName = Split(param, "|")(8)
                        pRemarks = Split(param, "|")(9)
                        pOrderNo = Split(param, "|")(10)
                        pETDVendorOrder = Split(param, "|")(11)
                        pETDVendorSupplier = Split(param, "|")(12)
                        pETDPort = Split(param, "|")(13)
                        pETAPort = Split(param, "|")(14)
                        pETAFactory = Split(param, "|")(15)
                        pPO = Split(param, "|")(16)

                        If pAffiliateCode <> "" Then btnSubMenu.Text = "BACK"

                        Call bindDataHeader(pAffiliateCode, pPO, pOrderNo, pSupplierCode)
                        Call bindDataDetail(pAffiliateCode, pPO, pOrderNo, pSupplierCode)

                        Session("pFilter") = pFilter
                        Session.Remove("EmergencyUrl")
                        btnSubMenu.Text = "BACK"

                    ElseIf Session("GOTOStatus") = "tiga" Then
                        lblInfo.Text = ""
                        pPeriod = Split(param, "|")(1)
                        pCommercial = Split(param, "|")(2)
                        pPOEmergency = Split(param, "|")(3)
                        pShipBy = Split(param, "|")(4)
                        pAffiliateCode = Split(param, "|")(5)
                        pAffiliateName = Split(param, "|")(6)
                        pDeliveryCode = Split(param, "|")(7)
                        pDeliveryName = Split(param, "|")(8)
                        pRemarks = Split(param, "|")(9)
                        pOrderNo = Split(param, "|")(10)
                        pETDVendorOrder = Split(param, "|")(11)
                        pETDVendorSupplier = Split(param, "|")(12)
                        pETDPort = Split(param, "|")(13)
                        pETAPort = Split(param, "|")(14)
                        pETAFactory = Split(param, "|")(15)
                        pPO = Split(param, "|")(16)

                        If pAffiliateCode <> "" Then btnSubMenu.Text = "BACK"

                        Call bindDataHeader(pAffiliateCode, pPO, pOrderNo, pSupplierCode)
                        Call bindDataDetail(pAffiliateCode, pPO, pOrderNo, pSupplierCode)

                        Session("pFilter") = pFilter
                        Session.Remove("EmergencyUrl")
                        btnSubMenu.Text = "BACK"
                    End If
                End If
                btnSubMenu.Text = "BACK"
            End If
            '===============================================================================

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblInfo.Text = ""
                'dt1.Value = Format(txtkanbandate.text, "MMM yyyy")
            End If

            'Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            grid.JSProperties("cpMessage") = lblInfo.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        If btnSubMenu.Text = "BACK" And Session("GOTOStatus") = "3" Then
            Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
        ElseIf btnSubMenu.Text = "BACK" And Session("GOTOStatus") = "tiga" Then
            Response.Redirect("~/PurchaseOrderExport/POExportFinalApprovalList.aspx")
        ElseIf btnSubMenu.Text = "BACK" Then
            Response.Redirect("~/PurchaseOrderExport/POExportFinalApprovalList.aspx")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If

        Session.Remove("GOTOStatus")
    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Call uf_Approve()
                    Call bindDataDetail(cboAffiliate.Text, txtpono.Text, txtOrderNo.Text, Session("SupplierID"))
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "1009", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    Else
                        Call clsMsg.DisplayMessage(lblInfo, "1009", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If
            End Select

EndProcedure:
            Session("AA220Msg") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        If x > grid.VisibleRowCount Then Exit Sub

        With grid
            If .VisibleRowCount > 0 Then
                If e.GetValue("AffiliateName") = "SUPPLIER APPROVAL" Then
                    If e.DataColumn.FieldName = "UnitDesc" Or e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" _
                        Or e.DataColumn.FieldName = "Forecast1" Or e.DataColumn.FieldName = "Forecast2" _
                        Or e.DataColumn.FieldName = "Forecast3" Or e.DataColumn.FieldName = "Variance" Or e.DataColumn.FieldName = "VarPecentage" _
                        Or e.DataColumn.FieldName = "PreviousForecast" Then
                        e.Cell.Text = ""
                    End If
                    If CDbl(e.GetValue("POQty")) <> CDbl(e.GetValue("POQtyOld")) Then
                        If e.DataColumn.FieldName = "POQty" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("Week1")) <> CDbl(e.GetValue("Week1Old")) Then
                        If e.DataColumn.FieldName = "Week1" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("Week2")) <> CDbl(e.GetValue("Week2Old")) Then
                        If e.DataColumn.FieldName = "Week2" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("Week3")) <> CDbl(e.GetValue("Week3Old")) Then
                        If e.DataColumn.FieldName = "Week3" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("Week4")) <> CDbl(e.GetValue("Week4Old")) Then
                        If e.DataColumn.FieldName = "Week4" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("Week5")) <> CDbl(e.GetValue("Week5Old")) Then
                        If e.DataColumn.FieldName = "Week5" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                End If
            End If
        End With
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub uf_Approve()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                ls_sql = " Update PO_Master_Export set PASIApproveDate = getdate(), PASIApproveUser = '" & Session("UserID") & "', finalApprovalCls = '1'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf & _
                            " AND PONo = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                            " AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf & _
                            " AND SupplierID = '" & Session("SupplierID") & "'" & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_Fillcombo()
        Dim ls_SQL As String = ""
        'Combo Affiliate
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' AffiliateID, '" & clsGlobal.gs_All & "' AffiliateName UNION ALL SELECT RTRIM(AffiliateID) AffiliateID,AffiliateName FROM dbo.MS_Affiliate where isnull(overseascls, '0') = '1'" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 50
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 120

                .TextField = "AffiliateID"
                .DataBind()
                .SelectedIndex = 0
                txtAffiliate.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataHeader(ByVal pAffCode As String, ByVal pPONO As String, ByVal pOrderNo As String, ByVal pSuppCode As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT  " & vbCrLf & _
                  " 	a.Period, a.PONo, a.OrderNo1, a.CommercialCls, a.EmergencyCls, a.ShipCls, a.AffiliateID, x.AffiliateName, " & vbCrLf & _
                  " 	a.SupplierID, y.SupplierName, b.ETDVendor1, a.ETDPort1, a.ETAPort1, a.ETAFactory1," & vbCrLf & _
                  " 	a.ForwarderID, z.ForwarderName, b.Remarks " & vbCrLf & _
                  " FROM PO_Master_Export a " & vbCrLf & _
                  " INNER JOIN PO_MasterUpload_Export b ON a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.ForwarderID = b.ForwarderID " & vbCrLf & _
                  " Left JOin MS_Affiliate x on x.AffiliateID = a.AffiliateID" & vbCrLf & _
                  " Left JOin MS_Supplier y on y.SupplierID = a.SupplierID" & vbCrLf & _
                  " Left JOin MS_Forwarder z on z.ForwarderID = a.ForwarderID" & vbCrLf & _
                  " WHERE a.PONo = '" & pPONO & "'" & vbCrLf & _
                  " AND a.OrderNo1 = '" & pOrderNo & "'" & vbCrLf & _
                  " AND a.AffiliateID = '" & pAffCode & "'" & vbCrLf & _
                  " AND a.SupplierID = '" & pSuppCode & "' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                dtPeriodFrom.Text = Format(ds.Tables(0).Rows(0)("Period"), "yyyy-MM") 'ds.Tables(0).Rows(0)("Period") & ""
                If ds.Tables(0).Rows(0)("CommercialCls") = "1" Then
                    rdrCom1.Checked = True
                Else
                    rdrCom2.Checked = True
                End If

                If ds.Tables(0).Rows(0)("EmergencyCls") = "E" Then
                    rdEmergency.Checked = True
                Else
                    rdMonthly.Checked = True
                End If

                If ds.Tables(0).Rows(0)("ShipCls") = "B" Then
                    rdrShipBy2.Checked = True
                Else
                    rdrShipBy3.Checked = True
                End If

                cboAffiliate.Text = ds.Tables(0).Rows(0)("AffiliateID") & "" '"AffiliateCode"
                txtAffiliate.Text = ds.Tables(0).Rows(0)("AffiliateName") & "" '"AffiliateName"
                txtpono.Text = ds.Tables(0).Rows(0)("PONo") & "" 'PONo
                txtOrderNo.Text = ds.Tables(0).Rows(0)("OrderNo1") & "" 'SPLIT ORDER NO
                cboDelLoc.Text = ds.Tables(0).Rows(0)("ForwarderID") & ""
                txtDelLoc.Text = ds.Tables(0).Rows(0)("ForwarderName") & ""
                dtETDVendor.Text = Format(ds.Tables(0).Rows(0)("ETDVendor1"), "yyyy-MM-dd")
                dtETDPort.Text = Format(ds.Tables(0).Rows(0)("ETDPort1"), "yyyy-MM-dd")
                dtETAPort.Text = Format(ds.Tables(0).Rows(0)("ETAPort1"), "yyyy-MM-dd")
                dtETAFactory.Text = Format(ds.Tables(0).Rows(0)("ETAFactory1"), "yyyy-MM-dd")

                Session("SupplierID") = ds.Tables(0).Rows(0)("SupplierID") & ""

                Call clsMsg.DisplayMessage(lblInfo, "1008", clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                Session("YA010IsSubmit") = lblInfo.Text
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataDetail(ByVal pAffCode As String, ByVal pPONO As String, ByVal pOrderNo As String, ByVal psuppCode As String)
        Dim ls_SQL As String = ""
        Dim jsScript As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select convert(char,row_number() over (order by tbl1.PartNo, tbl1.urutan asc))as NoUrutAbal, tbl2.NoUrut, tbl1.AffiliateName, tbl2.PartNo, tbl1.PartNo1, tbl2.PartName, tbl2.UnitDesc, tbl2.MOQ, tbl2.QtyBox,  " & vbCrLf & _
                  " ISNULL(TotalPOQty,0)POQty, ISNULL(TotalPOQty,0)POQtyOld,  " & vbCrLf & _
                  " ISNULL(Week1,0)Week1, ISNULL(Week2,0)Week2, ISNULL(Week3,0)Week3, ISNULL(Week4,0)Week4, ISNULL(Week5,0)Week5,   " & vbCrLf & _
                  " ISNULL(Week1Old,0)Week1Old, ISNULL(Week2Old,0)Week2Old, ISNULL(Week3Old,0)Week3Old, ISNULL(Week4Old,0)Week4Old, ISNULL(Week5Old,0)Week5Old, " & vbCrLf & _
                  " ISNULL(PreviousForecast,0)PreviousForecast, ISNULL(Variance,0)Variance,ISNULL(VariancePercentage,0)VarPecentage,ISNULL(Forecast1,0)Forecast1,ISNULL(Forecast2,0)Forecast2,ISNULL(Forecast3,0)Forecast3, SupplierID " & vbCrLf & _
                  " from   " & vbCrLf & _
                  " (  " & vbCrLf & _
                  " 	select distinct * from  " & vbCrLf & _
                  " 	(  " & vbCrLf & _
                  " 		select distinct '1' NoUrutDesc, 'ORDER' AffiliateName, urutan = 1  " & vbCrLf & _
                  " 		union all "

            ls_SQL = ls_SQL + " 		select distinct '2' NoUrutDesc, 'SUPPLIER APPROVAL' AffiliateName, urutan = 2 " & vbCrLf & _
                              " 	)tbla  " & vbCrLf & _
                              " 	cross join  " & vbCrLf & _
                              " 	(  " & vbCrLf & _
                              " 		select distinct                       " & vbCrLf & _
                              " 			b.PartNo, b.PartNo PartNo1  " & vbCrLf & _
                              " 		from PO_Master_Export a  " & vbCrLf & _
                              " 		inner join po_Detail_export b on a.OrderNo1 = b.OrderNo1 and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID					 " & vbCrLf & _
                              " 		where a.OrderNo1 = '" & pOrderNo & "' and a.PoNo = '" & Trim(pPONO) & "' and a.AffiliateID = '" & pAffCode & "' and a.SupplierID = '" & Trim(psuppCode) & "' " & vbCrLf & _
                              " 	)tb1b  " & vbCrLf & _
                              " )tbl1  "

            ls_SQL = ls_SQL + " left join  " & vbCrLf & _
                              " (  " & vbCrLf & _
                              " 	select convert(char,row_number() over (order by PartNo asc))as NoUrut, * from ( " & vbCrLf & _
                              "     select distinct    " & vbCrLf & _
                              " 	'ORDER' AffiliateName, '1' NoUrutDesc,   " & vbCrLf & _
                              " 	b.PartNo, b.PartNo PartNo1, c.PartName, d.Description UnitDesc,  " & vbCrLf & _
                              " 	MOQ = ISNULL(b.POMOQ,e.MOQ), QtyBox = ISNULL(b.POQtyBox,e.QtyBox) , b.TotalPOQty, 0 TotalPOQtyOld,  " & vbCrLf & _
                              " 	b.Week1, b.Week2, b.Week3, b.Week4, b.Week5, " & vbCrLf & _
                              " 	0 Week1Old, 0 Week2Old, 0 Week3Old, 0 Week4Old, 0 Week5Old, " & vbCrLf & _
                              " 	PreviousForecast = CASE WHEN a.EmergencyCls = 'E' then 0 else ISNULL(PrevQty.Forecast1,0) END, " & vbCrLf & _
                              "     Variance = CASE WHEN a.EmergencyCls = 'E' then 0 else CASE WHEN ISNULL(PrevQty.Forecast1,0) = 0 THEN 0 ELSE B.Week1 - PrevQty.Forecast1 END END, " & vbCrLf & _
                              "     VariancePercentage = CASE WHEN a.EmergencyCls = 'E' then 0 else CASE WHEN ISNULL(PrevQty.Forecast1,0) = 0 THEN 0 ELSE ((B.Week1 - PrevQty.Forecast1) / PrevQty.Forecast1) * 100 END END," & vbCrLf & _
                              " 	b.Forecast1, b.Forecast2, b.Forecast3, urutan = 1, a.SupplierID " & vbCrLf & _
                              " 	from PO_Master_Export a  "

            ls_SQL = ls_SQL + " 	inner join PO_Detail_Export b on a.OrderNo1 = b.OrderNo1 and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                              "    LEFT JOIN ( " & vbCrLf & _
                              "       SELECT Forecast1, PartNo, a.AffiliateID, a.PONo, a.OrderNo1 FROM PO_Detail_Export a " & vbCrLf & _
                              "       INNER JOIN PO_Master_Export b ON a.PONo = b.PONo and a.OrderNo1 = b.OrderNo1 and a.AffiliateID = b.AffiliateID  and a.SupplierID = b.SupplierID " & vbCrLf & _
                              "       WHERE Period = '" & DateAdd(DateInterval.Month, -1, dtPeriodFrom.Value) & "' and a.PONo = a.PONo and b.EmergencyCls <> 'E' and a.OrderNo1 = b.OrderNo1 and Forecast1 > 0" & vbCrLf & _
                              "    )PrevQty ON PrevQty.PartNo = b.PartNo and PrevQty.AffiliateID = b.AffiliateID --and PrevQty.PONo = b.PONo and PrevQty.OrderNo1 = b.OrderNo1" & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "     left join MS_PartMapping e on e.AffiliateID = b.AffiliateID and e.SupplierID = b.SupplierID and e.PartNo = b.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              " 	where a.OrderNo1 = '" & pOrderNo & "' and a.PoNo = '" & Trim(pPONO) & "' and a.AffiliateID = '" & pAffCode & "'and a.supplierID = '" & Trim(psuppCode) & "')zyx " & vbCrLf & _
                              " 	union all  " & vbCrLf & _
                              " 	select distinct   " & vbCrLf & _
                              " 	'' NoUrut, 'SUPPLIER APPROVAL' AffiliateName, '2' NoUrutDesc,   " & vbCrLf & _
                              " 	'' PartNo, b.PartNo PartNo1, '' PartName, '' UnitDesc,  " & vbCrLf & _
                              " 	0 MOQ, 0 QtyBox, b.TotalPOQty, b.TotalPOQtyOld,  " & vbCrLf & _
                              " 	0 Week1, 0 Week2, 0 Week3, 0 Week4, 0 Week5, " & vbCrLf & _
                              " 	0 Week1Old, 0 Week2Old, 0 Week3Old, 0 Week4Old, 0 Week5Old,  "

            ls_SQL = ls_SQL + " 	0 PreviousForecast, Variance = 0 , VarPecentage = 0 , " & vbCrLf & _
                              " 	0 Forecast1, 0 Forecast2, 0 Forecast3, urutan = 2, '' SupplierID " & vbCrLf & _
                              " 	from PO_MasterUpload_Export a  " & vbCrLf & _
                              " 	inner join PO_DetailUpload_Export b on a.OrderNo1 = b.OrderNo1 and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls 	 " & vbCrLf & _
                              " 	where a.PONo = '" & pPONO & "'" & vbCrLf & _
                              "     and a.OrderNo1 = '" & pOrderNo & "' " & vbCrLf & _
                              "     and a.AffiliateID = '" & pAffCode & "' " & vbCrLf & _
                              "     and a.supplierID = '" & Trim(psuppCode) & "' " & vbCrLf & _
                              " )tbl2 on tbl2.AffiliateName = tbl1.AffiliateName and tbl1.PartNo = tbl2.PartNo1 and tbl1.NoUrutDesc = tbl2.NoUrutDesc and tbl1.urutan = tbl2.urutan " & vbCrLf & _
                              "  Order By tbl1.PartNo, tbl1.urutan "
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

    'Private Function EmailToEmailCC() As DataSet
    '    Dim ls_SQL As String = ""

    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()
    '        'ls_SQL = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID='" & ls_value & "'"

    '        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
    '                 " select 'AFF' flag,affiliatepocc, affiliatepoto='',toEmail='' from ms_emailaffiliate where AffiliateID='" & Trim(txtAffiliateID.Text) & "'" & vbCrLf & _
    '                 " union all " & vbCrLf & _
    '                 " --PASI TO -CC " & vbCrLf & _
    '                 " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI where AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
    '                 " union all " & vbCrLf & _
    '                 " --Supplier TO- CC " & vbCrLf & _
    '                 " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail='' from ms_emailSupplier where SupplierID='" & Trim(txtSupplierCode.Text) & "'"

    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)
    '        If ds.Tables(0).Rows.Count > 0 Then
    '            Return ds
    '        End If
    '    End Using
    'End Function

#End Region
End Class