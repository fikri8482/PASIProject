Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing

Public Class AffReceivingConfExport
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
#End Region

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'AFFILIATE
        ls_sql = "SELECT distinct AffiliateID = '" & clsGlobal.gs_All & "', AffiliateName = '" & clsGlobal.gs_All & "' from MS_Affiliate " & vbCrLf & _
                 "UNION Select AffiliateID = RTRIM(AffiliateID) ,AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 70
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtaffiliate.Text = clsGlobal.gs_All
                .TextField = "AffiliateID"
                .DataBind()
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If checkbox1.Checked = True Then
                ls_Filter = ls_Filter + "AND  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(DSM.DeliveryDate,'')), 106) BETWEEN '" & dtDeliveryDateFrom.Value & "' AND '" & Format(dtDeliveryDateTo.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If rbInvoiceByPasi.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(IM.InvoiceNo, '') = '' " & vbCrLf
            ElseIf rbInvoiceByPasi.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(IM.InvoiceNo,'') <> '' " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_Filter = ls_Filter + " AND  ISNULL(DSM.SuratJalanNo, '') LIKE '%" & txtsj.Text & "%' " & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_Filter = ls_Filter + " AND IM.InvoiceDate BETWEEN '" & dtPasiInvoiceDateFrom.Value & "' AND '" & dtPasiInvoiceDateTo.Value & "' " & vbCrLf
            End If

            If cboaffiliate.Text <> "" Then
                If cboaffiliate.Text <> "== ALL ==" Then
                    ls_Filter = ls_Filter + " AND POM.AffiliateID = '" & cboaffiliate.Text & "' " & vbCrLf
                End If
            End If

            If txtorderno.Text <> "" Then
                ls_Filter = ls_Filter + " AND SD.OrderNo LIKE '%" & txtorderno.Text & "%'" & vbCrLf
            End If


            ls_SQL = " SELECT coldetail, coldetailname, act, colno = CONVERT(char,ROW_NUMBER() OVER(ORDER BY colorderno)), colperiod, colaffiliatecode, colaffiliatename, colorderno, colsupplierid, " & vbCrLf & _
                     " colsuppliername, colsuppplandeldate, coldeldate, colsj, colpasiinvno, colpasiinvdate, colshipping " & vbCrLf & _
                     " FROM ( " & vbCrLf & _
                     " SELECT DISTINCT " & vbCrLf & _
                     "         coldetailname = CASE WHEN ISNULL(IM.InvoiceNo,'') <> '' THEN 'DETAIL' ELSE 'INVOICE' END , " & vbCrLf & _
                     "         coldetail = 'InvToAffExport.aspx?prm=' + Rtrim(SD.OrderNo) + '|' + RTRIM(SM.AffiliateID)  + '|' + RTRIM(RM.SupplierID) + '|' + RTRIM(DSM.SuratJalanNo)+ '|' + RTRIM(ISNULL(IM.InvoiceNo,'')) + '|' + rtrim(isnull(SM.ShippingInstructionNo,'')) + '|' + rtrim(isnull(MA.AffiliateName,'')),  " & vbCrLf & _
                     "         act = '0'  , " & vbCrLf & _
                     "         colno = '' , " & vbCrLf & _
                     "         colperiod = RIGHT(CONVERT(CHAR(11), CONVERT(DATETIME, POM.Period), 106), " & vbCrLf & _
                     "                           8) , " & vbCrLf & _
                     "         colaffiliatecode = SM.AffiliateID , " & vbCrLf & _
                     "         colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
                     "         colorderno = SD.OrderNo , " & vbCrLf & _
                     "         colsupplierid = RM.SupplierID , "

            ls_SQL = ls_SQL + "         colsuppliername = MS.SupplierName , " & vbCrLf & _
                              "         colsuppplandeldate = CONVERT(CHAR(12), CONVERT(DATETIME, POM.ETDVendor), 106) , " & vbCrLf & _
                              "         coldeldate = CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(DSM.DeliveryDate, " & vbCrLf & _
                              "                                                               '')), 106) , " & vbCrLf & _
                              "         colsj = ISNULL(DSM.SuratJalanNo, '') , " & vbCrLf & _
                              "         colpasiinvno = isnull(IM.InvoiceNo,'') , " & vbCrLf & _
                              "         colpasiinvdate = CONVERT(CHAR(12), CONVERT(DATETIME, IM.InvoiceDate), 106), colshipping = isnull(SM.ShippingInstructionNo,'')  " & vbCrLf & _
                              " FROM    ShippingInstruction_Master SM " & vbCrLf & _
                              "         LEFT JOIN ShippingInstruction_Detail SD ON SM.affiliateID = SD.AffiliateID " & vbCrLf & _
                              "                                                    AND SM.forwarderID = SD.ForwarderID " & vbCrLf & _
                              "                                                    AND SM.ShippingInstructionNo = SD.ShippingInstructionNo "

            ls_SQL = ls_SQL + "         LEFT JOIN ReceiveForwarder_Master RM ON RM.AffiliateID = SM.AffiliateID " & vbCrLf & _
                              "                                                 AND RM.ForwarderID = SM.ForwarderID " & vbCrLf & _
                              "                                                 AND RM.OrderNo = SD.OrderNo " & vbCrLf & _
                              "                                                 AND RM.SuratjalanNo = SD.suratJalanNo " & vbCrLf & _
                              "         LEFT JOIN ReceiveForwarder_Detail RD ON RD.SuratJalanNo = RM.SuratJalanNo " & vbCrLf & _
                              "                                                 AND RD.SupplierID = RM.SupplierID " & vbCrLf & _
                              "                                                 AND RD.AffiliateID = RM.AffiliateID " & vbCrLf & _
                              "                                                 AND RD.OrderNo = SD.OrderNo " & vbCrLf & _
                              "                                                 AND RD.PartNo = SD.PartNo " & vbCrLf & _
                              "         LEFT JOIN DOSupplier_Master_Export DSM ON DSM.affiliateID = RM.AffiliateID " & vbCrLf & _
                              "                                                   AND DSM.orderNo = RM.OrderNo "

            ls_SQL = ls_SQL + "         LEFT JOIN ( SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo1 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                             ETAPort = ETAPort1 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory1 , " & vbCrLf & _
                              "                             week = 1 " & vbCrLf & _
                              "                     FROM    Po_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo2 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor2 , "

            ls_SQL = ls_SQL + "                             ETAPort = ETAPort2 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory2 , " & vbCrLf & _
                              "                             week = 2 " & vbCrLf & _
                              "                     FROM    Po_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo3 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                             ETAPort = ETAPort3 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory3 , " & vbCrLf & _
                              "                             week = 3 "

            ls_SQL = ls_SQL + "                     FROM    Po_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo4 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                             ETAPort = ETAPort4 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory4 , " & vbCrLf & _
                              "                             week = 4 " & vbCrLf & _
                              "                     FROM    Po_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , "

            ls_SQL = ls_SQL + "                             OrderNO = OrderNo5 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                             ETAPort = ETAPort5 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory5 , " & vbCrLf & _
                              "                             week = 5 " & vbCrLf & _
                              "                     FROM    Po_Master_Export " & vbCrLf & _
                              "                   ) POM ON POM.AffiliateID = DSM.AffiliateID " & vbCrLf & _
                              "                            AND POM.SupplierID = DSM.SupplierID " & vbCrLf & _
                              "                            AND POM.orderno = DSM.OrderNo " & vbCrLf & _
                              " 		LEFT JOIN InvoiceOverseas_Master IM ON IM.AffiliateID = SM.AffiliateID " & vbCrLf & _
                              " 												AND IM.ShippingInstructionNo = SM.ShippingInstructionNo "

            ls_SQL = ls_SQL + " 		LEFT JOIN InvoiceOverseas_Detail ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                              " 												AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                              " 												AND ID.ShippingInstructionNo = IM.ShippingInstructionNo " & vbCrLf & _
                              " 												AND ID.OrderNo = SD.OrderNo " & vbCrLf & _
                              " 												AND ID.Partno = SD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = SM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN MS_Supplier MS ON MS.SupplierID = RM.SupplierID "

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + " )x"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_fillcombo()
                lblerrmessage.Text = ""
                grid.JSProperties("cpdtfrom") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdtto") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpAll") = "ALL"

            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())

        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 5, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_AffiliateCode As String = ""
        Dim ls_supplierCode As String = ""
        Dim ls_PasiSj As String = ""
        Dim ls_PasiInvoiceno As String = ""
        Dim ls_orderno As String = ""
        Dim ls_shipping As String
        Dim ls_affliateName As String = ""

        Dim ls_Notes As String = ""

        With grid
            If e.UpdateValues.Count = 0 Then Exit Sub
            If (e.UpdateValues(0).NewValues("act").ToString()) = 1 Then
                'ls_DeliveryDate = Trim(e.UpdateValues(0).NewValues("colpasideliverydate").ToString())
                ls_AffiliateCode = Trim(e.UpdateValues(0).NewValues("colaffiliatecode").ToString())
                ls_affliateName = Trim(e.UpdateValues(0).NewValues("colaffiliatename").ToString())
                ls_supplierCode = Trim(e.UpdateValues(0).NewValues("colsupplierid").ToString())
                ls_orderno = Trim(e.UpdateValues(0).NewValues("colorderno").ToString())
                ls_PasiSj = "'" & Trim(e.UpdateValues(0).NewValues("colsj").ToString()) & "'"
                ls_shipping = "'" & Trim(e.UpdateValues(0).NewValues("colshipping").ToString()) & "'"

                If Trim(e.UpdateValues(0).NewValues("colpasiinvno").ToString()) <> "" Then
                    ls_PasiInvoiceno = "'" & Trim(e.UpdateValues(0).NewValues("colpasiinvno").ToString()) & "'"
                End If

                ls_Notes = ""
            End If

            If e.UpdateValues.Count > 1 Then
                For i = 1 To e.UpdateValues.Count - 1
                    If (e.UpdateValues(i).NewValues("act").ToString()) = 1 Then
                        ls_orderno = ls_orderno + ",'" & Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) & "'"
                        ls_PasiSj = ls_PasiSj + ",'" & Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        ls_Shipping = ls_shipping + ", '" & Trim(e.UpdateValues(i).NewValues("colshipping").ToString()) & "'"
                        If ls_PasiInvoiceno <> "" Then
                            ls_PasiInvoiceno = ls_PasiInvoiceno + ",'" & Trim(e.UpdateValues(i).NewValues("colpasiinvno").ToString()) & "'"
                        Else
                            If Trim(e.UpdateValues(i).NewValues("colpasiinvno").ToString()) <> "" Then
                                ls_PasiInvoiceno = ls_PasiInvoiceno + ",'" & Trim(e.UpdateValues(i).NewValues("colpasiinvno").ToString()) & "'"
                            End If
                        End If
                    End If
                Next
            End If
        End With
        Session("POListInv") = ls_orderno & "|" & ls_AffiliateCode & "|" & ls_supplierCode & _
                            "|" & ls_PasiSj & "|" & ls_PasiInvoiceno & _
                            "|" & ls_shipping & "|" & ls_affliateName

    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 5, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)

            Select Case pAction
                Case "gridload"
                    Call up_GridLoad()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "kosong"

            End Select

EndProcedure:
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If Not (e.DataColumn.FieldName = "coldetail" Or e.DataColumn.FieldName = "act") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        End If

        If e.DataColumn.FieldName = "act" Then
            If (e.GetValue("colorderno") = "") Then
                e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
                e.Cell.Controls("0").Controls.Clear()
            End If
        End If

        'If (e.GetValue("colkanbanno") = "" Or Left(e.GetValue("colpokanban"), 2) = "NO") Then
        '    e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        'End If

        If (e.DataColumn.FieldName = "coldetail") Then
            If (e.GetValue("colorderno") = "") Then
                e.Cell.Controls("0").Controls.Clear()
            End If
        End If
    End Sub


    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub btndeliver_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btndeliver.Click
        Response.Redirect("~/InvoiceExport/InvToAffExport.aspx")
    End Sub
End Class