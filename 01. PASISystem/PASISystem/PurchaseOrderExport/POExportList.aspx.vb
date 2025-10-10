Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports OfficeOpenXml
Imports System.IO
Imports System.Drawing

Public Class POExportList
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
    Dim flag As Boolean = True
    Dim pFilter As String
    Dim pStatus As Boolean

    Dim ls_Period As String
    Dim ls_AffiliateCode As String = ""
    Dim ls_Order As String = ""
    Dim ls_Emergency As String
    Dim ls_Commercial As String
    Dim ls_Ship As String
    Dim ls_Error As String
    Dim ls_partno As String
    Dim ls_supplier As String

    Dim ls_pono As String = ""
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim param As String = ""
        Dim filterQty As String = ""

        Try
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then

                Session("MenuDesc") = "PO STATUS INQUIRY"

                If Session("POEList") <> "" Then

                    If param = "'back'" Then
                        btnSubMenu.Text = "BACK"
                    Else
                        If pStatus = False Then

                            pStatus = True
                            Call bindDataList()

                            Call up_FillCombo()
                            rdrCom1.Checked = True
                            rdrEAll.Checked = True
                            lblInfo.Text = ""

                            Session("pFilter") = pFilter
                            Session.Remove("M01Url")

                        End If
                    End If
                    btnSubMenu.Text = "BACK"
                    'End If
                Else
                    Call up_FillCombo()
                    rdrCom1.Checked = True
                    rdrEAll.Checked = True
                    lblInfo.Text = ""
                End If
            End If

            Session.Remove("POEList")
            Session.Remove("M01Url")

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblInfo.Text = ""
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            grid.JSProperties("cpMessage") = lblInfo.Text
        Finally

        End Try
    End Sub

    Private Sub grid_AfterPerformCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewAfterPerformCallbackEventArgs) Handles grid.AfterPerformCallback
        If e.CallbackName <> "CANCELEDIT" And e.CallbackName <> "CUSTOMCALLBACK" Then
            bindData()
        End If
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 0
        Dim pIsUpdate As Boolean
        Dim ls_PONo As String = "", ls_OrderNo As String = "", ls_AffiliateID As String = "", ls_SupplierID As String = "", ls_ForwarderID As String = ""
        Dim ls_StatusError As String = ""
        Dim ls_AdaData As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("PONo")

                If grid.VisibleRowCount = 0 Then
                    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager, False, False)
                    Exit Sub
                End If

                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    Session("AA220Msg") = lblInfo.Text
                    Exit Sub
                End If

                Dim a As Integer
                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1

                    ls_Active = (e.UpdateValues(iLoop).NewValues("cols").ToString())
                    If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

                    ls_AffiliateID = Trim(e.UpdateValues(iLoop).NewValues("AffiliateID").ToString())
                    ls_SupplierID = Trim(e.UpdateValues(iLoop).NewValues("SupplierID").ToString())
                    ls_ForwarderID = Trim(e.UpdateValues(iLoop).NewValues("ForwarderID").ToString())
                    ls_StatusError = Trim(e.UpdateValues(iLoop).NewValues("ErrorStatus").ToString())
                    ls_PONo = Trim(e.UpdateValues(iLoop).NewValues("PONo").ToString())
                    ls_OrderNo = Trim(e.UpdateValues(iLoop).NewValues("OrderNo").ToString())

                    Dim sqlstring As String
                    sqlstring = "SELECT * FROM PO_Master_Export WHERE PONo = '" & Trim(ls_PONo) & "' AND OrderNo1 = '" & Trim(ls_OrderNo) & "' AND AffiliateID = '" & Trim(ls_AffiliateID) & "' AND SupplierID = '" & Trim(ls_SupplierID) & "' AND ForwarderID = '" & Trim(ls_ForwarderID) & "' AND ErrorStatus = 'OK'"

                    Dim sqlComm As New SqlCommand(sqlstring, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If ls_Active = "1" Then
                        If pIsUpdate = False Then
                            ls_MsgID = "6015"
                            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                            Session("AA220Msg") = lblInfo.Text
                            Exit Sub
                        Else
                            ls_SQL = " 	UPDATE dbo.PO_Master_Export " & vbCrLf & _
                                     " 	   SET ExcelCls = '1' , " & vbCrLf & _
                                     " 	       PASISendToSupplierDate = GETDATE(), " & vbCrLf & _
                                     " 	       PASISendToSupplierUser = '" & Session("UserID").ToString & "', " & vbCrLf & _
                                     " 	       UpdateDate = GETDATE(), " & vbCrLf & _
                                     " 	       UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                     " 	 WHERE PONo = '" & Trim(ls_PONo) & "' AND OrderNo1 = '" & Trim(ls_OrderNo) & "' AND AffiliateID = '" & Trim(ls_AffiliateID) & "' AND SupplierID = '" & Trim(ls_SupplierID) & "' AND ForwarderID = '" & Trim(ls_ForwarderID) & "'"
                            ls_MsgID = "1008"

                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()
                        End If
                    ElseIf ls_Active = "0" And pIsUpdate = False Then
                        lblInfo.Text = "[ Please give a checkmark to save data ! ] "
                        Session("AA220Msg") = lblInfo.Text
                        Exit Sub
                    End If
                Next iLoop

                sqlTran.Commit()

            End Using

            sqlConn.Close()
        End Using

        'Call ColorGrid()
        'Session("POStatus") = "1"
        'Call bindData()
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        Session("AA220Msg") = lblInfo.Text
        lblInfo.Visible = True
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        'e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If e.GetValue("ErrorStatus") = "NG" Then
            e.Cell.BackColor = Color.Red
        End If
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "DetailPage" Or e.Column.FieldName = "RevisePage" Or e.Column.FieldName = "Period" _
            Or e.Column.FieldName = "AffiliateID" Or e.Column.FieldName = "OrderNo" Or e.Column.FieldName = "EmergencyCls" Or e.Column.FieldName = "CommercialCls" Or e.Column.FieldName = "ShipCls" _
            Or e.Column.FieldName = "SupplierID" Or e.Column.FieldName = "PartNo" _
            Or e.Column.FieldName = "EntryDate" Or e.Column.FieldName = "EntryUser" _
            Or e.Column.FieldName = "POStatus1" Or e.Column.FieldName = "POStatus2" Or e.Column.FieldName = "POStatus3" _
            Or e.Column.FieldName = "POStatus4" Or e.Column.FieldName = "POStatus5" Or e.Column.FieldName = "POStatus6") _
            And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session("M01Url") = "~/PurchaseOrderExport/POExportList.aspx"
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
        Session.Remove("Status2")
        Session.Remove("Status7")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        bindData()
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try

            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction

                Case "POStatus1"
                    Session("POStatus") = "1"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "POStatus2"
                    Session("POStatus") = "2"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "POStatus3"
                    Session("POStatus") = "3"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "POStatus4"
                    Session("POStatus") = "4"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "POStatus5"
                    Session("POStatus") = "5"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "POStatus6"
                    Session("POStatus") = "6"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "POStatus7"
                    Session("POStatus") = "7"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "load"
                    Session("POStatus") = "1"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If
                Case "kosong"
                    Call up_GridLoadWhenEventChange()

            End Select

EndProcedure:
            Session("AA220Msg") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub
#End Region

#Region "PROCEDURE"

    Private Sub ColorGrid()
        grid.VisibleColumns(1).CellStyle.BackColor = Color.White
    End Sub

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim pYear1 As String = "", pYear2 As String = ""
        Dim pMonth1 As String = "", pMonth2 As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  select ROW_NUMBER() over (order by PONo, OrderNo, SupplierID, AffiliateID) as NoUrut , * " & vbCrLf & _
                  "  from " & vbCrLf & _
                  "  ( " & vbCrLf & _
                  "  select  distinct  " & vbCrLf & _
                  "  	'DETAIL' DetailPage, coldetail = 'POExportEntryMonthly.aspx?prm='" & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.Period),'') + '|'   " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.Period),'') + '|'   " & vbCrLf & _
                  "  	+  RTRIM(ISNULL(a.CommercialCls,0)) + '|'   " & vbCrLf & _
                  "  	+  RTRIM(ISNULL(EmergencyCls,'E')) + '|'   " & vbCrLf & _
                  "  	+  RTRIM(ISNULL(ShipCls,0)) + '|'   " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.AffiliateID),'') + '|' +  ISNULL(RTRIM(c.AffiliateName),'') + '|'   " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.SupplierID),'') + '|' +  ISNULL(RTRIM(d.SupplierName),'') + '|'   " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.OrderNo1),'') + '|'   " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.ETDVendor1),'') + '|'  " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.ETDPort1),'') + '|'  " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.ETAPort1),'') + '|'  " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.ETAFactory1),'') + '|'  " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.ForwarderID),'') + '|' +  ISNULL(RTRIM(e.ForwarderName),'') + '|'  " & vbCrLf & _
                  "  	+  RTRIM(ISNULL(a.PONo,' ')) + '|' + ISNULL(RTRIM(c.ConsigneeCode),'') + '|' + ISNULL(RTRIM(a.SplitReffPONo),'') + '|' + ISNULL(RTRIM(a.SplitStatus),'0') " & vbCrLf & _
                  "  	, CASE WHEN a.PASIApproveDate is not null then '0' else '0' end cols , a.Period,a.AffiliateID,   " & vbCrLf

            ls_SQL = ls_SQL + "  	a.SupplierID, a.ForwarderID,  OrderNomor = a.PONo, " & vbCrLf & _
                              "  	PONo = a.Pono,  " & vbCrLf & _
                              " 	case PASISendToSupplierCls when '0' then 'NO' else 'YES' end PASISendToSupplierCls, " & vbCrLf & _
                              " 	case SupplierApprovalCls when '0' then 'NO' else 'YES' end SupplierApprovalCls, " & vbCrLf & _
                              "  	OrderNo = CASE WHEN ISNULL(EmergencyCls,'M') = 'M' THEN (ISNULL(RTRIM(a.OrderNo1),''))   " & vbCrLf & _
                              "  	          ELSE ISNULL(RTRIM(a.OrderNo1),'') END  " & vbCrLf & _
                              "  	,a.EmergencyCls, CASE WHEN a.CommercialCls = '1' then   'YES' else 'NO' END CommercialCls, a.ShipCls, ISNULL(a.ErrorStatus,'OK') ErrorStatus  " & vbCrLf

            If Session("POStatus") = "1" Then
                Session("GOTOStatus") = "1"
                ls_SQL = ls_SQL + " ,detailGOTO = 'GOTO'" & vbCrLf & _
                              "     ,CASE WHEN a.UploadDate is not null then CONVERT(CHAR(17), CONVERT(datetime,UploadDate),113) + ', ' + CONVERT(CHAR(3),UploadUser) else '' END POStatus1	  " & vbCrLf & _
                              "  	,GOTOPOStatus2 = 'POExportEntryMonthly.aspx?prm='" & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.Period),'') + '|'   " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.Period),'') + '|'   " & vbCrLf & _
                              "  	+  RTRIM(ISNULL(CommercialCls,0)) + '|'   " & vbCrLf & _
                              "  	+  RTRIM(ISNULL(EmergencyCls,'E')) + '|'   " & vbCrLf & _
                              "  	+  RTRIM(ISNULL(ShipCls,0)) + '|'   " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.AffiliateID),'') + '|' +  ISNULL(RTRIM(c.AffiliateName),'') + '|'   " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.SupplierID),'') + '|' +  ISNULL(RTRIM(d.SupplierName),'') + '|'   " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.OrderNo1),'') + '|'   " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.ETDVendor1),'') + '|'  " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.ETDPort1),'') + '|'  " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.ETAPort1),'') + '|'  " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.ETAFactory1),'') + '|'  " & vbCrLf & _
                              "  	+  ISNULL(RTRIM(a.ForwarderID),'') + '|' +  ISNULL(RTRIM(e.ForwarderName),'') + '|'  " & vbCrLf & _
                              "  	+  RTRIM(ISNULL(a.PONo,' ')) + '|' + ISNULL(RTRIM(c.ConsigneeCode),'') + '|' + ISNULL(RTRIM(a.SplitReffPONo),'') + '|' + ISNULL(RTRIM(a.SplitStatus),'0') " & vbCrLf & _
                              "  	,'' POStatus2	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierApproveUser) else '' END POStatus3	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierApprovePartialDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApprovePartialDate),113) + ', ' + CONVERT(CHAR(3),SupplierApprovePartialUser) else '' END POStatus4	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierUnApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierUnApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierUnApproveUser) else '' END POStatus5	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus6	  " & vbCrLf & _
                              "  	,CASE WHEN a.PASIApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASIApproveDate),113) + ', ' + CONVERT(CHAR(3),PASIApproveUser) else '' END POStatus6  " & vbCrLf & _
                              "  	,'' GOTOPOStatus7	  " & vbCrLf & _
                              "  	,'' POStatus7	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus8	  " & vbCrLf & _
                              "  	,'' POStatus8	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus9	  " & vbCrLf & _
                              "  	,'' POStatus9	  " & vbCrLf
            End If

            If Session("POStatus") = "2" Then
                Session("GOTOStatus") = "2"
                Session("Status2") = "Klik"
                ls_SQL = ls_SQL + " ,detailGOTO = 'GOTO'" & vbCrLf & _
                              "     ,CASE WHEN a.UploadDate is not null then CONVERT(CHAR(17), CONVERT(datetime,UploadDate),113) + ', ' + CONVERT(CHAR(3),UploadUser) else '' END POStatus1	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus2	  " & vbCrLf & _
                              "  	,CASE WHEN a.PASISendToSupplierDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASISendToSupplierDate),113) + ', ' + CONVERT(CHAR(3),PASISendToSupplierUser) else '' END POStatus2	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierApproveUser) else '' END POStatus3	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierApprovePartialDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApprovePartialDate),113) + ', ' + CONVERT(CHAR(3),SupplierApprovePartialUser) else '' END POStatus4	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierUnApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierUnApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierUnApproveUser) else '' END POStatus5	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus6	  " & vbCrLf & _
                              "  	,CASE WHEN a.PASIApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASIApproveDate),113) + ', ' + CONVERT(CHAR(3),PASIApproveUser) else '' END POStatus6  " & vbCrLf & _
                              "  	,'' GOTOPOStatus7	  " & vbCrLf & _
                              "  	,'' POStatus7	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus8	  " & vbCrLf & _
                              "  	,'' POStatus8	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus9	  " & vbCrLf & _
                              "  	,'' POStatus9	  " & vbCrLf
            End If

            If Session("POStatus") = "3" Then
                Session("GOTOStatus") = "3"
                ls_SQL = ls_SQL + " ,detailGOTO = 'GOTO' " & vbCrLf & _
                                  "      ,CASE WHEN a.UploadDate is not null then CONVERT(CHAR(17), CONVERT(datetime,Uploaddate),113) + ', ' + CONVERT(CHAR(3),UploadUser) else '' END POStatus1	   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus2	   " & vbCrLf & _
                                  "   	,CASE WHEN a.PASISendToSupplierDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASISendToSupplierDate),113) + ', ' + CONVERT(CHAR(3),PASISendToSupplierUser) else '' END POStatus2	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierApproveUser) else '' END POStatus3	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierApprovePartialDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApprovePartialDate),113) + ', ' + CONVERT(CHAR(3),SupplierApprovePartialUser) else '' END POStatus4	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierUnApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierUnApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierUnApproveUser) else '' END POStatus5	   " & vbCrLf & _
                                  "   	,GOTOPOStatus6 = 'POExportFinalApprovalMonthly.aspx?prm=' " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.Period),'') + '|' " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.SupplierID),'') + '|'    " & vbCrLf & _
                                  " 	+  RTRIM(ISNULL(CommercialCls,0)) + '|'    " & vbCrLf

                ls_SQL = ls_SQL + " 	+  RTRIM(ISNULL(EmergencyCls,'E')) + '|'    " & vbCrLf & _
                                  " 	+  RTRIM(ISNULL(ShipCls,0)) + '|'    " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.AffiliateID),'') + '|' +  ISNULL(RTRIM(c.AffiliateName),'') + '|' " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.ForwarderID),'') + '|' +  ISNULL(RTRIM(e.ForwarderName),'') + '|'  " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(pme.Remarks),'') + '|' " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.OrderNo1),'') + '|'       " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.ETDVendor1),'') + '|'   " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(pme.ETDVendor1),'') + '|'   " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.ETDPort1),'') + '|'   " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.ETAPort1),'') + '|'   " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(a.ETAFactory1),'') + '|' " & vbCrLf

                ls_SQL = ls_SQL + " 	+  RTRIM(ISNULL(a.PONo,' ')) + '|' " & vbCrLf & _
                                  " 	+  ISNULL(RTRIM(c.ConsigneeCode),'') + '|'   " & vbCrLf & _
                                  "   	+  ISNULL(RTRIM(a.SupplierID),'') + '|' +  ISNULL(RTRIM(d.SupplierName),'') " & vbCrLf & _
                                  "   	,CASE WHEN a.PASIApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASIApproveDate),113) + ', ' + CONVERT(CHAR(3),PASIApproveUser) else '' END POStatus6   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus7	   " & vbCrLf & _
                                  "   	,'' POStatus7	   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus8	   " & vbCrLf & _
                                  "   	,'' POStatus8	   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus9	   " & vbCrLf & _
                                  "   	,'' POStatus9	   " & vbCrLf
            End If

            If Session("POStatus") = "4" Then
                Session("GOTOStatus") = "4"
                ls_SQL = ls_SQL + " ,detailGOTO = 'GOTO' " & vbCrLf & _
                                  "     ,CASE WHEN a.UploadDate is not null then CONVERT(CHAR(17), CONVERT(datetime,Uploaddate),113) + ', ' + CONVERT(CHAR(3),UploadUser) else '' END POStatus1	   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus2	   " & vbCrLf & _
                                  "   	,CASE WHEN a.PASISendToSupplierDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASISendToSupplierDate),113) + ', ' + CONVERT(CHAR(3),PASISendToSupplierUser) else '' END POStatus2	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierApproveUser) else '' END POStatus3	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierApprovePartialDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApprovePartialDate),113) + ', ' + CONVERT(CHAR(3),SupplierApprovePartialUser) else '' END POStatus4	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierUnApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierUnApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierUnApproveUser) else '' END POStatus5	   " & vbCrLf & _
                                  "   	,GOTOPOStatus6 = 'FinalApprovalDetail.aspx?prm='" & vbCrLf & _
                                  "                         +  ISNULL(RTRIM(a.Period),'') + '|'   " & vbCrLf & _
                                  "                         +  ISNULL(RTRIM(a.PONo),'') + '|' " & vbCrLf & _
                                  "                         +  ISNULL(RTRIM(a.OrderNo1),'') + '|' " & vbCrLf & _
                                  "                         +  ISNULL(RTRIM(a.AffiliateID),'') + '|' " & vbCrLf & _
                                  "                         +  ISNULL(RTRIM(c.AffiliateName),'') + '|'" & vbCrLf & _
                                  "                         +  ISNULL(RTRIM(a.SupplierID),'') + '|' " & vbCrLf & _
                                  "                         +  ISNULL(RTRIM(d.SupplierName),'') " & vbCrLf & _
                                  "   	,CASE WHEN a.PASIApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASIApproveDate),113) + ', ' + CONVERT(CHAR(3),PASIApproveUser) else '' END POStatus6   " & vbCrLf & _
                                  "   	,GOTOPOStatus7 = '' --'~/DeliveryExport/DeliveryToAffListExport.aspx?prm=' " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.Period),'') + '|' " & vbCrLf

                ls_SQL = ls_SQL + "   	--+  ISNULL(RTRIM(a.AffiliateID),'') + '|' +  ISNULL(RTRIM(c.AffiliateName),'') + '|'    " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.SupplierID),'') + '|' +  ISNULL(RTRIM(d.SupplierName),'') + '|' " & vbCrLf & _
                                  "   	--+  RTRIM(ISNULL(a.PONo,' ')) + '|' + Isnull(Rtrim(c.ConsigneeCode),'') " & vbCrLf & _
                                  "   	,'' POStatus7	   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus8	   " & vbCrLf & _
                                  "   	,'' POStatus8	   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus9	   " & vbCrLf & _
                                  "   	,'' POStatus9	   " & vbCrLf
            End If

            If Session("POStatus") = "5" Then
                Session("GOTOStatus") = "5"
                ls_SQL = ls_SQL + " ,detailGOTO = 'GOTO' " & vbCrLf & _
                                    "     ,CASE WHEN a.UploadDate is not null then CONVERT(CHAR(17), CONVERT(datetime,Uploaddate),113) + ', ' + CONVERT(CHAR(3),UploadUser) else '' END POStatus1	   " & vbCrLf & _
                                    "   	,'' GOTOPOStatus2	   " & vbCrLf & _
                                    "   	,CASE WHEN a.PASISendToSupplierDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASISendToSupplierDate),113) + ', ' + CONVERT(CHAR(3),PASISendToSupplierUser) else '' END POStatus2	   " & vbCrLf & _
                                    "   	,CASE WHEN a.SupplierApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierApproveUser) else '' END POStatus3	   " & vbCrLf & _
                                    "   	,CASE WHEN a.SupplierApprovePartialDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApprovePartialDate),113) + ', ' + CONVERT(CHAR(3),SupplierApprovePartialUser) else '' END POStatus4	   " & vbCrLf & _
                                    "   	,CASE WHEN a.SupplierUnApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierUnApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierUnApproveUser) else '' END POStatus5	   " & vbCrLf & _
                                    "   	,'' GOTOPOStatus6	   " & vbCrLf & _
                                    "   	,CASE WHEN a.PASIApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASIApproveDate),113) + ', ' + CONVERT(CHAR(3),PASIApproveUser) else '' END POStatus6   " & vbCrLf & _
                                    "   	,'' GOTOPOStatus7	   " & vbCrLf & _
                                    "   	,CASE WHEN do.EntryDate is not null then CONVERT(CHAR(17), CONVERT(datetime,do.EntryDate),113) + ', ' + CONVERT(CHAR(3),do.EntryUser) else '' END POStatus7  " & vbCrLf

                ls_SQL = ls_SQL + "   	,GOTOPOStatus8 = '' --'~/DeliveryExport/DeliveryToAffListExport.aspx?prm=' " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.Period),'') + '|' " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.AffiliateID),'') + '|' +  ISNULL(RTRIM(c.AffiliateName),'') + '|' " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.SupplierID),'') + '|' +  ISNULL(RTRIM(d.SupplierName),'') + '|' " & vbCrLf & _
                                  "   	--+  RTRIM(ISNULL(a.PONo,' ')) + '|' + Isnull(Rtrim(c.ConsigneeCode),'')  " & vbCrLf & _
                                  "   	,'' POStatus8" & vbCrLf & _
                                  "   	,'' GOTOPOStatus9	   " & vbCrLf & _
                                  "   	,'' POStatus9	   " & vbCrLf
            End If

            If Session("POStatus") = "6" Then
                Session("GOTOStatus") = "6"
                ls_SQL = ls_SQL + " ,detailGOTO = 'GOTO' " & vbCrLf & _
                                  "     ,CASE WHEN a.UploadDate is not null then CONVERT(CHAR(17), CONVERT(datetime,Uploaddate),113) + ', ' + CONVERT(CHAR(3),UploadUser) else '' END POStatus1	   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus2	   " & vbCrLf & _
                                  "   	,CASE WHEN a.PASISendToSupplierDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASISendToSupplierDate),113) + ', ' + CONVERT(CHAR(3),PASISendToSupplierUser) else '' END POStatus2	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierApproveUser) else '' END POStatus3	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierApprovePartialDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApprovePartialDate),113) + ', ' + CONVERT(CHAR(3),SupplierApprovePartialUser) else '' END POStatus4	   " & vbCrLf & _
                                  "   	,CASE WHEN a.SupplierUnApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierUnApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierUnApproveUser) else '' END POStatus5	   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus6	   " & vbCrLf & _
                                  "   	,CASE WHEN a.PASIApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASIApproveDate),113) + ', ' + CONVERT(CHAR(3),PASIApproveUser) else '' END POStatus6   " & vbCrLf & _
                                  "   	,'' GOTOPOStatus7	   " & vbCrLf & _
                                  "      ,'' POStatus7  " & vbCrLf

                ls_SQL = ls_SQL + "   	,'' GOTOPOStatus8	   " & vbCrLf & _
                                  "   	,CASE WHEN rm.EntryDate is not null then CONVERT(CHAR(17), CONVERT(datetime,rm.EntryDate),113) + ', ' + CONVERT(CHAR(3),rm.EntryUser) else '' END POStatus8  " & vbCrLf & _
                                  "   	,GOTOPOStatus9 = '' --'~/ShippingInstruction/ShippingInstructionToForwarder.aspx?prm=' " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.Period),'') + '|' " & vbCrLf & _
                                  "   	--+  RTRIM('UPDATE') + '|'   " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(Convert(Char(16),a.ETDPort1,106)),'') + '|' " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.AffiliateID),'') + '|' +  ISNULL(RTRIM(c.AffiliateName),'') + '|' " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.ForwarderID),'') + '|' +  ISNULL(RTRIM(e.ForwarderName),'') + '|'     " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(sid.ShippingInstructionNo),'') + '|' " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(sim.ShippingInstructionDate),'') + '|'  " & vbCrLf & _
                                  "   	--+  ISNULL(RTRIM(a.SupplierID),'') + '|' +  ISNULL(RTRIM(d.SupplierName),'') + '|'       " & vbCrLf

                ls_SQL = ls_SQL + "   	--+  ISNULL(RTRIM(sid.PartNo),'') + '|' +  ISNULL(RTRIM(mp.PartName),'') + '|' " & vbCrLf & _
                                  "   	--+  RTRIM('ALREADY SEND') + '|' " & vbCrLf & _
                                  "   	--+  RTRIM(ISNULL(a.PONo,' ')) + '|' + Isnull(Rtrim(c.ConsigneeCode),'')	   " & vbCrLf & _
                                  "      ,'' POStatus9 " & vbCrLf
            End If

            If Session("POStatus") = "7" Then
                Session("GOTOStatus") = "7"
                Session("Status7") = "Klik"
                ls_SQL = ls_SQL + " ,detailGOTO = 'GOTO'" & vbCrLf & _
                              "     ,CASE WHEN a.UploadDate is not null then CONVERT(CHAR(17), CONVERT(datetime,Uploaddate),113) + ', ' + CONVERT(CHAR(3),UploadUser) else '' END POStatus1	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus2	  " & vbCrLf & _
                              "  	,CASE WHEN a.PASISendToSupplierDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASISendToSupplierDate),113) + ', ' + CONVERT(CHAR(3),PASISendToSupplierUser) else '' END POStatus2	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierApproveUser) else '' END POStatus3	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierApprovePartialDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierApprovePartialDate),113) + ', ' + CONVERT(CHAR(3),SupplierApprovePartialUser) else '' END POStatus4	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierUnApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,SupplierUnApproveDate),113) + ', ' + CONVERT(CHAR(3),SupplierUnApproveUser) else '' END POStatus5	  " & vbCrLf & _
                              "  	,'' GOTOPOStatus6	  " & vbCrLf & _
                              "  	,CASE WHEN a.PASIApproveDate is not null then CONVERT(CHAR(17), CONVERT(datetime,PASIApproveDate),113) + ', ' + CONVERT(CHAR(3),PASIApproveUser) else '' END POStatus6  " & vbCrLf & _
                              "  	,'' GOTOPOStatus7	  " & vbCrLf & _
                              "     ,'' POStatus7 " & vbCrLf & _
                              "  	,'' GOTOPOStatus8	  " & vbCrLf & _
                              "     ,'' POStatus8 " & vbCrLf & _
                              "  	,'' GOTOPOStatus9	  " & vbCrLf & _
                              "  	,CASE WHEN sid.EntryDate is not null then CONVERT(CHAR(2),(DAY(sid.EntryDate))) + CONVERT(CHAR(2),(MONTH(sid.EntryDate))) + RIGHT(CONVERT(CHAR(4),(YEAR(sid.EntryDate))),2) + ', ' + CONVERT(CHAR(5),sid.EntryDate,108) + ', ' + CONVERT(CHAR(3),sid.EntryUser) else '' END POStatus9 " & vbCrLf
            End If

            ls_SQL = ls_SQL + "  from PO_Master_Export a  "

            If Session("POStatus") = "1" Or Session("POStatus") = "2" Or Session("POStatus") = "4" Then
                ls_SQL = ls_SQL + "  inner join PO_Detail_Export b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.OrderNo1 = b.OrderNo1 " & vbCrLf
            End If

            If Session("POStatus") = "3" Then
                ls_SQL = ls_SQL + " inner join PO_Detail_Export b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.OrderNo1 = b.OrderNo1" & vbCrLf & _
                                  " inner join PO_MasterUpload_Export pme on a.PONo = pme.PONo and a.AffiliateID = pme.AffiliateID and a.SupplierID = pme.SupplierID and a.ForwarderID = pme.ForwarderID and a.OrderNo1 = pme.OrderNo1" & vbCrLf
            End If

            If Session("POStatus") = "4" Then
                ls_SQL = ls_SQL + " left join DOSupplier_Master_Export do on a.PONo = do.PONo and a.AffiliateID = do.AffiliateID and a.SupplierID = do.SupplierID and a.OrderNo1 = do.OrderNo " & vbCrLf
            End If

            If Session("POStatus") = "5" Then
                ls_SQL = ls_SQL + " inner join DOSupplier_Master_Export do on do.SupplierID = a.SupplierID and do.AffiliateID = a.AffiliateID and do.PONo = a.PONo and do.OrderNo = a.OrderNo1 " & vbCrLf & _
                                  " left join ReceiveForwarder_Master rm on rm.SupplierID = do.SupplierID and rm.AffiliateID = do.AffiliateID and rm.PONo = do.PONo and rm.OrderNo = do.OrderNo   " & vbCrLf
            End If

            If Session("POStatus") = "6" Then
                ls_SQL = ls_SQL + " inner join DOSupplier_Master_Export do on do.SupplierID = a.SupplierID and do.AffiliateID = a.AffiliateID and do.PONo = a.PONo and do.OrderNo = a.OrderNo1 " & vbCrLf & _
                                  " inner join ReceiveForwarder_Master rm on rm.SupplierID = a.SupplierID and rm.AffiliateID = a.AffiliateID and rm.PONo = a.PONo and rm.OrderNo = a.OrderNo1 " & vbCrLf & _
                                  " left join ShippingInstruction_Detail sid on sid.SupplierID = rm.SupplierID  and sid.AffiliateID = rm.AffiliateID and sid.OrderNo = a.PONo" & vbCrLf & _
                                  " left join ShippingInstruction_Master sim on sid.AffiliateID = sim.AffiliateID  and sid.ForwarderID = sim.ForwarderID and sid.ShippingInstructionNo = sim.ShippingInstructionNo" & vbCrLf & _
                                  " left join MS_Parts mp on sid.PartNo = mp.PartNo " & vbCrLf
            End If

            If Session("POStatus") = "7" Then
                ls_SQL = ls_SQL + " inner join ShippingInstruction_Detail sid on sid.SupplierID = a.SupplierID and sid.AffiliateID = a.AffiliateID and sid.OrderNo = a.PONo " & vbCrLf
            End If

            ls_SQL = ls_SQL + "  left join MS_Affiliate c on c.AffiliateID = a.AffiliateID  " & vbCrLf & _
                              "  left join MS_Supplier d on a.SupplierID = d.SupplierID  " & vbCrLf & _
                              "  left join MS_Forwarder e on a.ForwarderID = e.ForwarderID  " & vbCrLf & _
                              " WHERE 'A' = 'A' " & vbCrLf

            'If txtOrderNo.Text <> "" Then
            '    ls_SQL = ls_SQL + " AND a.OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            'End If

            If Session("POStatus") = "1" Then
                ls_SQL = ls_SQL + " AND UploadDate is Not Null AND PASISendToSupplierDate is Null " & vbCrLf
            End If

            If Session("POStatus") = "2" Then
                ls_SQL = ls_SQL + " AND PASISendToSupplierDate is Not Null " & vbCrLf & _
                    " AND SupplierApproveDate is Null AND SupplierApprovePartialDate is Null AND SupplierUnApproveDate is Null " & vbCrLf
            End If

            If Session("POStatus") = "3" Then
                ls_SQL = ls_SQL + " AND PASISendToSupplierDate is Not Null " & vbCrLf & _
                    " AND (SupplierApproveDate is Not Null OR SupplierApprovePartialDate is Not Null OR SupplierUnApproveDate is Not Null) " & vbCrLf & _
                    " AND PASIApproveDate is Null " & vbCrLf
            End If

            If Session("POStatus") = "4" Then
                ls_SQL = ls_SQL + " AND PASISendToSupplierDate is Not Null " & vbCrLf & _
                    " AND (SupplierApproveDate is Not Null OR SupplierApprovePartialDate is Not Null OR SupplierUnApproveDate is Not Null) " & vbCrLf & _
                    " AND PASIApproveDate is Not Null" & vbCrLf & _
                    " AND do.SuratJalanNo is Null " & vbCrLf
            End If

            If Session("POStatus") = "5" Then
                ls_SQL = ls_SQL + " AND PASISendToSupplierDate is Not Null " & vbCrLf & _
                    " AND (SupplierApproveDate is Not Null OR SupplierApprovePartialDate is Not Null OR SupplierUnApproveDate is Not Null) " & vbCrLf & _
                    " AND PASIApproveDate is Not Null" & vbCrLf & _
                    " AND do.SuratJalanNo is Not Null " & vbCrLf & _
                    " AND rm.SuratJalanNo is Null " & vbCrLf
            End If

            If Session("POStatus") = "6" Then
                ls_SQL = ls_SQL + " AND PASISendToSupplierDate is Not Null " & vbCrLf & _
                    " AND (SupplierApproveDate is Not Null OR SupplierApprovePartialDate is Not Null OR SupplierUnApproveDate is Not Null) " & vbCrLf & _
                    " AND PASIApproveDate is Not Null" & vbCrLf & _
                    " AND do.SuratJalanNo is Not Null " & vbCrLf & _
                    " AND rm.SuratJalanNo is Not Null " & vbCrLf & _
                    " AND sid.ShippingInstructionNo is Null " & vbCrLf
            End If

            ls_SQL = ls_SQL + "        )x " & vbCrLf & _
                              " WHERE 'A' = 'A' " & vbCrLf

            'If condition
            If cboAffiliate.Text.Trim <> "== ALL ==" Then
                ls_SQL = ls_SQL + " AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' "
            End If

            If cboSupplierCode.Text.Trim <> "== ALL ==" Then
                ls_SQL = ls_SQL + " AND SupplierID = '" & Trim(cboSupplierCode.Text) & "' "
            End If

            If rdrEM.Checked = True Then
                ls_SQL = ls_SQL + " AND EmergencyCls = 'M' "
            End If

            If rdrEE.Checked = True Then
                ls_SQL = ls_SQL + " AND EmergencyCls = 'E'"
            End If

            If rdrCom2.Checked = True Then
                ls_SQL = ls_SQL + " AND CommercialCls = 'YES'"
            End If

            If rdrCom3.Checked = True Then
                ls_SQL = ls_SQL + " AND CommercialCls = 'NO'"
            End If

            If txtOrderNo.Text <> "" Then
                'ls_SQL = ls_SQL + " AND PONO like '%" & Trim(txtOrderNo.Text) & "%'"
                ls_SQL = ls_SQL + " AND OrderNo like '%" & Trim(txtOrderNo.Text) & "%'"
            End If

            ls_SQL = ls_SQL + " ORDER BY PONo, OrderNo, SupplierID, AffiliateID "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()

            End With
            sqlConn.Close()

            'If Session("POStatus") = "5" Then
            '    grid.Columns("OrderNo").Visible = False
            'Else
            '    grid.Columns("OrderNo").Visible = True
            'End If

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  ''Nourut, '' DetailPage, '' Period, ''AffiliateID, ''OrderNo, ''EmergencyCls, ''CommercialCls, ''ShipCls, '' EntryDate, ''EntryUser, '' POStatus1, ''POStatus2, ''POStatus3, ''POStatus4, ''POStatus5, ''POStatus6, ''POStatus7, ''POStatus8, ''Remarks"

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

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        'Affiliate ID
        'ls_SQL = "select distinct RTRIM(a.AffiliateID) AffiliateCode, b.AffiliateName from MS_PartMapping a left join MS_Affiliate b on a.AffiliateID = b.AffiliateID order by AffiliateCode " & vbCrLf
        ls_SQL = " SELECT Distinct '== ALL ==' AffiliateCode, '== ALL ==' AffiliateName union all  " & vbCrLf & _
                    " select RTRIM(AffiliateID)AffiliateCode, isnull(AffiliateName,'') AffiliateName from MS_Affiliate " & vbCrLf & _
                    " where isnull(overseascls, '0') = '1' " & vbCrLf & _
                    " order by AffiliateCode " & vbCrLf & _
                    "  "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateCode")
                .Columns(0).Width = 75
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 400

                .TextField = "AffiliateCode"
                .DataBind()
                .SelectedIndex = 0
                txtAffiliate.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        'SupplierID
        ls_SQL = "SELECT [Supplier Code] = '== ALL ==' , [Supplier Name] = '== ALL ==' UNION ALL SELECT [Supplier Code] = RTRIM(supplierID) ,[Supplier Name] = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Supplier Code")
                .Columns(0).Width = 100
                .Columns.Add("Supplier Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                .TextField = "Supplier Code"
                .DataBind()
                txtSupplierName.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub bindDataList()
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  select row_number() over (order by SupplierID asc) as NoUrut, * " & vbCrLf & _
                  "   from  " & vbCrLf & _
                  "   (  " & vbCrLf & _
                  "   select  distinct   " & vbCrLf & _
                  "  	'0' cols , 'DETAIL' DetailPage, '' coldetail, a.AffiliateID, d.SupplierID, Left(a.Period, '7')Period,  " & vbCrLf & _
                  "  	'' PASISendToSupplierCls,  " & vbCrLf & _
                  "  	'' SupplierApprovalCls,  " & vbCrLf & _
                  "  	PONo = POno,  " & vbCrLf & _
                  "   	OrderNo = CASE WHEN ISNULL(a.EmergencyCls,'M') = 'M' THEN (ISNULL(RTRIM(a.OrderNo1),'') + ', ' + ISNULL(RTRIM(a.OrderNo2),'') + ', ' +  ISNULL(RTRIM(a.OrderNo3),'') + ', ' +  ISNULL(RTRIM(a.OrderNo4),'') + ', ' +  ISNULL(RTRIM(a.OrderNo5),''))    " & vbCrLf & _
                  "   	          ELSE ISNULL(RTRIM(a.OrderNo1),'') END   " & vbCrLf & _
                  "   	,a.EmergencyCls, CASE WHEN a.CommercialCls = '1' then   'YES' else 'NO' END CommercialCls, a.ShipCls,'OK' ErrorStatus   " & vbCrLf & _
                  "   	,'' POStatus1	   "

            ls_SQL = ls_SQL + "   	,'' POStatus2	   " & vbCrLf & _
                              "   	,'' POStatus3	   " & vbCrLf & _
                              "   	,'' POStatus4	   " & vbCrLf & _
                              "   	,'' POStatus5	   " & vbCrLf & _
                              "   	,'' POStatus6   " & vbCrLf & _
                              "  from UploadPOExport a   " & vbCrLf & _
                              "  left join MS_Parts b on a.PartNo = b.PartNo  " & vbCrLf & _
                              "  left join MS_UnitCls c on b.UnitCls = c.UnitCls  " & vbCrLf & _
                              "  left join MS_PartMapping d on a.AffiliateID = d.AffiliateID and a.PartNo = d.PartNo  " & vbCrLf & _
                              "  where a.AffiliateID = a.AffiliateID and a.SupplierID = a.SupplierID and a.PONo = a.PONo)x "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()

            End With
            sqlConn.Close()

            'clsGlobal.HideColumTanggal1(Session("Period"), grid)
        End Using
    End Sub

    Private Sub bindDataSend()
        'Dim lspon As String = ls_pono
        'Dim lsper As String = ls_period
        'Dim lsaff As String = ls_affiliateCode
        'Dim lsorder As String = ls_order
        'Dim lsem As String = ls_Emergency
        'Dim lscom As String = ls_Commercial
        'Dim lsship As String = ls_ship
        'Dim lssup As String = ls_supplier
        'Dim lserr As String = ls_error
        'Dim lspar As String = ls_PartNo

        'ls_pono = Trim(e.UpdateValues(0).NewValues("PONo").ToString())
        'ls_period = Trim(e.UpdateValues(0).NewValues("Period").ToString())
        'ls_affiliateCode = Trim(e.UpdateValues(0).NewValues("AffiliateID").ToString())
        'ls_order = Trim(e.UpdateValues(0).NewValues("OrderNo").ToString())
        'ls_Emergency = Trim(e.UpdateValues(0).NewValues("EmergencyCls").ToString())
        'ls_Commercial = Trim(e.UpdateValues(0).NewValues("CommercialCls").ToString())
        'ls_ship = Trim(e.UpdateValues(0).NewValues("ShipCls").ToString())

        'ls_error = "'" & Trim(e.UpdateValues(0).NewValues("ErrorStatus").ToString()) & "'"
        'ls_supplier = Trim(e.UpdateValues(0).NewValues("SupplierID").ToString())

        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim pYear1 As String = "", pYear2 As String = ""
        Dim pMonth1 As String = "", pMonth2 As String = ""

        Dim i As Integer
        'pYear1 = Year(dtPeriodFrom.Value)
        'pYear2 = Year(dtPeriodTo.Value)

        'pMonth1 = Month(dtPeriodFrom.Value)
        'pMonth2 = Month(dtPeriodTo.Value)

        If cboAffiliate.Text.Trim <> "== ALL ==" Then
            pWhere = pWhere + " and a.AffiliateID = '" & cboAffiliate.Text.Trim & "' "
        End If

        'If rdrAff2.Checked = True Then
        '    pWhere = pWhere + " and a.FinalApproveDate is not null "
        'End If

        'If rdrAff3.Checked = True Then
        '    pWhere = pWhere + " and a.FinalApproveDate is null "
        'End If

        If rdrEM.Checked = True Then
            pWhere = pWhere + " and a.EmergencyCls = 'M' "
        End If

        If rdrEE.Checked = True Then
            pWhere = pWhere + " and a.EmergencyCls = 'E' "
        End If

        If rdrCom2.Checked = True Then
            pWhere = pWhere + " and a.CommercialCls = '1' "
        End If

        If rdrCom3.Checked = True Then
            pWhere = pWhere + " and a.CommercialCls = '0' "
        End If

        'If rdrSup2.Checked = True Then
        '    pWhere = pWhere + " and a.SupplierApprovalCls = '1' "
        'End If

        'If rdrSup3.Checked = True Then
        '    pWhere = pWhere + " and a.SupplierApprovalCls = '0' "
        'End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  select ROW_NUMBER() over (order by PONo, SupplierID, AffiliateID) as NoUrut , * " & vbCrLf & _
                  "  from " & vbCrLf & _
                  "  ( " & vbCrLf & _
                  "  select  distinct  " & vbCrLf & _
                  "  	'DETAIL' DetailPage, coldetail = CASE WHEN ISNULL(EmergencyCls,'M') = 'M' THEN 'POExportEntryMonthly.aspx?prm=' ELSE 'POExportEntryEmergency.aspx?prm=' END  " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.Period),'') + '|'   " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.AffiliateID),'') + '|' +  ISNULL(RTRIM(c.AffiliateName),'') + '|'  " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.SupplierID),'') + '|' +  ISNULL(RTRIM(d.SupplierName),'') + '|'  " & vbCrLf & _
                  "  	+  ISNULL(RTRIM(a.ForwarderID),'') + '|' +  ISNULL(RTRIM(e.ForwarderName),'') + '|'  " & vbCrLf & _
                  "  	+ RTRIM(ISNULL(CommercialCls,0)) + '|' + RTRIM(ISNULL(EmergencyCls,'E')) + '|' + RTRIM(ISNULL(ShipCls,0)) + '|' + RTRIM(ISNULL(a.PONo,' ')) " & vbCrLf & _
                  "  	, CASE WHEN a.PASIApproveDate is not null then '1' else '0' end cols , a.Period,a.AffiliateID,   "

            ls_SQL = ls_SQL + "  	a.SupplierID,  " & vbCrLf & _
                              "  	a.PONo, b.PartNo,  " & vbCrLf & _
                              " 	case PASISendToSupplierCls when '0' then 'NO' else 'YES' end PASISendToSupplierCls, " & vbCrLf & _
                              " 	case SupplierApprovalCls when '0' then 'NO' else 'YES' end SupplierApprovalCls, " & vbCrLf & _
                              "  	OrderNo = CASE WHEN ISNULL(EmergencyCls,'M') = 'M' THEN (ISNULL(RTRIM(OrderNo1),'') + ', ' + ISNULL(RTRIM(OrderNo2),'') + ', ' +  ISNULL(RTRIM(OrderNo3),'') + ', ' +  ISNULL(RTRIM(OrderNo4),'') + ', ' +  ISNULL(RTRIM(OrderNo5),''))   " & vbCrLf & _
                              "  	          ELSE ISNULL(RTRIM(OrderNo1),'') END  " & vbCrLf & _
                              "  	,a.EmergencyCls, CASE WHEN a.CommercialCls = '1' then   'YES' else 'NO' END CommercialCls, a.ShipCls, ISNULL(a.ErrorStatus,'OK') ErrorStatus  " & vbCrLf & _
                              "  	,CASE WHEN a.UploadDate is not null then CONVERT(CHAR(2),(DAY(UploadDate))) + CONVERT(CHAR(2),(MONTH(UploadDate))) + RIGHT(CONVERT(CHAR(4),(YEAR(UploadDate))),2) + ', ' + CONVERT(CHAR(5),UploadDate,108) + ', ' + CONVERT(CHAR(3),UploadUser) else '' END POStatus1	  " & vbCrLf & _
                              "  	,CASE WHEN a.PASISendToSupplierDate is not null then CONVERT(CHAR(2),(DAY(PASISendToSupplierDate))) + CONVERT(CHAR(2),(MONTH(PASISendToSupplierDate))) + RIGHT(CONVERT(CHAR(4),(YEAR(PASISendToSupplierDate))),2) + ', ' + CONVERT(CHAR(5),PASISendToSupplierDate,108) + ', ' + CONVERT(CHAR(3),PASISendToSupplierUser) else '' END POStatus2	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierApproveDate is not null then CONVERT(CHAR(2),(DAY(SupplierApproveDate))) + CONVERT(CHAR(2),(MONTH(SupplierApproveDate))) + RIGHT(CONVERT(CHAR(4),(YEAR(SupplierApproveDate))),2) + ', ' + CONVERT(CHAR(5),SupplierApproveDate,108) + ', ' + CONVERT(CHAR(3),SupplierApproveUser) else '' END POStatus3	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierApprovePartialDate is not null then CONVERT(CHAR(2),(DAY(SupplierApprovePartialDate))) + CONVERT(CHAR(2),(MONTH(SupplierApprovePartialDate))) + RIGHT(CONVERT(CHAR(4),(YEAR(SupplierApprovePartialDate))),2) + ', ' + CONVERT(CHAR(5),SupplierApprovePartialDate,108) + ', ' + CONVERT(CHAR(3),SupplierApprovePartialUser) else '' END POStatus4	  " & vbCrLf & _
                              "  	,CASE WHEN a.SupplierUnApproveDate is not null then CONVERT(CHAR(2),(DAY(SupplierUnApproveDate))) + CONVERT(CHAR(2),(MONTH(SupplierUnApproveDate))) + RIGHT(CONVERT(CHAR(4),(YEAR(SupplierUnApproveDate))),2) + ', ' + CONVERT(CHAR(5),SupplierUnApproveDate,108) + ', ' + CONVERT(CHAR(3),SupplierUnApproveUser) else '' END POStatus5	  " & vbCrLf & _
                              "  	,CASE WHEN a.PASIApproveDate is not null then CONVERT(CHAR(2),(DAY(PASIApproveDate))) + CONVERT(CHAR(2),(MONTH(PASIApproveDate))) + RIGHT(CONVERT(CHAR(4),(YEAR(PASIApproveDate))),2) + ', ' + CONVERT(CHAR(5),PASIApproveDate,108) + ', ' + CONVERT(CHAR(3),PASIApproveUser) else '' END POStatus6  "

            ls_SQL = ls_SQL + "  from PO_Master_Export a  " & vbCrLf & _
                              "  inner join PO_Detail_Export b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID  " & vbCrLf & _
                              "  left join MS_Affiliate c on c.AffiliateID = a.AffiliateID  " & vbCrLf & _
                              "  left join MS_Supplier d on a.SupplierID = d.SupplierID  " & vbCrLf & _
                              "  left join MS_Forwarder e on a.ForwarderID = e.ForwarderID  " & vbCrLf & _
                              "  --left join PO_MasterUpload_Export f on a.PONo = f.PONo and a.AffiliateID = f.AffiliateID and a.SupplierID = f.SupplierID and a.ForwarderID = f.ForwarderID  " & vbCrLf & _
                              " WHERE (YEAR(Period) between '" & pYear1 & "' and '" & pYear2 & "') and (MONTH(Period) between '" & pMonth1 & "' and '" & pMonth2 & "') " & pWhere & " " & vbCrLf & _
                              "        )x "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            For i = 1 To ds.Tables(0).Rows.Count - 1
                ls_pono = ds.Tables(0).Rows(i)("PONo")
                ls_Period = ds.Tables(0).Rows(i)("Period")
                ls_AffiliateCode = ds.Tables(0).Rows(i)("AffiliateID")
                ls_Order = ds.Tables(0).Rows(i)("OrderNo")
                ls_Emergency = ds.Tables(0).Rows(i)("EmergencyCls")
                ls_Commercial = ds.Tables(0).Rows(i)("CommercialCls")
                ls_Ship = ds.Tables(0).Rows(i)("ShipCls")
                ls_supplier = ds.Tables(0).Rows(i)("SupplierID")
                ls_Error = ds.Tables(0).Rows(i)("ErrorStatus")
                ls_partno = ds.Tables(0).Rows(i)("PartNo")
            Next
            'With grid
            '    .DataSource = ds.Tables(0)
            '    .DataBind()
            '    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            'End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub grid_CustomColumnDisplayText(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewColumnDisplayTextEventArgs) Handles grid.CustomColumnDisplayText
        With e.Column
            If .FieldName = "GOTOPOStatus2" Then
                If e.GetFieldValue("GOTOPOStatus2") = "" Then
                    .Width = 0
                Else
                    .Width = 60
                End If
            End If

            If .FieldName = "POStatus2" Then
                If e.GetFieldValue("GOTOPOStatus2") <> "" And Session("Status2") <> "Klik" Then
                    .Width = 0
                End If
            End If

            If .FieldName = "GOTOPOStatus6" Then
                If e.GetFieldValue("GOTOPOStatus6") = "" Then
                    .Width = 0
                Else
                    .Width = 60
                End If
            End If

            If .FieldName = "POStatus6" Then
                If e.GetFieldValue("GOTOPOStatus6") <> "" Then
                    .Width = 0
                End If
            End If

            If .FieldName = "GOTOPOStatus7" Then
                If e.GetFieldValue("GOTOPOStatus7") = "" Then
                    .Width = 0
                Else
                    .Width = 60
                End If
            End If

            If .FieldName = "POStatus7" Then
                If e.GetFieldValue("GOTOPOStatus7") <> "" Then
                    .Width = 0
                End If
            End If

            If .FieldName = "GOTOPOStatus8" Then
                If e.GetFieldValue("GOTOPOStatus8") = "" Then
                    .Width = 0
                Else
                    .Width = 60
                End If
            End If

            If .FieldName = "POStatus8" Then
                If e.GetFieldValue("GOTOPOStatus8") <> "" Then
                    .Width = 0
                End If
            End If

            If .FieldName = "GOTOPOStatus9" Then
                If e.GetFieldValue("GOTOPOStatus9") = "" Then
                    .Width = 0
                Else
                    .Width = 60
                End If
            End If

            If .FieldName = "POStatus9" Then
                If e.GetFieldValue("GOTOPOStatus9") <> "" And Session("Status7") <> "Klik" Then
                    .Width = 0
                End If
            End If

        End With
    End Sub

#End Region
End Class