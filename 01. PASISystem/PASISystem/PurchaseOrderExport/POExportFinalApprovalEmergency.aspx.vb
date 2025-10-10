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

Public Class POExportFinalApprovalEmergency
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
                Session("MenuDesc") = "PO FROM SUPPLIER APPROVE BY PASI (EMERGENCY)"

                If Session("PORevExportList") <> "" Then
                    param = Session("PORevExportList").ToString()
                ElseIf Session("TampungDelivery") <> "" Then
                    param = Session("TampungDelivery").ToString()
                Else
                    param = Request.QueryString("prm").ToString
                End If

                If param = "  'back'" Then
                    btnSubMenu.Text = "BACK"
                Else
                    If pStatus = False Then

                        pPeriod = Split(param, "|")(0)
                        pAffiliateCode = Split(param, "|")(1)
                        pAffiliateName = Split(param, "|")(2)
                        pSupplierCode = Split(param, "|")(3)
                        pSupplierName = Split(param, "|")(4)
                        pDeliveryCode = Split(param, "|")(5)
                        pDeliveryName = Split(param, "|")(6)
                        pCommercial = Split(param, "|")(7)
                        pPOEmergency = Split(param, "|")(8)
                        pShipBy = Split(param, "|")(9)
                        pRemarks = Split(param, "|")(10)
                        pPO = Split(param, "|")(11)

                        If pAffiliateCode <> "" Then btnSubMenu.Text = "BACK"
                        'If Trim(pPeriod) = "01 Jan 1900" Then pPeriod = Format(Now, "dd MMM yyyy")
                        'If Trim(pPeriod) = "" Then pPeriod = Format(Now, "dd MMM yyyy")

                        dtPeriodFrom.Value = pPeriod
                        rblCommercial.Value = pCommercial
                        cboAffiliateCode.Text = pAffiliateCode
                        txtAffiliateName.Text = pAffiliateName
                        cboDeliveryLoc.Text = pDeliveryCode
                        txtDeliveryLoc.Text = pDeliveryName
                        'cboSupplierCode.Text = pSupplierCode
                        'txtSupplierName.Text = pSupplierName
                        txtShipBy.Text = pShipBy
                        txtPOEmergency.Text = pPOEmergency
                        'txtRevisionNo.Text = pPORevNo

                        txtRemarks.Text = pRemarks
                        pStatus = True

                        Call bindDataHeader(pPOEmergency, pAffiliateCode, pPO)
                        Call bindDataDetail(pPOEmergency, pAffiliateCode, pPO)
                        'Call InitializeComponent(pPOEmergency, pAffiliateCode, pPO)
                        Session("pFilter") = pFilter
                        Session.Remove("POList")
                    End If
                End If
                btnSubMenu.Text = "BACK"
                'End If
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

        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/PurchaseOrderExport/POExportFinalApprovalList.aspx")
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        If x > grid.VisibleRowCount Then Exit Sub

        With grid
            If .VisibleRowCount > 0 Then
                If CDbl(e.GetValue("POQty")) <> CDbl(e.GetValue("POQtyOld")) Then
                    If e.DataColumn.FieldName = "POQty" Then
                        e.Cell.BackColor = Color.GreenYellow
                    End If
                End If
                If e.GetValue("ETDVendor1") <> e.GetValue("ETDVendor1Old") Then
                    If e.DataColumn.FieldName = "ETDVendor1" Then
                        e.Cell.BackColor = Color.GreenYellow
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Call uf_Approve()

                    If grid.VisibleRowCount = 0 Then
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
#End Region

#Region "PROCEDURE"
    Private Sub uf_Approve()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                ls_sql = " Update PO_Master_Export set PASIApproveDate = getdate(), PASIApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' and PONo = '" & Trim(pPO) & "' and SupplierID = '" & Trim(pSupplierCode) & "'" & vbCrLf

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
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' AffiliateID, '" & clsGlobal.gs_All & "' AffiliateName UNION ALL SELECT RTRIM(AffiliateID) AffiliateID,AffiliateName FROM dbo.MS_Affiliate" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
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
                txtAffiliateName.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataHeader(ByVal pPOEmergency As String, ByVal pAffCode As String, ByVal pPONO As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT  " & vbCrLf & _
                  " 	a.OrderNo1, a.OrderNo2, a.OrderNo3, a.OrderNo4, a.OrderNo5, " & vbCrLf & _
                  " 	a.ETDVendor1 ETDVendorOld1,b.ETDVendor1 ,a.ETDPort1 ETDPortOld1, a.ETAPort1 ETAPortOld1, a.ETAFactory1 ETAFactoryOld1, " & vbCrLf & _
                  " 	a.ETDVendor2 ETDVendorOld2,b.ETDVendor2 ,a.ETDPort2 ETDPortOld2, a.ETAPort2 ETAPortOld2, a.ETAFactory2 ETAFactoryOld2, " & vbCrLf & _
                  " 	a.ETDVendor3 ETDVendorOld3,b.ETDVendor3 ,a.ETDPort3 ETDPortOld3, a.ETAPort3 ETAPortOld3, a.ETAFactory3 ETAFactoryOld3, " & vbCrLf & _
                  " 	a.ETDVendor4 ETDVendorOld4,b.ETDVendor4 ,a.ETDPort4 ETDPortOld4, a.ETAPort4 ETAPortOld4, a.ETAFactory4 ETAFactoryOld4, " & vbCrLf & _
                  " 	a.ETDVendor5 ETDVendorOld5,b.ETDVendor5 ,a.ETDPort5 ETDPortOld5, a.ETAPort5 ETAPortOld5, a.ETAFactory5 ETAFactoryOld5 " & vbCrLf & _
                  " FROM PO_Master_Export a " & vbCrLf & _
                  " INNER JOIN PO_MasterUpload_Export b ON a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.ForwarderID = b.ForwarderID " & vbCrLf & _
                  " WHERE a.PONo = '" & pPONO & "' and a.AffiliateID = '" & pAffCode & "' and a.EmergencyCls = '" & pPOEmergency & "' " & vbCrLf & _
                  "  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtOrderNoWeek1.Text = ds.Tables(0).Rows(0)("OrderNo1") & ""
                dtWeekETDVendorOld1.Text = If(IsDBNull(ds.Tables(0).Rows(0)("ETDVendorOld1")), "", Format(ds.Tables(0).Rows(0)("ETDVendorOld1"), "yyyy-MM-dd"))
                dtWeekETDPortOld1.Text = If(IsDBNull(ds.Tables(0).Rows(0)("ETDPortOld1")), "", Format(ds.Tables(0).Rows(0)("ETDPortOld1"), "yyyy-MM-dd"))
                dtWeekETAPortOld1.Text = If(IsDBNull(ds.Tables(0).Rows(0)("ETAPortOld1")), "", Format(ds.Tables(0).Rows(0)("ETAPortOld1"), "yyyy-MM-dd"))
                dtETAFactWeekOld1.Text = If(IsDBNull(ds.Tables(0).Rows(0)("ETAFactoryOld1")), "", Format(ds.Tables(0).Rows(0)("ETAFactoryOld1"), "yyyy-MM-dd"))
                Call clsMsg.DisplayMessage(lblInfo, "1008", clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                Session("YA010IsSubmit") = lblInfo.Text
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataDetail(ByVal pPOEmergency As String, ByVal pAffCode As String, ByVal pPONO As String)
        Dim ls_SQL As String = ""
        Dim jsScript As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select " & vbCrLf & _
                      " 	convert(char,row_number() over (order by b.PartNo asc))as NoUrut, " & vbCrLf & _
                      " 	b.PartNo, d.PartName, e.Description UnitDesc, d.MOQ, d.QtyBox, " & vbCrLf & _
                      " 	b.Week1 POQty, b.Week1Old POQtyOld, a.ETDVendor1 ETDVendor1Old, c.ETDVendor1 " & vbCrLf & _
                      " from PO_Master_Export a   " & vbCrLf & _
                      " inner join PO_DetailUpload_Export b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                      " inner join PO_MasterUpload_Export c on a.PONo = c.PONo and a.SupplierID = c.SupplierID and a.AffiliateID = c.AffiliateID " & vbCrLf & _
                      " left join MS_Parts d on b.PartNo = d.PartNo " & vbCrLf & _
                      " left join MS_UnitCls e on e.UnitCls = d.UnitCls " & vbCrLf & _
                      " where a.PONo = '" & pPONO & "' and a.AffiliateID = '" & pAffCode & "'  " & vbCrLf & _
                      "  "

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