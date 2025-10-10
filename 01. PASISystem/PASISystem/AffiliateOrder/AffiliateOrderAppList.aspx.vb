Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors

Public Class AffiliateOrderAppList
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance    
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "AFFILIATE ORDER APPROVAL LIST"
            up_Fillcombo()
            dtPeriodFrom.Value = Now
            dtPeriodTo.Value = Now
            If Session("M01Url") <> "" Then
                'Call bindData()

                Session.Remove("M01Url")
            End If
            lblInfo.Text = ""
        End If
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "DetailPage" Or e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "Period" Or e.Column.FieldName = "AffiliateID" _
            Or e.Column.FieldName = "AffiliateName" Or e.Column.FieldName = "PONo" Or e.Column.FieldName = "CommercialCls" _
            Or e.Column.FieldName = "SupplierID" Or e.Column.FieldName = "SupplierName" Or e.Column.FieldName = "ShipCls" _
            Or e.Column.FieldName = "CurrAff" Or e.Column.FieldName = "AmountAff" Or e.Column.FieldName = "CurrSupp" _
            Or e.Column.FieldName = "AmountSupp" Or e.Column.FieldName = "EntryDate" Or e.Column.FieldName = "EntryUser" _
            Or e.Column.FieldName = "POStatus1" Or e.Column.FieldName = "POStatus2" Or e.Column.FieldName = "POStatus3" _
            Or e.Column.FieldName = "POStatus4" Or e.Column.FieldName = "POStatus5" Or e.Column.FieldName = "POStatus6" _
            Or e.Column.FieldName = "POStatus7" Or e.Column.FieldName = "POStatus8") _
        And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Dim pDateFrom As Date = Split(e.Parameters, "|")(1)
            Dim pDateTo As Date = Split(e.Parameters, "|")(2)
            Dim pPONo As String = Split(e.Parameters, "|")(3)
            Dim pAffCode As String = Split(e.Parameters, "|")(4)
            Dim pSuppCode As String = Split(e.Parameters, "|")(5)
            Dim pSendTo As String = Split(e.Parameters, "|")(6)
            Dim pComm As String = Split(e.Parameters, "|")(7)
            Select Case pAction
                Case "load"
                    Call bindData(pDateFrom, pDateTo, pPONo, pAffCode, pSuppCode, pSendTo, pComm)
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
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

#End Region

#Region "PROCEDURE"
    Private Sub bindData(ByVal pDateFrom As Date, ByVal pDateTo As Date, ByVal pPONo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pPASIApp As String, ByVal pComm As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim pYear1 As String = "", pYear2 As String = ""
        Dim pMonth1 As String = "", pMonth2 As String = ""

        pYear1 = Year(dtPeriodFrom.Value)
        pYear2 = Year(dtPeriodTo.Value)

        pMonth1 = Month(dtPeriodFrom.Value)
        pMonth2 = Month(dtPeriodTo.Value)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  distinct  " & vbCrLf & _
                  "  	'DETAIL' DetailPage,  " & vbCrLf & _
                  "  	DeliveryByPASICls,  " & vbCrLf & _
                  "  	Period " & vbCrLf & _
                  "  	,a.AffiliateID,AffiliateName " & vbCrLf & _
                  "  	,RTRIM(a.PONo) PONo " & vbCrLf & _
                  "  	,RTRIM(a.PONo) + '-' + RTRIM(a.SupplierID) POMarking " & vbCrLf & _
                  "  	,case CommercialCls when '0' then 'NO' else 'YES' end CommercialCls " & vbCrLf & _
                  "  	,RTRIM(a.SupplierID) SupplierID " & vbCrLf & _
                  "  	,SupplierName " & vbCrLf & _
                  "  	,RTRIM(ShipCls) ShipCls " & vbCrLf 

            ls_SQL = ls_SQL + "  	,a.EntryDate " & vbCrLf & _
                              "  	,a.EntryUser,  	 " & vbCrLf & _
                              "  	case ISNULL(a.EntryDate,0) when 0 then 0 else 1 end POStatus1,  " & vbCrLf & _
                              "  	case ISNULL(AffiliateApproveDate,0) when 0 then 0 else 1 end POStatus2,  " & vbCrLf & _
                              "  	case ISNULL(PASISendAffiliateDate,0) when 0 then 0 else 1 end POStatus3,  " & vbCrLf & _
                              "  	case ISNULL(SupplierApproveDate,0) when 0 then 0 else 1 end POStatus4,  " & vbCrLf & _
                              "  	case ISNULL(SupplierApprovePendingDate,0) when 0 then 0 else 1 end POStatus5,  " & vbCrLf & _
                              "  	case ISNULL(SupplierUnApproveDate,0) when 0 then 0 else 1 end POStatus6,  " & vbCrLf

            ls_SQL = ls_SQL + "  	case ISNULL(PASIApproveDate,0) when 0 then 0 else 1 end POStatus7,  " & vbCrLf & _
                              "  	case ISNULL(FinalApproveDate,0) when 0 then 0 else 1 end POStatus8 " & vbCrLf & _
                              "  	,ISNULL(Remarks,'')Remarks " & vbCrLf & _
                              " from po_master a " & vbCrLf & _
                              " INNER JOIN dbo.PO_Detail b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.PONo     " & vbCrLf & _
                              "  LEFT JOIN dbo.Affiliate_Detail AD ON b.AffiliateID = AD.AffiliateID AND b.PONo = AD.PONo AND AD.SupplierID=a.SupplierID " & vbCrLf & _
                              "  LEFT JOIN dbo.MS_Affiliate MA ON a.AffiliateID = MA.AffiliateID " & vbCrLf & _
                              "  LEFT JOIN dbo.MS_Supplier MS ON a.SupplierID=ms.SupplierID " & vbCrLf & _
                              "  LEFT JOIN PO_MasterUpload d on d.PONo = b.PONo and d.AffiliateID = b.AffiliateID and d.SupplierID = b.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " WHERE Period BETWEEN '" & Format(pDateFrom, "yyyy-MM-dd") & "' AND '" & Format(pDateTo, "yyyy-MM-dd") & "' "

            If pPONo <> "" Then
                ls_SQL = ls_SQL + " AND a.PONo like '%" & pPONo & "%' " & vbCrLf
            End If

            If pAff <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND a.AffiliateID='" & pAff & "' " & vbCrLf
            End If

            If pSupp <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND b.SupplierID='" & pSupp & "' " & vbCrLf
            End If

            If pComm <> "2" Then
                ls_SQL = ls_SQL + " AND ISNULL(CommercialCls,0)='" & pComm & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " AND (ISNULL(SupplierApproveDate,'') <> '' OR ISNULL(SupplierApprovePendingDate,'') <> '' OR ISNULL(SupplierUnApproveDate,'') <> '')" & vbCrLf

            Select Case pPASIApp
                Case "0" 'No
                    ls_SQL = ls_SQL + " AND ISNULL(PASIApproveDate,'') = '' " & vbCrLf
                Case "1" 'Yes
                    ls_SQL = ls_SQL + " AND ISNULL(PASIApproveDate,'') <> '' " & vbCrLf
            End Select

            ls_SQL = ls_SQL + " GROUP BY DeliveryByPASICls, Period	,a.AffiliateID,AffiliateName,a.PONo	,CommercialCls,a.SupplierID	,SupplierName	,ShipCls " & vbCrLf & _                  
                  "  ,a.EntryDate, a.EntryUser, a.EntryDate, AffiliateApproveDate, 	PASISendAffiliateDate, SupplierApproveDate, SupplierApprovePendingDate " & vbCrLf & _
                  "  , SupplierUnApproveDate, PASIApproveDate, FinalApproveDate, Remarks  " & vbCrLf


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
        'Combo Supplier
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' SupplierCode, '" & clsGlobal.gs_All & "' SupplierName union all select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 50
                .Columns.Add("SupplierName")
                .Columns(1).Width = 120

                .TextField = "SupplierID"
                .DataBind()
                .SelectedIndex = 0
                txtSupplierName.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' NoUrut, '' as AffiliateID, '' AffiliateName, ''Address, ''City, '' PostalCode, ''Phone1, '' Phone2, ''Fax, ''NPWP, ''PODeliveryBy, ''DetailPage"

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
        GetAffiliateID = container.Grid.GetRowValues(container.ItemIndex, "AffiliateID")
    End Function

    Protected Function GetAffiliateName(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateName = container.Grid.GetRowValues(container.ItemIndex, "AffiliateName")
    End Function

    Protected Function GetPeriod(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPeriod = container.Grid.GetRowValues(container.ItemIndex, "Period")
    End Function

    Protected Function GetShipCls(ByVal container As GridViewDataItemTemplateContainer) As String
        GetShipCls = container.Grid.GetRowValues(container.ItemIndex, "ShipCls")
    End Function

    Protected Function GetCommercialCls(ByVal container As GridViewDataItemTemplateContainer) As String
        GetCommercialCls = container.Grid.GetRowValues(container.ItemIndex, "CommercialCls")
    End Function

    Protected Function GetSupplierID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetSupplierID = container.Grid.GetRowValues(container.ItemIndex, "SupplierID")
    End Function

    Protected Function GetRemarks(ByVal container As GridViewDataItemTemplateContainer) As String
        GetRemarks = container.Grid.GetRowValues(container.ItemIndex, "Remarks")
    End Function

    Protected Function GetFinalApproval(ByVal container As GridViewDataItemTemplateContainer) As String
        GetFinalApproval = container.Grid.GetRowValues(container.ItemIndex, "POStatus7")
    End Function

    Protected Function GetDeliveryBy(ByVal container As GridViewDataItemTemplateContainer) As String
        GetDeliveryBy = container.Grid.GetRowValues(container.ItemIndex, "DeliveryByPASICls")
    End Function

    Protected Function GetSupplierName(ByVal container As GridViewDataItemTemplateContainer) As String
        GetSupplierName = container.Grid.GetRowValues(container.ItemIndex, "SupplierName")
    End Function
#End Region

End Class