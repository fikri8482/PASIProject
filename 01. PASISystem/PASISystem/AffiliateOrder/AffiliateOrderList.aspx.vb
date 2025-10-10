Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors

Public Class AffiliateOrderList
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    'Dim menuID As String = "B01"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "AFFILIATE ORDER LIST"
            up_Fillcombo()
            dtPeriodFrom.Value = Now
            dtPeriodTo.Value = Now
            If Session("M01Url") <> "" Then
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
            Or e.Column.FieldName = "EntryDate" Or e.Column.FieldName = "EntryUser" _
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
    Private Sub bindData(ByVal pDateFrom As Date, ByVal pDateTo As Date, ByVal pPONo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pSend As String, ByVal pComm As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT DISTINCT " & vbCrLf & _
                          " 	row_number() over (order by POM.PONo) as NoUrut " & vbCrLf & _
                          " 	, Period  " & vbCrLf & _
                          " 	, POM.AffiliateID " & vbCrLf & _
                          " 	, AffiliateName " & vbCrLf & _
                          " 	, RTRIM(POM.PONo)PONo " & vbCrLf & _
                          " 	, CASE WHEN POD.KanbanCls = 0 then RTRIM(POM.PONo) + '-' + RTRIM(POM.SupplierID) ELSE POM.PONo END POMarking    " & vbCrLf & _
                          " 	, CASE WHEN ISNULL(POM.CommercialCls,0) = 0 THEN 'NO' ELSE 'YES' END CommercialCls    " & vbCrLf & _
                          " 	, SupplierID = ISNULL(POM.SupplierID,'') " & vbCrLf & _
                          " 	, SupplierName = ISNULL(SupplierName,'') " & vbCrLf & _
                          " 	, POM.ShipCls "

            ls_SQL = ls_SQL + " 	, POM.EntryDate " & vbCrLf & _
                              " 	, POM.EntryUser " & vbCrLf & _
                              " 	, CASE WHEN ISNULL(POM.EntryUser,'')  <> '' OR ISNULL(POM.EntryDate,'')  <> '' THEN 1 ELSE 0 END POStatus1    " & vbCrLf & _
                              " 	, CASE WHEN ISNULL(POM.AffiliateApproveUser,'')  <> '' OR ISNULL(POM.AffiliateApproveDate,'')  <> '' THEN 1 ELSE 0 END POStatus2    " & vbCrLf & _
                              " 	, CASE WHEN ISNULL(POM.PASISendAffiliateUser,'') <> '' OR ISNULL(POM.PASISendAffiliateDate,'') <> '' THEN 1 ELSE 0 END POStatus3    " & vbCrLf & _
                              " 	, CASE WHEN ISNULL(POM.SupplierApproveUser,'') <> '' OR ISNULL(POM.SupplierApproveDate,'') <> '' THEN 1 ELSE 0 END POStatus4      " & vbCrLf & _
                              " 	, CASE WHEN ISNULL(POM.SupplierApprovePendingUser,'') <> '' OR ISNULL(POM.SupplierApprovePendingDate,'') <> ''  THEN 1 ELSE 0 END POStatus5    " & vbCrLf & _
                              " 	, CASE WHEN ISNULL(POM.SupplierUnApproveUser,'') <> '' OR ISNULL(POM.SupplierUnApproveDate,'') <> ''  THEN 1 ELSE 0 END POStatus6    " & vbCrLf & _
                              " 	, CASE WHEN ISNULL(POM.PASIApproveUser,'') <> '' OR ISNULL(POM.PASIApproveDate,'') <> ''  THEN 1 ELSE 0 END POStatus7     " & vbCrLf & _
                              " 	, CASE WHEN ISNULL(POM.FinalApproveUser,'')<> '' OR ISNULL(POM.FinalApproveDate,'') <> '' THEN 1 ELSE 0 END POStatus8    " & vbCrLf & _
                              " 	, 'Detail' DetailPage    "

            ls_SQL = ls_SQL + " FROM dbo.PO_Master POM    " & vbCrLf & _
                              " 	INNER JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo and POM.AffiliateID = POD.AffiliateID and POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_Affiliate MAF ON POM.AffiliateID = MAF.AffiliateID    " & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_Supplier MSU ON POM.SupplierID = MSU.SupplierID    " & vbCrLf & _
                              " WHERE Period BETWEEN '" & Format(dtPeriodFrom.Value, "yyyy-MM") & "-01' AND '" & Format(dtPeriodTo.Value, "yyyy-MM-dd") & "' AND ISNULL(AffiliateApproveUser,'')  <> ''  " & vbCrLf

            If pPONo <> "" Then
                ls_SQL = ls_SQL + " AND POM.PONo like '%" & pPONo & "%' " & vbCrLf
            End If

            If pAff <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND POM.AffiliateID='" & pAff & "' " & vbCrLf
            End If

            If pSupp <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND POD.SupplierID='" & pSupp & "' " & vbCrLf
            End If

            If pComm <> "2" Then
                ls_SQL = ls_SQL + " AND ISNULL(CommercialCls,0)='" & pComm & "' " & vbCrLf
            End If

            Select Case pSend
                Case "0"
                    ls_SQL = ls_SQL + " AND ISNULL(PASISendAffiliateUser,'') = '' " & vbCrLf
                Case "1"
                    ls_SQL = ls_SQL + " AND ISNULL(PASISendAffiliateUser,'') <> '' " & vbCrLf
            End Select

            ls_SQL = ls_SQL + " AND ISNULL(AffiliateApproveUser,'')  <> '' " & vbCrLf & _
                  "  GROUP BY Period,POD.PONo,POM.PONo,POM.AffiliateID,AffiliateName,CommercialCls,POD.SupplierID,POD.KanbanCls,POM.SupplierID,SupplierName,ShipCls  " & vbCrLf & _
                  "   ,POM.EntryDate,POM.EntryUser,AffiliateApproveUser ,AffiliateApproveDate " & vbCrLf & _
                  "   ,PASISendAffiliateUser,PASISendAffiliateDate,SupplierApproveUser,SupplierApproveDate,SupplierApprovePendingUser,SupplierApprovePendingDate " & vbCrLf & _
                  "   ,SupplierUnApproveUser,SupplierUnApproveDate,PASIApproveUser,PASIApproveDate,FinalApproveUser,FinalApproveDate "

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

    'id
    Protected Function GetRowValue(ByVal container As GridViewDataItemTemplateContainer) As String
        If Not IsNothing(container.KeyValue) Then
            Return container.Grid.GetRowValuesByKeyValue(container.KeyValue, "PONo").ToString()
        Else
            Return Nothing
        End If
    End Function

    't3
    Protected Function GetPeriod(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPeriod = container.Grid.GetRowValues(container.ItemIndex, "Period")
    End Function

    't1
    Protected Function GetAffiliateID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateID = container.Grid.GetRowValues(container.ItemIndex, "AffiliateID")
    End Function

    't2
    Protected Function GetSupplierID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetSupplierID = container.Grid.GetRowValues(container.ItemIndex, "SupplierID")
    End Function
    
#End Region
End Class