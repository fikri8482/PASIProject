Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors

Public Class PORevisionExportList
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
#End Region

#Region "CONTROL EVENTS"
    Private Sub btnSubMenu_Click(sender As Object, e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
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
            Or e.Column.FieldName = "PONo" Or e.Column.FieldName = "PORevNo" Or e.Column.FieldName = "CommercialCls" Or e.Column.FieldName = "ShipCls" _
            Or e.Column.FieldName = "EntryDate" Or e.Column.FieldName = "EntryUser" _
            Or e.Column.FieldName = "POStatus1" Or e.Column.FieldName = "POStatus2" Or e.Column.FieldName = "POStatus3" _
            Or e.Column.FieldName = "POStatus4" Or e.Column.FieldName = "POStatus5" Or e.Column.FieldName = "POStatus6") _
        And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            grid.JSProperties("cpMessage") = Session("AA220Msg")
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Dim pDateFrom As Date = Split(e.Parameters, "|")(1)
            Dim pDateTo As Date = Split(e.Parameters, "|")(2)
            Dim pAffCode As String = Split(e.Parameters, "|")(3)
            Dim pSendTo As String = Split(e.Parameters, "|")(4)
            Dim pMonthly As String = Split(e.Parameters, "|")(5)
            Dim pComm As String = Split(e.Parameters, "|")(6)
            Select Case pAction
                Case "load"
                    Call bindData(pDateFrom, pDateTo, pAffCode, pSendTo, pMonthly, pComm)
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
    Private Sub bindData(ByVal pDateFrom As Date, ByVal pDateTo As Date, ByVal pAff As String, ByVal pSend As String, ByVal pMonthly As String, ByVal pComm As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT coldetail = CASE WHEN ISNULL(EmergencyCls,'M') = 'M' THEN 'PORevisionExportMonthly.aspx?prm=' ELSE 'PORevisionExportEmergency.aspx?prm=' END   " & vbCrLf & _
                  "            +  ISNULL(RTRIM(POM.AffiliateID),'') + '|' +  ISNULL(RTRIM(MAF.AffiliateName),'') + '|' " & vbCrLf & _
                  "			   +  ISNULL(RTRIM(POM.SupplierID),'') + '|' +  ISNULL(RTRIM(MSU.SupplierName),'') + '|' " & vbCrLf & _
                  "			   +  ISNULL(RTRIM(POM.ForwarderID),'') + '|' +  ISNULL(RTRIM(MSF.ForwarderName),'') + '|' " & vbCrLf & _
                  "            + RTRIM(ISNULL(CommercialCls,0)) + '|' + RTRIM(ISNULL(EmergencyCls,'E')) + '|' + RTRIM(ISNULL(ShipCls,0)) " & vbCrLf & _
                  "    ,ROW_NUMBER() over (order by POM.PONo) as NoUrut,Period " & vbCrLf & _
                  "    ,POM.AffiliateID,AffiliateName,POM.SupplierID, MSU.SupplierName,RTRIM(POM.PONo)PONo,RTRIM(PORD.PORevNo)PORevNo  " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(CommercialCls,0) = 0 THEN 'NO' ELSE 'YES' END CommercialCls " & vbCrLf & _
                  "    ,ISNULL(EmergencyCls,'M') EmergencyCls, ShipCls, POM.EntryDate, POM.EntryUser   " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(POM.EntryUser,'')  <> '' OR ISNULL(POM.EntryDate,'')  <> '' THEN 1 ELSE 0 END POStatus1    " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(PASISendAffiliateUser,'') <> '' OR ISNULL(POM.PASISendAffiliateDate,'') <> '' THEN 1 ELSE 0 END POStatus2    " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(SupplierApproveUser,'') <> '' OR ISNULL(POM.SupplierApproveDate,'') <> '' THEN 1 ELSE 0 END POStatus3  " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(SupplierApprovePendingUser,'') <> '' OR ISNULL(POM.SupplierApprovePendingDate,'') <> ''  THEN 1 ELSE 0 END POStatus4 " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(SupplierUnApproveUser,'') <> '' OR ISNULL(POM.SupplierUnApproveDate,'') <> ''  THEN 1 ELSE 0 END POStatus5 " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(FinalApproveUser,'')<> '' OR ISNULL(POM.FinalApproveDate,'') <> '' THEN 1 ELSE 0 END POStatus6 " & vbCrLf & _
                  "    ,'Detail' DetailPage "

            ls_SQL = ls_SQL + "    FROM dbo.PO_Master_Export POM    " & vbCrLf & _
                              "    LEFT JOIN dbo.PO_Detail_Export POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID and POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                              "    LEFT JOIN dbo.PORev_Detail_Export PORD ON PORD.PONo = POD.PONo AND PORD.AffiliateID = POD.AffiliateID and PORD.SupplierID = POD.SupplierID  " & vbCrLf & _
                              "    LEFT JOIN dbo.MS_Affiliate MAF ON POM.AffiliateID = MAF.AffiliateID    " & vbCrLf & _
                              "    LEFT JOIN dbo.MS_Supplier MSU ON POD.SupplierID = MSU.SupplierID    " & vbCrLf & _
                              "    LEFT JOIN dbo.MS_Forwarder MSF ON POM.ForwarderID = MSF.ForwarderID " & vbCrLf

            ls_SQL = ls_SQL + " WHERE MONTH(Period) BETWEEN MONTH('" & Format(pDateFrom, "yyyy-MM-dd") & "') AND MONTH('" & Format(pDateTo, "yyyy-MM-dd") & "') " & vbCrLf & _
                              " AND YEAR(Period) BETWEEN YEAR('" & Format(pDateFrom, "yyyy-MM-dd") & "') AND YEAR('" & Format(pDateTo, "yyyy-MM-dd") & "') " & vbCrLf

            If pAff <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND POM.AffiliateID='" & pAff & "' " & vbCrLf
            End If

            If pComm <> "2" Then
                ls_SQL = ls_SQL + " AND ISNULL(CommercialCls,0)='" & pComm & "' " & vbCrLf
            End If

            If pMonthly <> "2" Then
                ls_SQL = ls_SQL + " AND ISNULL(EmergencyCls,'M')='" & pMonthly & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --AND ISNULL(FinalApproveUser,'')  <> '' " & vbCrLf & _
                              " GROUP BY Period,POD.PONo,POM.PONo,PORD.PORevNo,POM.ForwarderID,MSF.ForwarderName,POM.AffiliateID,AffiliateName,POM.SupplierID,SupplierName,EmergencyCls,CommercialCls,ShipCls   " & vbCrLf & _
                              " ,POM.EntryDate,POM.EntryUser,PASISendAffiliateUser,PASISendAffiliateDate,SupplierApproveUser,SupplierApproveDate,SupplierApprovePendingUser " & vbCrLf & _
                              " ,SupplierApprovePendingDate,SupplierUnApproveUser,SupplierUnApproveDate,FinalApproveUser,FinalApproveDate  "

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

    End Sub

    Protected Function GetRowValue(ByVal container As GridViewDataItemTemplateContainer) As String
        If Not IsNothing(container.KeyValue) Then
            Return container.Grid.GetRowValuesByKeyValue(container.KeyValue, "PONo")
        End If
    End Function

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

    't1
    Protected Function GetPeriod(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPeriod = container.Grid.GetRowValues(container.ItemIndex, "Period")
    End Function
    't2
    Protected Function GetAffiliateID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateID = container.Grid.GetRowValues(container.ItemIndex, "AffiliateID")
    End Function
    't3
    Protected Function GetSupplierID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetSupplierID = container.Grid.GetRowValues(container.ItemIndex, "SupplierID")
    End Function
    't4
    Protected Function GetPO(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPO = container.Grid.GetRowValues(container.ItemIndex, "PONo")
    End Function
    't5
    Protected Function GetDeliveryLocation(ByVal container As GridViewDataItemTemplateContainer) As String
        GetDeliveryLocation = container.Grid.GetRowValues(container.ItemIndex, "ForwarderID")
    End Function
    't6
    Protected Function GetRevNo(ByVal container As GridViewDataItemTemplateContainer) As String
        GetRevNo = container.Grid.GetRowValues(container.ItemIndex, "PORevNo")
    End Function
#End Region

End Class