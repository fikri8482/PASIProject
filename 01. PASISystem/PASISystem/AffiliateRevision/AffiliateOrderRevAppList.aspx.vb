Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel

Public Class AffiliateOrderRevAppList
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            If Session("M01Url") <> "" Then
                'Call bindData()
                'AFFILIATE ORDER REV. APPROVAL LIST
                Session("MenuDesc") = "AFFILIATE ORDER REV. APPROVAL LIST"
                Session.Remove("M01Url")
            End If
            dtPeriodFrom.Value = Now
            dtPeriodTo.Value = Now
            up_FillCombo(dtPeriodFrom.Value)
            rblSendToSupp.SelectedIndex = 0
            rblCommercial.SelectedIndex = 0
            lblInfo.Text = ""
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "DetailPage" Or e.Column.FieldName = "Period" Or e.Column.FieldName = "AffiliateID" Or e.Column.FieldName = "AffiliateName" _
            Or e.Column.FieldName = "PORevNo" Or e.Column.FieldName = "PONo" Or e.Column.FieldName = "CommercialCls" Or e.Column.FieldName = "SupplierID" _
            Or e.Column.FieldName = "SupplierName" Or e.Column.FieldName = "ShipCls" Or e.Column.FieldName = "CurrAff" _
            Or e.Column.FieldName = "AmountAff" Or e.Column.FieldName = "EntryDate" Or e.Column.FieldName = "EntryUser" _
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

    'Private Sub btnADD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnADD.Click
    '    Session("M01Url") = "~/AffiliateRevision/AffiliateOrderRevList.aspx"
    '    Response.Redirect("~/AffiliateRevision/AffiliateOrderRevEntry.aspx")
    'End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        bindData(dtPeriodFrom.Value, dtPeriodTo.Value, Trim(cboPONoRev.Text), Trim(cboPONo.Text), Trim(cboAffiliateCode.Text), Trim(cboSupplierCode.Text), rblSendToSupp.Value, rblCommercial.Value)
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Dim pDateFrom As Date = Split(e.Parameters, "|")(1)
                    Dim pDateTo As Date = Split(e.Parameters, "|")(2)
                    Dim pPORevNo As String = Split(e.Parameters, "|")(3)
                    Dim pPONo As String = Split(e.Parameters, "|")(4)
                    Dim pAffCode As String = Split(e.Parameters, "|")(5)
                    Dim pSuppCode As String = Split(e.Parameters, "|")(6)
                    Dim pSendTo As String = Split(e.Parameters, "|")(7)
                    Dim pComm As String = Split(e.Parameters, "|")(8)
                    bindData(pDateFrom, pDateTo, pPORevNo, pPONo, pAffCode, pSuppCode, pSendTo, pComm)

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

    Private Sub cboPONo_Callback(sender As Object, e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboPONo.Callback
        If String.IsNullOrEmpty(e.Parameter) Then
            Return
        End If

        Dim ls_value As String = Split(e.Parameter, "|")(0)
        Dim ls_dateFrom As Date = FormatDateTime(Split(e.Parameter, "|")(1), DateFormat.ShortDate)
        Dim ls_dateTo As Date = FormatDateTime(Split(e.Parameter, "|")(2), DateFormat.ShortDate)
        Dim ls_PORev As String = Split(e.Parameter, "|")(3)

        Dim ls_sql As String = ""
        ls_sql = "select '" & clsGlobal.gs_All & "' PONo union all select RTRIM(a.PONo) PONo " & vbCrLf & _
                " from PORev_Master a inner join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                " WHERE (YEAR(a.Period) between YEAR('" & ls_dateFrom & "') and YEAR('" & ls_dateTo & "')) " & vbCrLf & _
                " and (MONTH(a.Period) between MONTH('" & ls_dateFrom & "') and MONTH('" & ls_dateTo & "')) " & vbCrLf & _
                " AND PORevNo='" & ls_PORev & "' order by PONo " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
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

    Private Sub cboAffiliateCode_Callback(sender As Object, e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboAffiliateCode.Callback
        If String.IsNullOrEmpty(e.Parameter) Then
            Return
        End If

        Dim ls_value As String = Split(e.Parameter, "|")(0)
        Dim ls_dateFrom As Date = FormatDateTime(Split(e.Parameter, "|")(1), DateFormat.ShortDate)
        Dim ls_dateTo As Date = FormatDateTime(Split(e.Parameter, "|")(2), DateFormat.ShortDate)
        Dim ls_PORev As String = Split(e.Parameter, "|")(3)
        Dim ls_PO As String = Split(e.Parameter, "|")(4)

        Dim ls_sql As String = ""
        ls_sql = " select '" & clsGlobal.gs_All & "' AffiliateID,'" & clsGlobal.gs_All & "' AffiliateName union all " & vbCrLf & _
                " select RTRIM(a.AffiliateID)AffiliateID ,AffiliateName " & vbCrLf & _
                " from PORev_Master a inner join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                " LEFT JOIN dbo.MS_Affiliate MA ON b.AffiliateID = MA.AffiliateID" & vbCrLf & _
                " LEFT JOIN dbo.MS_Supplier MS ON a.SupplierID = MS.SupplierID " & vbCrLf & _
                " WHERE (YEAR(a.Period) between YEAR('" & ls_dateFrom & "') and YEAR('" & ls_dateTo & "')) " & vbCrLf & _
                " and (MONTH(a.Period) between MONTH('" & ls_dateFrom & "') and MONTH('" & ls_dateTo & "')) " & vbCrLf & _
                " AND a.PORevNo='" & ls_PORev & "' " & vbCrLf & _
                " AND a.PONo='" & ls_PO & "' --ORDER BY a.AffiliateID " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
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
    Private Sub cboSupplierCode_Callback(sender As Object, e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboSupplierCode.Callback
        If String.IsNullOrEmpty(e.Parameter) Then
            Return
        End If

        Dim ls_value As String = Split(e.Parameter, "|")(0)
        Dim ls_dateFrom As Date = FormatDateTime(Split(e.Parameter, "|")(1), DateFormat.ShortDate)
        Dim ls_dateTo As Date = FormatDateTime(Split(e.Parameter, "|")(2), DateFormat.ShortDate)
        Dim ls_PORev As String = Split(e.Parameter, "|")(3)
        Dim ls_PO As String = Split(e.Parameter, "|")(4)

        Dim ls_sql As String = ""
        ls_sql = " select '" & clsGlobal.gs_All & "' SupplierID,'" & clsGlobal.gs_All & "' SupplierName union all " & vbCrLf & _
                " select a.SupplierID,SupplierName " & vbCrLf & _
                " from PORev_Master a inner join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                " LEFT JOIN dbo.MS_Affiliate MA ON b.AffiliateID = MA.AffiliateID" & vbCrLf & _
                " LEFT JOIN dbo.MS_Supplier MS ON a.SupplierID = MS.SupplierID " & vbCrLf & _
                " WHERE (YEAR(a.Period) between YEAR('" & ls_dateFrom & "') and YEAR('" & ls_dateTo & "')) " & vbCrLf & _
                " and (MONTH(a.Period) between MONTH('" & ls_dateFrom & "') and MONTH('" & ls_dateTo & "')) " & vbCrLf & _
                " AND a.PORevNo='" & ls_PORev & "' " & vbCrLf & _
                " AND a.PONo='" & ls_PO & "' --ORDER BY a.AffiliateID " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierID")
                .Columns(0).Width = 50
                .Columns.Add("SupplierName")
                .Columns(1).Width = 120

                .TextField = "SupplierID"
                .DataBind()
                .SelectedIndex = 0
                txtAffiliateName.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData(ByVal pDateFrom As Date, ByVal pDateTo As Date, ByVal pPORevNo As String, ByVal pPONo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pSend As String, ByVal pComm As String)
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

            ls_SQL = "    SELECT row_number() over (order by PORM.PONo) as NoUrut,PORM.Period,PORM.AffiliateID,AffiliateName,RTRIM(PORM.PORevNo)PORevNo,RTRIM(PORM.PONo)PONo    " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(POM.CommercialCls,0) = 0 THEN 'NO' ELSE 'YES' END CommercialCls    " & vbCrLf & _
                  "    ,PORD.SupplierID,SupplierName    " & vbCrLf & _
                  "    ,POM.ShipCls   " & vbCrLf & _
                  "    ,CASE WHEN KanbanCls = 0 then RTRIM(PORM.PONo) + '-' + RTRIM(PORM.SupplierID) ELSE PORM.PONo END POMarking   " & vbCrLf & _
                  "    ,KanbanCls " & vbCrLf & _
                  "    ,PORM.EntryDate,PORM.EntryUser    " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(PORM.EntryUser,'')  <> '' OR ISNULL(PORM.EntryDate,'')  <> '' THEN 1 ELSE 0 END POStatus1    " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(PORM.AffiliateApproveUser,'')  <> '' OR ISNULL(PORM.AffiliateApproveDate,'')  <> '' THEN 1 ELSE 0 END POStatus2    " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(PORM.PASISendAffiliateUser,'') <> '' OR ISNULL(PORM.PASISendAffiliateDate,'') <> '' THEN 1 ELSE 0 END POStatus3    " & vbCrLf & _
                  "    ,CASE WHEN ISNULL(PORM.SupplierApproveUser,'') <> '' OR ISNULL(PORM.SupplierApproveDate,'') <> '' THEN 1 ELSE 0 END POStatus4      "

            ls_SQL = ls_SQL + "    ,CASE WHEN ISNULL(PORM.SupplierApprovePendingUser,'') <> '' OR ISNULL(PORM.SupplierApprovePendingDate,'') <> ''  THEN 1 ELSE 0 END POStatus5    " & vbCrLf & _
                              "    ,CASE WHEN ISNULL(PORM.SupplierUnApproveUser,'') <> '' OR ISNULL(PORM.SupplierUnApproveDate,'') <> ''  THEN 1 ELSE 0 END POStatus6    " & vbCrLf & _
                              "    ,CASE WHEN ISNULL(PORM.PASIApproveUser,'') <> '' OR ISNULL(PORM.PASIApproveDate,'') <> ''  THEN 1 ELSE 0 END POStatus7     " & vbCrLf & _
                              "    ,CASE WHEN ISNULL(PORM.FinalApproveUser,'')<> '' OR ISNULL(PORM.FinalApproveDate,'') <> '' THEN 1 ELSE 0 END POStatus8    " & vbCrLf & _
                              "    ,ISNULL(Remarks,'')Remarks " & vbCrLf & _
                              "    ,'Detail' DetailPage    " & vbCrLf & _
                              "    FROM dbo.PORev_Master PORM    " & vbCrLf & _
                              "    INNER JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID    " & vbCrLf & _
                              "    LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "    INNER JOIN dbo.PO_Detail POD ON POM.AffiliateID = POD.AffiliateID AND POM.PONo = POD.PONo AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "    LEFT JOIN dbo.MS_Affiliate MAF ON PORM.AffiliateID = MAF.AffiliateID    " & vbCrLf & _
                              "    LEFT JOIN dbo.MS_Supplier MSU ON PORD.SupplierID = MSU.SupplierID    " & vbCrLf 

            ls_SQL = ls_SQL + "    LEFT JOIN dbo.PORev_MasterUpload d on d.PONo = PORD.PONo and d.AffiliateID = PORD.AffiliateID " & vbCrLf & _
                              "    WHERE MONTH(PORM.Period) BETWEEN MONTH('" & Format(pDateFrom, "yyyy-MM-dd") & "') AND MONTH('" & Format(pDateTo, "yyyy-MM-dd") & "') " & vbCrLf & _
                              "    AND YEAR(PORM.Period) BETWEEN YEAR('" & Format(pDateFrom, "yyyy-MM-dd") & "') AND YEAR('" & Format(pDateTo, "yyyy-MM-dd") & "') " & vbCrLf & _
                              "    AND (ISNULL(PORM.SupplierApproveDate,'') <> '' OR ISNULL(PORM.SupplierApprovePendingDate,'') <> '' OR ISNULL(PORM.SupplierUnApproveDate,'') <> '')  " & vbCrLf
            If pPORevNo <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND PORD.PORevNo= '" & pPORevNo & "' " & vbCrLf
            End If

            If pPONo <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND PORM.PONo = '" & pPONo & "' " & vbCrLf
            End If

            If pAff <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND PORM.AffiliateID='" & pAff & "' " & vbCrLf
            End If

            If pSupp <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND PORM.SupplierID='" & pSupp & "' " & vbCrLf
            End If

            If pComm <> "2" Then
                ls_SQL = ls_SQL + " AND ISNULL(CommercialCls,0)='" & pComm & "' " & vbCrLf
            End If

            Select Case pSend
                Case "0"
                    ls_SQL = ls_SQL + " AND ISNULL(PORM.PASISendAffiliateUser,'') = '' OR ISNULL(PORM.PASISendAffiliateDate,'') = NULL " & vbCrLf
                Case "1"
                    ls_SQL = ls_SQL + " AND ISNULL(PORM.PASISendAffiliateUser,'') <> '' OR ISNULL(PORM.PASISendAffiliateDate,'') <> NULL " & vbCrLf
            End Select
            ls_SQL = ls_SQL + "   GROUP BY PORM.Period,PORD.PONo,PORM.PORevNo,PORM.PONo,PORM.AffiliateID,AffiliateName,POM.CommercialCls,PORD.SupplierID,PORM.SupplierID,SupplierName,POM.ShipCls,Remarks,KanbanCls   " & vbCrLf & _
                              "    ,PORM.EntryDate,PORM.EntryUser,PORM.AffiliateApproveUser ,PORM.AffiliateApproveDate  " & vbCrLf & _
                              "    ,PORM.PASISendAffiliateUser,PORM.PASISendAffiliateDate,PORM.SupplierApproveUser,PORM.SupplierApproveDate,PORM.SupplierApprovePendingUser,PORM.SupplierApprovePendingDate  " & vbCrLf & _
                              "    ,PORM.SupplierUnApproveUser,PORM.SupplierUnApproveDate,PORM.PASIApproveUser,PORM.PASIApproveDate,PORM.FinalApproveUser,PORM.FinalApproveDate  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' DetailPage, '' as PORevNo, '' Period, ''PONo, ''CommercialCls, ''ShipCls, ''CurrAff, ''AmountAff, '' EntryDate, ''EntryUser, '' POStatus1, ''POStatus2, ''POStatus3, ''POStatus4, ''POStatus5, ''POStatus6, ''POStatus7, ''POStatus8"

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
        'Combo PONo
        ls_SQL = "select '" & clsGlobal.gs_All & "' PONo union all select RTRIM(a.PONo) PONo from PORev_Master a inner join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID where Year(b.Period) = '" & Year(pPeriod) & "' and month(b.Period) = '" & Month(pPeriod) & "' order by PONo " & vbCrLf
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

        ls_SQL = "select '" & clsGlobal.gs_All & "' PORevNo union all select RTRIM(PORevNo) PORevNo from PORev_Master a inner join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID where Year(b.Period) = '" & Year(pPeriod) & "' and month(b.Period) = '" & Month(pPeriod) & "' order by PORevNo " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            'Combo PONoRev
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPONoRev
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PORevNo")
                .Columns(0).Width = 180

                .TextField = "PORevNo"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using

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

    Protected Function GetRowValue(ByVal container As GridViewDataItemTemplateContainer) As String
        If Not IsNothing(container.KeyValue) Then
            Return container.Grid.GetRowValuesByKeyValue(container.KeyValue, "PORevNo").ToString()
        End If
    End Function

    Protected Function GetPeriod(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPeriod = container.Grid.GetRowValues(container.ItemIndex, "Period")
    End Function

    Protected Function GetPORevNo(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPORevNo = container.Grid.GetRowValues(container.ItemIndex, "PORevNo")
    End Function

    Protected Function GetPONo(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPONo = container.Grid.GetRowValues(container.ItemIndex, "PONo")
    End Function

    Protected Function GetCommercial(ByVal container As GridViewDataItemTemplateContainer) As String
        GetCommercial = container.Grid.GetRowValues(container.ItemIndex, "CommercialCls")
    End Function

    Protected Function GetAffiliateID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateID = container.Grid.GetRowValues(container.ItemIndex, "AffiliateID")
    End Function

    Protected Function GetAffiliateName(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateName = container.Grid.GetRowValues(container.ItemIndex, "AffiliateName")
    End Function

    Protected Function GetSupplierID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetSupplierID = container.Grid.GetRowValues(container.ItemIndex, "SupplierID")
    End Function

    Protected Function GetSupplierName(ByVal container As GridViewDataItemTemplateContainer) As String
        GetSupplierName = container.Grid.GetRowValues(container.ItemIndex, "SupplierName")
    End Function

    Protected Function GetKanban(ByVal container As GridViewDataItemTemplateContainer) As String
        GetKanban = container.Grid.GetRowValues(container.ItemIndex, "KanbanCls")
    End Function

    Protected Function GetRemarks(ByVal container As GridViewDataItemTemplateContainer) As String
        GetRemarks = container.Grid.GetRowValues(container.ItemIndex, "Remarks")
    End Function

#End Region

End Class