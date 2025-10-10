Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing
Imports OfficeOpenXml
Imports System.IO

Public Class SupplierPriceMaster
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim menuID As String = "A14"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ls_AllowDownload As String = clsGlobal.Auth_UserConfirm(Session("UserID"), menuID)
        Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
        Dim ls_AllowDelete As String = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "PURCHASE PRICE MASTER"
            up_FillCombo()
            DeleteHistory()
            dt1.Value = Now
            dt2.Value = Now
            dt3.Value = Now
            dt4.Value = Now
            dt5.Value = Now
            dt6.Value = Now

            'Call bindData()
            ScriptManager.RegisterStartupScript(grid, grid.GetType(), "init", "grid.SetFocusedRowIndex(-1);", True)
            
            lblInfo.Text = ""            
        End If

        If ls_AllowDownload = False Then btnDownload.Enabled = False
        If ls_AllowUpdate = False Then btnUpload.Enabled = False
        If ls_AllowUpdate = False Then btnSubmit.Enabled = False
        If ls_AllowDelete = False Then btnDelete.Enabled = False

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName" _
            Or e.Column.FieldName = "SupplierID" Or e.Column.FieldName = "SupplierName" Or e.Column.FieldName = "StartDate" _
            Or e.Column.FieldName = "EndDate" Or e.Column.FieldName = "EntryDate" Or e.Column.FieldName = "CurrCls" _
            Or e.Column.FieldName = "Price" Or e.Column.FieldName = "PriceDesc" Or e.Column.FieldName = "PackingDesc") _
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

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        bindData()
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Response.Redirect("~/Upload/UploadSupplierPrice.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim lb_IsUpdate As Boolean
                    Call SaveData(lb_IsUpdate, _
                                     Split(e.Parameters, "|")(2), _
                                     Split(e.Parameters, "|")(3), _
                                     Split(e.Parameters, "|")(4), _
                                     Split(e.Parameters, "|")(5), _
                                     Split(e.Parameters, "|")(6), _
                                     Split(e.Parameters, "|")(7), _
                                     Split(e.Parameters, "|")(8), _
                                     Split(e.Parameters, "|")(9), _
                                     Split(e.Parameters, "|")(10), _
                                     Split(e.Parameters, "|")(11))
                Case "delete"
                    Dim pPartNo As String = Split(e.Parameters, "|")(1)
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim pStartDate As String = Split(e.Parameters, "|")(3)
                    Dim pCurrCls As String = Split(e.Parameters, "|")(4)
                    Dim pPackingCls As String = Split(e.Parameters, "|")(5)
                    Dim pDeliveryID As String = Split(e.Parameters, "|")(6)

                    If HF.Get("DeleteCls") = "0" Then
                        Call DeleteData(pAffiliateID, pPartNo, pCurrCls, pStartDate, pPackingCls, pDeliveryID)
                    Else
                        Call DeleteDataRec(pAffiliateID, pPartNo, pCurrCls, pStartDate, pPackingCls, pDeliveryID)
                    End If

                    Call bindData()
                Case "load"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    ElseIf grid.VisibleRowCount > 1 Then
                        grid.JSProperties("cpMessage") = ""
                        lblInfo.Text = ""
                    End If
                    grid.FocusedRowIndex = -1
                Case "loadaftersubmit"
                    Call bindData()
                    grid.FocusedRowIndex = -1
                    'txtMode.Text = "new"
                Case "yuhu"
                    Call up_kosong()
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim pSuppType As String = ""
                    Dim pAdditional As String = ""

                    If checkbox1.Checked = True Then
                        pAdditional = pAdditional + " AND StartDate = '" & Format(dt1.Value, "dd MMM yyyy") & "' " & vbCrLf
                    End If

                    If checkbox2.Checked = True Then
                        pAdditional = pAdditional + " AND EndDate = '" & Format(dt2.Value, "dd MMM yyyy") & "' " & vbCrLf
                    End If

                    If checkbox3.Checked = True Then
                        pAdditional = pAdditional + " AND EffectiveDate = '" & Format(dt3.Value, "dd MMM yyyy") & "' " & vbCrLf
                    End If

                    Dim dtProd As DataTable = clsMaster.GetTablePriceSupplier(cboPartNo.Text, cboSupplier.Text, pAdditional)
                    FileName = "TemplateMSSupplierPrice.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:3", psERR)
                    End If
            End Select

EndProcedure:
            Session("AA220Msg") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared        
        If grid.VisibleRowCount > 0 Then
            If e.GetValue("DeleteCls") = "1" Then
                e.Cell.BackColor = Color.Fuchsia
            End If
        End If
    End Sub

#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT row_number() over (order by PartNo, SupplierID) NoUrut, * FROM (" & vbCrLf & _
                        " select " & vbCrLf & _
                        " 	RTRIM(a.PartNo)PartNo, " & vbCrLf & _
                        " 	RTRIM(d.PartName)PartName, " & vbCrLf & _
                        " 	RTRIM(a.AffiliateID)SupplierID, " & vbCrLf & _
                        " 	RTRIM(c.SupplierName)SupplierName, " & vbCrLf & _
                        " 	RTRIM(a.PriceCls)PriceCls, " & vbCrLf & _
                        " 	RTRIM(g.Description)PriceDesc, " & vbCrLf & _
                        " 	RTRIM(a.PackingCls)PackingCls, " & vbCrLf & _
                        " 	RTRIM(f.Description)PackingDesc, " & vbCrLf & _
                        " 	RTRIM(a.DeliveryLocationID)DeliveryID, " & vbCrLf & _
                        " 	RTRIM(h.AffiliateName)DeliveryDesc, " & vbCrLf & _
                        " 	Convert(Char(15),a.StartDate,113)StartDate,  " & vbCrLf & _
                        " 	Convert(Char(15),a.EndDate,113)EndDate, " & vbCrLf & _
                        " 	Convert(Char(12),a.EffectiveDate,113)EffectiveDate, " & vbCrLf & _
                        " 	Rtrim(e.Description)CurrCls, " & vbCrLf & _
                        " 	a.price as Price, 0 DeleteCls, a.EntryDate, a.EntryUser, a.UpdateDate, a.UpdateUser " & vbCrLf & _
                        " from MS_Price a " & vbCrLf & _
                        " inner join MS_Supplier c on a.AffiliateID = c.SupplierID " & vbCrLf & _
                        " inner join MS_Parts d on a.PartNo = d.PartNo " & vbCrLf & _
                        " left join MS_CurrCls e on a.CurrCls = e.CurrCls " & vbCrLf & _
                        " left join MS_PackingCls f on a.PackingCls = f.PackingCls " & vbCrLf & _
                        " left join MS_PriceCls g on a.PriceCls = g.PriceCls " & vbCrLf & _
                        " left join (SELECT AffiliateID, AffiliateName FROM MS_Affiliate UNION ALL SELECT '0000' AffiliateID, 'COMMON' AffiliateName) h on a.DeliveryLocationID = h.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "UNION ALL" & vbCrLf & _
                        " select " & vbCrLf & _
                        " 	RTRIM(a.PartNo)PartNo, " & vbCrLf & _
                        " 	RTRIM(d.PartName)PartName, " & vbCrLf & _
                        " 	RTRIM(a.AffiliateID)SupplierID, " & vbCrLf & _
                        " 	RTRIM(c.SupplierName)SupplierName, " & vbCrLf & _
                        " 	RTRIM(a.PriceCls)PriceCls, " & vbCrLf & _
                        " 	RTRIM(g.Description)PriceDesc, " & vbCrLf & _
                        " 	RTRIM(a.PackingCls)PackingCls, " & vbCrLf & _
                        " 	RTRIM(f.Description)PackingDesc, " & vbCrLf & _
                        " 	RTRIM(a.DeliveryLocationID)DeliveryID, " & vbCrLf & _
                        " 	RTRIM(h.AffiliateName)DeliveryDesc, " & vbCrLf & _
                        " 	Convert(Char(15),a.StartDate,113)StartDate,  " & vbCrLf & _
                        " 	Convert(Char(15),a.EndDate,113)EndDate, " & vbCrLf & _
                        " 	Convert(Char(12),a.EffectiveDate,113)EffectiveDate, " & vbCrLf & _
                        " 	Rtrim(e.Description)CurrCls, " & vbCrLf & _
                        " 	a.price as Price, 1 DeleteCls, a.EntryDate, a.EntryUser, a.UpdateDate, a.UpdateUser " & vbCrLf & _
                        " from MS_Price_History a " & vbCrLf & _
                        " inner join MS_Supplier c on a.AffiliateID = c.SupplierID " & vbCrLf & _
                        " inner join MS_Parts d on a.PartNo = d.PartNo " & vbCrLf & _
                        " left join MS_CurrCls e on a.CurrCls = e.CurrCls " & vbCrLf & _
                        " left join MS_PackingCls f on a.PackingCls = f.PackingCls " & vbCrLf & _
                        " left join MS_PriceCls g on a.PriceCls = g.PriceCls " & vbCrLf & _
                        " left join (SELECT AffiliateID, AffiliateName FROM MS_Affiliate UNION ALL SELECT '0000' AffiliateID, 'COMMON' AffiliateName) h on a.DeliveryLocationID = h.AffiliateID " & vbCrLf & _
                        " )XYZ where 'a' = 'a'" & vbCrLf

            If cboSupplier.Text.Trim <> "" Then
                If cboSupplier.Text <> clsGlobal.gs_All Then
                    ls_SQL = ls_SQL + "and SupplierID = '" & cboSupplier.Text.Trim & "' "
                End If
            End If

            If cboPartNo.Text <> "" Then
                If cboPartNo.Text <> clsGlobal.gs_All Then
                    ls_SQL = ls_SQL + "and PartNo = '" & cboPartNo.Text & "' "
                End If
            End If


            If checkbox1.Checked = True Then
                ls_SQL = ls_SQL + " AND StartDate = '" & Format(dt1.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_SQL = ls_SQL + " AND EndDate = '" & Format(dt2.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If checkbox3.Checked = True Then
                ls_SQL = ls_SQL + " AND EffectiveDate = '" & Format(dt3.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If


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
    Private Sub up_kosong()

    End Sub
    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        grid.FocusedRowIndex = -1

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ''AffiliateID, '' AffiliateName, 
            ls_SQL = " select top 0  '' NoUrut, '' as PartNo, '' PartName, '' SupplierID, '' SupplierName, '' StartDate, '' EndDate, '' EffectiveDate, '' CurrCls, '' Price, '' PackingCls, ''PackingDesc, '' PriceCls, '' PriceDesc, '' DeliveryID, '' DeliveryDesc"

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

        'Person In Charge
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' SupplierCode, '" & clsGlobal.gs_All & "' SupplierName union all select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 75
                .Columns.Add("SupplierName")
                .Columns(1).Width = 400

                .TextField = "SupplierCode"
                .DataBind()
                .SelectedIndex = 0
                txtSupplier.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        'Person In Charge
        ls_SQL = "SELECT RTRIM(CurrCls) CurrCLs, RTRIM(Description) Description from MS_CurrCls order by CurrCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With CboCurrency
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("CurrCLs")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "CurrCLs"
                .DataBind()
                '.SelectedIndex = 0

            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplier2
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 75
                .Columns.Add("SupplierName")
                .Columns(1).Width = 400

                .TextField = "SupplierCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' PartCode, '" & clsGlobal.gs_All & "' PartName union all select RTRIM(PartNo) PartCode, PartName from MS_Parts order by PartCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartCode")
                .Columns(0).Width = 90
                .Columns.Add("PartName")
                .Columns(1).Width = 400

                .TextField = "PartCode"
                .DataBind()
                .SelectedIndex = 0
                txtPartNo.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(PartNo) PartCode, PartName from MS_Parts order by PartCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo2
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartCode")
                .Columns(0).Width = 90
                .Columns.Add("PartName")
                .Columns(1).Width = 400

                .TextField = "PartCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(PackingCls) [Packing Group], RTRIM(Description) [Packing Description] from MS_PackingCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPacking
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Packing Group")
                .Columns(0).Width = 75
                .Columns.Add("Packing Description")
                .Columns(1).Width = 400

                .TextField = "Packing Group"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(PriceCls) [Price Group], RTRIM(Description) [Price Description] from MS_PriceCls" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPriceCls
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Price Group")
                .Columns(0).Width = 75
                .Columns.Add("Price Description")
                .Columns(1).Width = 400

                .TextField = "Price Group"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT '0000' AffiliateID, 'COMMON' AffiliateName UNION ALL SELECT RTRIM(AffiliateID) AffiliateID, RTRIM(AffiliateName) AffiliateName FROM MS_Affiliate order by AffiliateID " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboDeliveryID
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 75
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 400

                .TextField = "AffiliateID"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub DeleteHistory()
        Dim ls_sql As String

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                    Dim sqlComm As New SqlCommand()

                    ls_sql = " delete MS_Price_History " & vbCrLf & _
                            " where exists " & vbCrLf & _
                            " (select * from MS_Price a where MS_Price_History.AffiliateID = a.AffiliateID and MS_Price_History.PartNo = a.PartNo " & vbCrLf & _
                            "   and MS_Price_History.StartDate = a.StartDate and MS_Price_History.PackingCls = a.PackingCls and MS_Price_History.CurrCls = a.CurrCls and MS_Price_History.DeliveryLocationID  = a.DeliveryLocationID ) "

                    sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    sqlComm.Dispose()
                    sqlTran.Commit()
                End Using

                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub DeleteData(ByVal pAffiliateID As String, ByVal pPartNo As String, ByVal pCurrency As String, ByVal pStartDate As Date, ByVal pPackingCls As String, ByVal pdeliveryID As String)

        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Dim shostname As String = System.Net.Dns.GetHostName

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("MS_Price")
                    ls_SQL = " INSERT INTO MS_Price_HISTORY " & vbCrLf & _
                             " SELECT * FROM MS_Price " & vbCrLf & _
                             " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "'" & vbCrLf & _
                             " and CurrCls='" & pCurrency & "' and StartDate ='" & pStartDate & "' and PackingCls = '" & pPackingCls & "'" & vbCrLf & _
                             " and DeliveryLocationID='" & pdeliveryID & "' "
                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    ls_SQL = " UPDATE MS_Price_HISTORY SET UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID") & "' " & vbCrLf & _
                                " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "'" & vbCrLf & _
                                " and CurrCls='" & pCurrency & "' and StartDate ='" & pStartDate & "' and PackingCls = '" & pPackingCls & "'" & vbCrLf
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    ls_SQL = " DELETE MS_Price " & vbCrLf & _
                                " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "'" & vbCrLf & _
                                " and CurrCls='" & pCurrency & "' and StartDate ='" & pStartDate & "' and PackingCls = '" & pPackingCls & "'" & vbCrLf & _
                                " and DeliveryLocationID='" & pdeliveryID & "' "
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                             "VALUES ('" & shostname & "','" & menuID & "','D','" & pPartNo & "','Delete PartNo " & pPartNo & ", Affiliate " & pAffiliateID & ", Curr " & pCurrency & ", StartDate " & pStartDate & ", PackingCls " & pPackingCls & "', GETDATE(),'" & Session("UserID") & "')  "
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    SqlComm.Dispose()
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using
            If x > 0 Then
                Call clsMsg.DisplayMessage(lblInfo, "1003", clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                grid.JSProperties("cpType") = "info"
                grid.JSProperties("cpFunction") = "delete"
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub DeleteDataRec(ByVal pAffiliateID As String, ByVal pPartNo As String, ByVal pCurrency As String, ByVal pStartDate As Date, ByVal pPackingCls As String, ByVal pdeliveryID As String)
        'ByVal pSupplierID As String,
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Dim shostname As String = System.Net.Dns.GetHostName

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                    ls_SQL = " INSERT INTO MS_Price" & vbCrLf & _
                             " SELECT * FROM MS_Price_HISTORY " & vbCrLf & _
                             " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "'" & vbCrLf & _
                             " and CurrCls='" & pCurrency & "' and StartDate ='" & pStartDate & "' and PackingCls = '" & pPackingCls & "'" & vbCrLf & _
                             " and DeliveryLocationID='" & pdeliveryID & "' "
                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    ls_SQL = " UPDATE MS_Price SET UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID") & "' " & vbCrLf & _
                                " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "'" & vbCrLf & _
                                " and CurrCls='" & pCurrency & "' and StartDate ='" & pStartDate & "' and PackingCls = '" & pPackingCls & "'" & vbCrLf
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    ls_SQL = " DELETE MS_Price_HISTORY " & vbCrLf & _
                                " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "'" & vbCrLf & _
                                " and CurrCls='" & pCurrency & "' and StartDate ='" & pStartDate & "' and PackingCls = '" & pPackingCls & "'" & vbCrLf & _
                                " and DeliveryLocationID='" & pdeliveryID & "' "
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    'insert into history
                    ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                             "VALUES ('" & shostname & "','" & menuID & "','R','" & pPartNo & "','Recovery PartNo " & pPartNo & ", Affiliate " & pAffiliateID & ", Curr " & pCurrency & ", StartDate " & pStartDate & ", PackingCls " & pPackingCls & "', GETDATE(),'" & Session("UserID") & "')  "
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()
                    SqlComm.Dispose()
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using
            If x > 0 Then
                Call clsMsg.DisplayMessage(lblInfo, "1016", clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                grid.JSProperties("cpType") = "info"
                grid.JSProperties("cpFunction") = "delete"
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pPartNo As String = "", _
                         Optional ByVal pAffiliateID As String = "", _
                         Optional ByVal pStartDate As String = "", _
                         Optional ByVal pEndDate As String = "", _
                         Optional ByVal pEffectiveDate As String = "", _
                         Optional ByVal pCurrency As String = "", _
                         Optional ByVal pPrice As Double = 0, _
                         Optional ByVal pPackingCls As String = "", _
                         Optional ByVal pPriceCls As String = "", _
                         Optional ByVal pDeliveryID As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = "", ls_Sql2 As String = ""
        Dim admin As String = Session("UserID").ToString

        Dim pub_effdate As Date
        Dim pub_enddate As Date
        Dim pub_price As Double
        Dim pub_pricecls As String
        Dim pub_packingCls As String
        Dim pub_affcode As String
        Dim shostname As String = System.Net.Dns.GetHostName

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                'Dim tempDelivery As String = HF.Get("hfTest")

                ls_SQL = " SELECT PartNo, PackingCls, EffectiveDate, EndDate, Price, PriceCls, PackingCls FROM MS_Price " & vbCrLf & _
                            "WHERE PartNo = '" & pPartNo & "' " & vbCrLf & _
                            "AND AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                            "AND StartDate = '" & pStartDate & "' " & vbCrLf & _
                            "AND CurrCls = '" & pCurrency & "' " & vbCrLf & _
                            "AND PackingCls in ('" & pPackingCls & "','') " & vbCrLf & _
                            "AND DeliveryLocationID = '" & pDeliveryID & "' "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    pIsNewData = False
                    ls_Sql2 = "UPDATE MS_Price SET " & vbCrLf & _
                           "EndDate = '" & pEndDate & "'," & vbCrLf & _
                           "EffectiveDate = '" & pEffectiveDate & "'," & vbCrLf & _
                           "Price = '" & CDec(pPrice) & "'," & vbCrLf & _
                           "DeliveryLocationID = '" & pDeliveryID & "'," & vbCrLf & _
                           "PriceCls = '" & pPriceCls & "'," & vbCrLf & _
                           "PackingCls = '" & pPackingCls & "'," & vbCrLf & _
                           "UpdateDate = getdate()," & vbCrLf & _
                           "UpdateUser = '" & Session("UserID").ToString & "'" & vbCrLf & _
                           "WHERE PartNo = '" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "'" & vbCrLf & _
                           "and StartDate = '" & pStartDate & "' and CurrCls='" & pCurrency & "' and PackingCls = '" & pPackingCls & "' and DeliveryLocationID = '" & pDeliveryID & "' "
                    pub_effdate = Format(ds.Tables(0).Rows(0)("EffectiveDate"), "yyyy-MM-dd")
                    pub_enddate = Format(ds.Tables(0).Rows(0)("EndDate"), "yyyy-MM-dd")
                    pub_price = ds.Tables(0).Rows(0)("Price")
                    pub_pricecls = ds.Tables(0).Rows(0)("PriceCls").ToString.Trim
                    pub_packingCls = ds.Tables(0).Rows(0)("PackingCls")
                Else
                    pIsNewData = True
                End If
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_Price " & _
                                "(PartNo, AffiliateID, StartDate, EndDate, EffectiveDate, EntryDate, EntryUser, CurrCls, Price, PackingCls, PriceCls, DeliveryLocationID)" & _
                                " VALUES ('" & pPartNo & "','" & pAffiliateID & "','" & pStartDate & "','" & pEndDate & "','" & pEffectiveDate & "'," & vbCrLf & _
                                " GETDATE(),'" & Session("UserID").ToString & "','" & pCurrency & "','" & pPrice & "','" & pPackingCls & "','" & pPriceCls & "','" & pDeliveryID & "')" & vbCrLf
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "insert"

                ElseIf pIsNewData = False Then
                    '#UPDATE DATA

                    'ls_SQL = "UPDATE MS_Price SET " & vbCrLf & _
                    '        "EndDate='" & pEndDate & "'," & vbCrLf & _
                    '        "EffectiveDate='" & pEffectiveDate & "'," & vbCrLf & _
                    '        "Price='" & CDec(pPrice) & "'," & vbCrLf & _
                    '        "UpdateDate = getdate()," & vbCrLf & _
                    '        "PriceCls='" & pPriceCls & "'," & vbCrLf & _
                    '        "UpdateUser ='" & Session("UserID").ToString & "'" & vbCrLf & _
                    '        "WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "'" & vbCrLf & _
                    '        " and StartDate ='" & pStartDate & "' and CurrCls='" & pCurrency & "' and PackingCls = '" & pPackingCls & "'" & vbCrLf
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_Sql2, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "update"

                    txtPartNo2.Text = ""
                    txtSupplier2.Text = ""
                    TxtPrice.Text = ""
                    cboPartNo2.Text = ""
                    cboSupplier2.Text = ""
                    dt4.Value = Now
                    dt5.Value = Now
                    dt6.Value = Now

                    Dim ls_Remarks As String = ""

                    If CDate(pEffectiveDate) <> CDate(pub_effdate) Then
                        ls_Remarks = ls_Remarks + "Effective Date " + pub_effdate & "->" & pEffectiveDate & " "
                    End If
                    If CDate(pEndDate) <> CDate(pub_enddate) Then
                        ls_Remarks = ls_Remarks + "End Date " + pub_enddate & "->" & pEndDate & " "
                    End If
                    If pPackingCls <> pub_packingCls Then
                        ls_Remarks = ls_Remarks + "Packing Cls " + pub_packingCls & "->" & pPackingCls & " "
                    End If
                    If pPrice <> pub_price Then
                        ls_Remarks = ls_Remarks + "Price " + pub_price & "->" & pPrice & " "
                    End If
                    If pPriceCls <> pub_pricecls Then
                        ls_Remarks = ls_Remarks + "PriceCls " + pub_pricecls & "->" & pPriceCls & " "
                    End If
                    If pDeliveryID <> pub_affcode Then
                        ls_Remarks = ls_Remarks + "AffiliateCode " + pub_affcode & "->" & pDeliveryID & " "
                    End If

                    Dim ls_Remarks2 As String = "Affiliate " & pAffiliateID & " StartDate " & pStartDate & " PackingCls " & pPackingCls & " Curr " & pCurrency & ", "

                    If ls_Remarks <> "" Then
                        'insert into history
                        ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                                 "VALUES ('" & shostname & "','" & menuID & "','U','" & pPartNo & "','Update [" & ls_Remarks2 & ls_Remarks & "]', " & vbCrLf & _
                                 "GETDATE(), '" & Session("UserID") & "')  "
                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                    End If

                End If

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        grid.JSProperties("cpType") = "info"


    End Sub

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Supplier Price Master " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Template\Result\" & tempFile & "")

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
                For irow = 0 To pData.Rows.Count - 1
                    For icol = 1 To pData.Columns.Count - 0
                        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                        If icol = 3 Or icol = 4 Or icol = 5 Then
                            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "dd-mmm-yy"
                        End If
                    Next
                Next

                ''ALIGNMENT
                ''.Cells(rowstart + 1, icol, irow, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
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
                Dim rgAll As ExcelRange = .Cells(2, 1, irow + 2, 10)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Template\Result\" & tempFile & "")

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