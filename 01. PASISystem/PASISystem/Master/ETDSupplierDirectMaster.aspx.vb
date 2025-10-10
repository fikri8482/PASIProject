Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing
Imports OfficeOpenXml
Imports System.IO

Public Class ETDSupplierDirectMaster
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
    Dim menuID As String = "A23"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ls_AllowDownload As String = clsGlobal.Auth_UserConfirm(Session("UserID"), menuID)
        Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
        Dim ls_AllowDelete As String = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "E.T.D SUPPLIER DIRECT TO AFFILIATE"
            up_FillCombo()
            DtPeriod.Focus()
            DtPeriod.Value = Now
            Format(DtPeriod.Value.Now, ("MMM yyyy"))
            dtSupplier.Value = Now
            dtAffiliate.Value = Now
            txtMode.ForeColor = Color.White
            If Session("M01Url") <> "" Then
                Call bindData()
                'ScriptManager.RegisterStartupScript(grid, grid.GetType(), "init", "grid.SetFocusedRowIndex(-1);", True)
                Session.Remove("M01Url")
            End If

            lblInfo.Text = ""

            grid.FocusedRowIndex = -1
        End If

        If ls_AllowDownload = False Then btnDownload.Enabled = False
        If ls_AllowUpdate = False Then btnUpload.Enabled = False
        If ls_AllowUpdate = False Then btnSubmit.Enabled = False
        If ls_AllowDelete = False Then btnDelete.Enabled = False

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "AffiliateID" Or e.Column.FieldName = "ETAAffiliate" Or e.Column.FieldName = "ETDSupplier" _
            Or e.Column.FieldName = "SupplierID" Or e.Column.FieldName = "SupplierName") _
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

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
            'grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliateID)

                    Dim affiliate As String = Split(e.Parameters, "|")(2)
                    Dim Supplier As String = Split(e.Parameters, "|")(3)
                    Dim ETAAffiliate As String = Split(e.Parameters, "|")(4)
                    Dim ETDSupplier As String = Split(e.Parameters, "|")(5)

                    Call SaveData(lb_IsUpdate, _
                                  affiliate.Trim, _
                                     Supplier.Trim, _
                                    ETAAffiliate, _
                                    ETDSupplier)
                    grid.FocusedRowIndex = -1
                    ScriptManager.RegisterStartupScript(grid, grid.GetType(), "init", "grid.SetFocusedRowIndex(-1);", True)
                    'bindData()
                    'Case "delete"
                    '    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    '    Dim pCurrCls As String = Split(e.Parameters, "|")(4)
                    '    Dim pStartDate As String = Split(e.Parameters, "|")(3)
                    '    Dim pPartNo As String = Split(e.Parameters, "|")(1)
                    '    'Dim ls_date As String = ""
                    '    'ls_date = Mid(pStartDate, 5, 11)
                    '    If AlreadyUsed(pAffiliateID, pPartNo, pCurrCls, pStartDate) = False Then
                    '        'pSupplierID,
                    '        Call DeleteData(pAffiliateID, pPartNo, pCurrCls, pStartDate)
                    '        Call bindData()
                    '    End If
                    '    txtMode.Text = "new"

                Case "delete"
                    Dim pSupplierID As String = Split(e.Parameters, "|")(1)
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim pStartDate As String = Split(e.Parameters, "|")(3)

                    'Dim ls_date As String = ""
                    'ls_date = Mid(pStartDate, 5, 11)
                    If AlreadyUsed(pSupplierID, pAffiliateID, pStartDate) = False Then
                        'pSupplierID,
                        Call DeleteData(pSupplierID, pAffiliateID, pStartDate)
                        Call bindData()
                    End If
                    txtMode.Text = "new"

                Case "load"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If
                    txtMode.Text = "new"
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                    grid.FocusedRowIndex = -1
                    'buat refresh grid, taruh di source aspx (grid.CollapseAll())
                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = clsMaster.GetTableETDDirect(DtPeriod.Value, cboAffiliate.Text, cboSupplier.Text)
                    FileName = "TemplateMSETDSupplierDirect.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:3", psERR)
                    End If
            End Select

EndProcedure:
            Session("AA220Msg") = ""
            grid.FocusedRowIndex = -1

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Response.Redirect("~/Upload/UploadETDDirect.aspx")
    End Sub

#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        'Dim pAffiliateID As String = ""
        Dim ld_Day As Integer = 0
        Dim JumlahHari As Integer = Date.DaysInMonth(Year(DtPeriod.Value), Month(DtPeriod.Value))
        'Dim JumlahHari As Integer = Date.DaysInMonth(DtPeriod.Value.year, DtPeriod.Value.month)

        If Format(DtPeriod.Value, "yyyy-MM-dd") <> "" Then
            pWhere = pWhere + " and ETAAffiliate like '%" & Format(DtPeriod.Value, "yyyy-MM-dd") & "%' "
        End If

        If cboAffiliate.Text.Trim <> "" Then
            ls_SQL = ls_SQL + "and AffiliateID = '" & cboAffiliate.Text.Trim & "' "
        End If

        If cboSupplier.Text.Trim <> "" Then
            ls_SQL = ls_SQL + "and SupplierID = '" & cboSupplier.Text.Trim & "' "
        End If

        Select Case Month(DtPeriod.Value)
            Case 1, 3, 5, 7, 8, 10, 12
                ld_Day = 31

            Case 4, 6, 9, 11
                ld_Day = 30

            Case 2
                If (Year(DtPeriod.Value) / 4) > 0 Then
                    ld_Day = 28
                Else
                    ld_Day = 29
                End If
        End Select

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            'ls_SQL = " DECLARE @Period AS DATETIME  " & vbCrLf & _
            '      " SET @Period = '" & Format(DtPeriod.Value, "yyyy-MMM-dd") & "'  " & vbCrLf & _
            '      "   " & vbCrLf & _
            '      " SELECT row_number() OVER ( ORDER BY CONVERT(numeric,DS.SeqNo)) NoUrut ,  " & vbCrLf & _
            '      " ETAAffiliate = COALESCE(ETAAffiliate, " & vbCrLf & _
            '      " CONVERT(VARCHAR, YEAR(@Period)) + '-' " & vbCrLf & _
            '      " + CONVERT(VARCHAR, MONTH(@Period)) + '-' " & vbCrLf & _
            '      " + CONVERT(VARCHAR, DS.SeqNo)) , " & vbCrLf & _
            '      " CONVERT(CHAR(15), MEP.ETDSupplier, 112) ETDSupplier , " & vbCrLf & _
            '      " MEP.SupplierID " & vbCrLf & _
            '      " FROM ( SELECT TOP 31 "

            'ls_SQL = ls_SQL + " * " & vbCrLf & _
            '                  " FROM DateSeqNo " & vbCrLf & _
            '                  " ORDER BY SeqNo " & vbCrLf & _
            '                  " ) DS " & vbCrLf & _
            '                  " LEFT JOIN MS_ETD_SUPPLIER_Direct MEP ON DAY(MEP.ETAAffiliate) = DS.SeqNo " & vbCrLf & _
            '                  " AND CONVERT(CHAR(6), ETAAffiliate, 112) = LEFT(CONVERT(CHAR(8), @Period, 112),6) " & vbCrLf & _
            '                  " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = MEP.AffiliateID " & vbCrLf & _
            '                  " LEFT JOIN MS_Supplier MS ON MS.SupplierID = MEP.SupplierID " & vbCrLf & _
            '                  " WHERE DS.SeqNo <= '" & JumlahHari & "' " & vbCrLf & _
            '                  " ORDER BY CONVERT(numeric,DS.SeqNo) " & vbCrLf & _
            '                  "  "
            ls_SQL = " DECLARE @Period AS DATETIME   " & vbCrLf & _
                              "  SET @Period = '" & Format(DtPeriod.Value, "yyyy-MMM-dd") & "'   " & vbCrLf & _
                              "     " & vbCrLf & _
                              "  SELECT row_number() OVER ( ORDER BY CONVERT(numeric,DS.SeqNo)) NoUrut ,   " & vbCrLf & _
                              " 		 ETAAffiliate = COALESCE(ETAAffiliate,  " & vbCrLf & _
                              " 								 CONVERT(VARCHAR, YEAR(@Period)) + '-'  " & vbCrLf & _
                              " 								 + CONVERT(VARCHAR, MONTH(@Period)) + '-'  " & vbCrLf & _
                              " 								 + CONVERT(VARCHAR, DS.SeqNo)) ,  " & vbCrLf & _
                              " 		 ETDSupplier  = ISNULL(CONVERT(CHAR(12), MEP.ETDSupplier, 113),''),  " & vbCrLf & _
                              " 		 AffiliateID = '" & cboAffiliate.Text.Trim & "',  " & vbCrLf & _
                              " 		 AffiliateName = '" & txtAffiliate.Text.Trim & "', " & vbCrLf & _
                              " 		 SupplierID = '" & cboSupplier.Text.Trim & "',  " & vbCrLf & _
                              " 		 SupplierName = '" & txtSupplier.Text.Trim & "' "

            ls_SQL = ls_SQL + "  FROM ( SELECT TOP 31  *  " & vbCrLf & _
                              " 		 FROM DateSeqNo  " & vbCrLf & _
                              " 		 ORDER BY SeqNo  " & vbCrLf & _
                              " 		 ) DS  " & vbCrLf & _
                              "          full JOIN ( " & vbCrLf & _
                              " 			select *  from MS_Affiliate where AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                              " 		) MA on  1 = 1 " & vbCrLf & _
                              "          full JOIN ( " & vbCrLf & _
                              " 			 select * from MS_Supplier where SupplierID = '" & cboSupplier.Text.Trim & "' " & vbCrLf & _
                              " 			) MS on 1 = 1 " & vbCrLf & _
                              " 		 LEFT JOIN  " & vbCrLf & _
                              " 		 (SELECT *  " & vbCrLf & _
                              " 		    FROM MS_ETD_Supplier_Direct  " & vbCrLf & _
                              " 		  WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & cboSupplier.Text.Trim & "' " & vbCrLf & _
                              " 		 ) MEP " & vbCrLf & _
                              " 		 ON DAY(MEP.ETAAffiliate) = DS.SeqNo  " & vbCrLf & _
                              " 		 AND CONVERT(CHAR(6), ETAAffiliate, 112) = LEFT(CONVERT(CHAR(8), @Period, 112),6)  "

            ls_SQL = ls_SQL + " 		 and MA.AffiliateID = MEP.AffiliateID  " & vbCrLf & _
                              " 		 and MS.SupplierID = MEP.SupplierID  " & vbCrLf & _
                              " WHERE  DS.SeqNo <= '" & JumlahHari & "' " & vbCrLf & _
                              " ORDER BY CONVERT(numeric,DS.SeqNo) " & vbCrLf & _
                              "  " & vbCrLf & _
                              "  "

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

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ''AffiliateID, '' AffiliateName, 
            ls_SQL = " select top 0  '' NoUrut, '' AffiliateID, '' ETAAffiliate, '' ETDSupplier"

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
        ls_SQL = "select RTRIM(AffiliateID) AffiliateCode, RTRIM(AffiliateName) AffiliateName from MS_Affiliate order by AffiliateCode " & vbCrLf
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
                .SelectedIndex = -1
                'txtAffiliate.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(AffiliateID) AffiliateCode, RTRIM(AffiliateName) AffiliateName from MS_Affiliate order by AffiliateCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliate2
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateCode")
                .Columns(0).Width = 75
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 400

                .TextField = "AffiliateCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
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
                .SelectedIndex = -1
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

    End Sub

    Private Function AlreadyUsed(ByVal pAffiliateID As String, ByVal pSupplierID As String, ByVal pStartDate As Date) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT * FROM MS_ETD_SUPPLIER_Direct WHERE AffiliateID ='" & pAffiliateID.Trim & "' and SupplierID ='" & pSupplierID.Trim & "' and ETAAffiliate='" & pStartDate & "' "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    'lblInfo.Text = "Affiliate ID already used in other screen"
                    Call clsMsg.DisplayMessage(lblInfo, "5004", clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    grid.JSProperties("cpType") = "error"
                    Return True
                Else
                    Return False
                End If
                Return True
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Function

    Private Sub DeleteData(ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pStartDate As Date)
        'ByVal pSupplierID As String,
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                pStartDate = Format(DateAdd(DateInterval.Month, 1, CDate(pStartDate)), "yyyy-MM-dd")
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("DelETDSupplierDirectMaster")
                    ls_SQL = " DELETE MS_ETD_Supplier_Direct " & vbCrLf & _
                                " WHERE SupplierID ='" & pSupplierID.Trim & "'" & vbCrLf & _
                                " and AffiliateID ='" & pAffiliateID.Trim & "'" & vbCrLf & _
                                " and ETAAffiliate ='" & pStartDate & "'" & vbCrLf

                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

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

    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
        Try
            'Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""

            'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            '    sqlConn.Open()

            '    ls_SQL = "SELECT AffiliateID" & vbCrLf & _
            '                " FROM MS_Affiliate " & _
            '                " WHERE AffiliateID= '" & Trim(pAffiliate) & "'"

            '    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            '    Dim ds As New DataSet
            '    sqlDA.Fill(ds)

            '    If ds.Tables(0).Rows.Count > 0 And grid.FocusedRowIndex = -1 Then
            '        ls_MsgID = "6018"
            '        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            '        AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
            '        flag = False
            '        Return False
            '    ElseIf ds.Tables(0).Rows.Count > 0 Then
            '        lblInfo.Text = "Affiliate ID with ID " & txtPartID.Text & " already exists in the database."
            '        Return False
            '    End If
            '    Return True
            '    sqlConn.Close()
            'End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

    End Function

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffiliateID As String = "", _
                         Optional ByVal pSupplierID As String = "", _
                         Optional ByVal pStartDate As String = "", _
                         Optional ByVal pEndDate As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString
        'and AffiliateID ='" & pAffiliateID & "'
        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                pStartDate = Format(DateAdd(DateInterval.Month, 1, CDate(pStartDate)), "yyyy-MM-dd")
                pEndDate = Format(DateAdd(DateInterval.Month, 1, CDate(pEndDate)), "yyyy-MM-dd")
                ls_SQL = " SELECT * FROM MS_ETD_SUPPLIER_Direct WHERE AffiliateID ='" & pAffiliateID.Trim & "' and SupplierID ='" & pSupplierID.Trim & "' and ETAAffiliate ='" & pStartDate & "' "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    pIsNewData = False
                Else
                    pIsNewData = True
                End If
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        If txtMode.Text = "update" Then
            flag = False
        Else
            flag = True
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_ETD_SUPPLIER_Direct " & _
                                "(AffiliateID, SupplierID, ETAAffiliate, ETDSupplier)" & _
                                " VALUES ('" & pAffiliateID.Trim & "','" & pSupplierID.Trim & "','" & pStartDate & "','" & pEndDate & "' )" & vbCrLf
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "insert"

                ElseIf pIsNewData = False Then
                    'ls_MsgID = "6018"
                    'Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    'grid.JSProperties("cpMessage") = lblInfo.Text
                    'grid.JSProperties("cpType") = "error"
                    ' Exit Sub

                    'ElseIf pIsNewData = False Then
                    '#UPDATE DATA
                    'and SupplierID = '" & pSupplierID & "'
                    '"CurrCls='" & CboCurrency.Value & "'," & _
                    ls_SQL = "UPDATE MS_ETD_SUPPLIER_Direct SET " & _
                            "ETDSupplier='" & pEndDate & "' " & _
                            "WHERE AffiliateID ='" & pAffiliateID.Trim & "' and SupplierID ='" & pSupplierID.Trim & "'" & vbCrLf & _
                            " and ETAAffiliate ='" & pStartDate & "' " & vbCrLf
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "update"

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
            Dim tempFile As String = "ETD Supplier Direct Master " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                    For icol = 1 To pData.Columns.Count
                        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
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
                Dim rgAll As ExcelRange = .Cells(2, 1, irow + 2, 4)
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

    'Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
    '    If e.GetValue("AffiliateID") = "" Then
    '        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
    '    End If
    'End Sub
End Class