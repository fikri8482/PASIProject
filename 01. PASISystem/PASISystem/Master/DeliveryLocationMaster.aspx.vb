Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing
Imports OfficeOpenXml
Imports System.IO

Public Class DeliveryLocationMaster
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
    Dim menuID As String = "A18"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ls_AllowDownload As String = clsGlobal.Auth_UserConfirm(Session("UserID"), menuID)
        Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
        Dim ls_AllowDelete As String = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "DELIVERY LOCATION MASTER"
            up_FillCombo()
            Call bindData()
            DeleteHistory()
            lblInfo.Text = ""
        End If

        If ls_AllowDownload = False Then btnDownload.Enabled = False
        If ls_AllowUpdate = False Then btnUpload.Enabled = False
        If ls_AllowUpdate = False Then btnSubmit.Enabled = False
        If ls_AllowDelete = False Then btnDelete.Enabled = False

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "AffiliateID" Or e.Column.FieldName = "DeliveryLocationCode" _
            Or e.Column.FieldName = "DeliveryLocationName" Or e.Column.FieldName = "Address" Or e.Column.FieldName = "City" Or e.Column.FieldName = "PostalCode" _
            Or e.Column.FieldName = "Phone1" Or e.Column.FieldName = "Phone2" Or e.Column.FieldName = "Fax" Or e.Column.FieldName = "NPWP" _
            Or e.Column.FieldName = "PODeliveryBy" Or e.Column.FieldName = "DefaultCls") _
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

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Response.Redirect("~/Upload/UploadDeliveryLocation.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
            'grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    'Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliateID)
                    Dim lb_IsUpdate As Boolean = True
                    'If cboAffiliateCode.ReadOnly = False Then
                    '    txtMode.Text = "update"
                    'Else
                    '    txtMode.Text = "new"
                    'End If

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
                                     Split(e.Parameters, "|")(11), _
                                     Split(e.Parameters, "|")(12), _
                                     Split(e.Parameters, "|")(13))
                    'Call bindData()

                Case "delete"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(1)
                    Dim pDeliveryLoc As String = Split(e.Parameters, "|")(2)

                    'If AlreadyUsed(pSupplierGroupCode) = True Then
                    'pSupplierID,
                    Call DeleteData(pAffiliateID, pDeliveryLoc)
                    Call bindData()
                    'End If
                    '                    txtMode.Text = "new"
                Case "load"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                    'cboAffiliateCode.Text = ""
                    'txtDeliveryLoc.Text = ""
                    'txtDeliveryLocName.Text = ""
                    'txtAddress.Text = ""
                    'txtCity.Text = ""
                    'txtPostalCode.Text = ""
                    'txtPhone1.Text = ""
                    'txtPhone2.Text = ""
                    'txtFax.Text = ""
                    'txtNPWP.Text = ""
                    'cboPODelby.Text = ""
                    'cboDefault.Text = ""
                    'lblInfo.Text = ""

                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim pSuppType As String = ""

                    Dim dtProd As DataTable = clsMaster.GetTableDeliveryLocation()
                    FileName = "TemplateMSDeliveryLocation.xlsx"
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

    Private Sub DeleteHistory()
        Dim ls_sql As String

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                    Dim sqlComm As New SqlCommand()

                    ls_sql = " delete MS_DeliveryPlace_History " & vbCrLf & _
                              " where exists " & vbCrLf & _
                              " (select * from MS_DeliveryPlace a where MS_DeliveryPlace_History.AffiliateID = a.AffiliateID  " & vbCrLf & _
                              "   and MS_DeliveryPlace_History.DeliveryLocationCode = a.DeliveryLocationCode) "

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

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  "  	row_number() over (order by AffiliateID, DeliveryLocationCode) NoUrut,  " & vbCrLf & _
                  "  	RTRIM(AffiliateID)AffiliateID,  " & vbCrLf & _
                  "  	RTRIM(DeliveryLocationCode)DeliveryLocationCode,  " & vbCrLf & _
                  "  	RTRIM(DeliveryLocationName)DeliveryLocationName,  " & vbCrLf & _
                  "  	RTRIM(Address)Address,  " & vbCrLf & _
                  "  	RTRIM(City)City,  " & vbCrLf & _
                  "  	RTRIM(PostalCode)PostalCode,  " & vbCrLf & _
                  "  	RTRIM(Phone1)Phone1,  " & vbCrLf & _
                  "  	RTRIM(Phone2)Phone2,  " & vbCrLf & _
                  "  	RTRIM(Fax)Fax,  "

            ls_SQL = ls_SQL + "  	RTRIM(NPWP)NPWP,  " & vbCrLf & _
                              "  	(case PODeliveryBy when '1' then 'PASI' " & vbCrLf & _
                              "  	      when '0' then 'SUPPLIER' " & vbCrLf & _
                              "  	      else '' " & vbCrLf & _
                              "  	      end) PODeliveryBy,  " & vbCrLf & _
                              "  	(case DefaultCls when '1' then 'YES' " & vbCrLf & _
                              "  	      when '0' then 'NO' " & vbCrLf & _
                              "  	      else '' " & vbCrLf & _
                              "  	      end) DefaultCls, 0 DeleteCls, EntryDate, EntryUser, UpdateDate, UpdateUser  " & vbCrLf & _
                              "  from MS_DeliveryPlace  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
            End With
            sqlConn.Close()

            grid.FocusedRowIndex = -1

            lblInfo.Text = ""
        End Using
    End Sub

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        'Person In Charge
        ls_SQL = "SELECT RTRIM(AffiliateID) AffiliateCode, AffiliateName from MS_Affiliate order by AffiliateCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateCode")
                .Columns(0).Width = 75


                .TextField = "AffiliateCode"
                .DataBind()
                '.SelectedIndex = -1

            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub tabIndex()
        cboAffiliateCode.TabIndex = 1
        txtDeliveryLoc.TabIndex = 2
        txtDeliveryLocName.TabIndex = 3
        txtAddress.TabIndex = 4
        txtCity.TabIndex = 5
        txtPostalCode.TabIndex = 6
        txtPhone1.TabIndex = 7
        txtPhone2.TabIndex = 8
        txtFax.TabIndex = 9
        txtNPWP.TabIndex = 10
        cboPODelby.TabIndex = 11
        cboDefault.TabIndex = 12
        btnSubmit.TabIndex = 13
        btnDelete.TabIndex = 14
        btnClear.TabIndex = 15
        btnSubMenu.TabIndex = 16
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ''AffiliateID, '' AffiliateName, 
            ls_SQL = " select top 0  '' NoUrut, ''AffiliateID, '' DeliveryLocationCode, '' DeliveryLocationName, '' Address, '' City, '' PostalCode, '' Phone1, '' Phone2, '' Fax, '' NPWP, '' PODeliveryBy, '' DefaultCls "

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

    Private Function AlreadyUsed(ByVal pAffiliateID As String, ByVal pDeliveryLoc As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT AffiliateID FROM MS_DeliveryPlace WHERE AffiliateID = '" & Trim(pAffiliateID) & "' and DeliveryLocationCode= '" & Trim(pDeliveryLoc) & "'"

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

    Private Sub DeleteData(ByVal pAffiliateID As String, ByVal pDeliveryLoc As String)
        'ByVal pSupplierID As String,
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                'and SupplierID = '" & pSupplierID & "'
                'and AffiliateID ='" & pAffiliateID & "'
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("DeliveryPlace")
                    ls_SQL = " DELETE MS_DeliveryPlace " & vbCrLf & _
                                " WHERE AffiliateID = '" & Trim(pAffiliateID) & "' and DeliveryLocationCode= '" & Trim(pDeliveryLoc) & "'" & vbCrLf

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

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffiliateID As String = "", _
                         Optional ByVal pDeliveryLoc As String = "", _
                         Optional ByVal pDeliveryLocName As String = "", _
                         Optional ByVal pAddress As String = "", _
                         Optional ByVal pCity As String = "", _
                         Optional ByVal pPostalCode As String = "", _
                         Optional ByVal pPhone1 As String = "", _
                         Optional ByVal pPhone2 As String = "", _
                         Optional ByVal pFax As String = "", _
                         Optional ByVal pNPWP As String = "", _
                         Optional ByVal pPODeliveryby As String = "", _
                        Optional ByVal pDefaultCls As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString
        'and AffiliateID ='" & pAffiliateID & "'
        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT AffiliateID FROM MS_DeliveryPlace WHERE AffiliateID ='" & pAffiliateID & "' and DeliveryLocationCode ='" & pDeliveryLoc & "'  "
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

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("DeliveryPlace")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = ""
                    'If cboDefault.Value = "1" Then
                    '    ls_SQL = " Update MS_DeliveryPlace set DefaultCls = 0 where affiliateID = '" & cboAffiliateCode.Text & "'" & vbCrLf
                    'End If

                    ls_SQL = ls_SQL + " INSERT INTO MS_DeliveryPlace " & _
                                "(AffiliateID, DeliveryLocationCode, DeliveryLocationName, Address, City, PostalCode, Phone1, Phone2, Fax, NPWP, PODeliveryby, DefaultCls, EntryDate, EntryUser, UpdateDate, UpdateUser)" & _
                                " VALUES ('" & pAffiliateID & "','" & pDeliveryLoc & "','" & pDeliveryLocName & "','" & pAddress & "', " & vbCrLf & _
                                " '" & pCity & "','" & pPostalCode & "','" & pPhone1 & "','" & pPhone2 & "', " & vbCrLf & _
                                " '" & pFax & "','" & pNPWP & "','" & pPODeliveryby & "','" & pDefaultCls & "', " & vbCrLf & _
                                " GETDATE(),'" & admin & "', GETDATE(),'" & admin & "')" & vbCrLf

                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "insert"

                    'ElseIf pIsNewData = False And flag = True Then
                    '    ls_MsgID = "6018"
                    '    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    '    grid.JSProperties("cpMessage") = lblInfo.Text
                    '    grid.JSProperties("cpType") = "error"
                    '    Exit Sub

                ElseIf pIsNewData = False Then
                    '#UPDATE DATA
                    'and SupplierID = '" & pSupplierID & "'
                    'ls_SQL = ""
                    'If cboDefault.Value = "1" Then
                    '    ls_SQL = " Update MS_DeliveryPlace set DefaultCls = 0 where affiliateID = '" & cboAffiliateCode.Text & "'" & vbCrLf
                    'End If

                    ls_SQL = ls_SQL + " UPDATE MS_DeliveryPlace SET " & _
                            "DeliveryLocationName = '" & pDeliveryLocName & "', " & vbCrLf & _
                            "Address = '" & pAddress & "', " & vbCrLf & _
                            "City = '" & pCity & "', " & vbCrLf & _
                            "PostalCode = '" & pPostalCode & "', " & vbCrLf & _
                            "Phone1 = '" & pPhone1 & "', " & vbCrLf & _
                            "Phone2 = '" & pPhone2 & "', " & vbCrLf & _
                            "Fax = '" & pFax & "', " & vbCrLf & _
                            "NPWP = '" & pNPWP & "', " & vbCrLf & _
                            "PODeliveryBy = '" & pPODeliveryby & "', " & vbCrLf & _
                            "DefaultCls = '" & pDefaultCls & "' " & vbCrLf & _
                    " WHERE AffiliateID ='" & pAffiliateID & "' and DeliveryLocationCode ='" & pDeliveryLoc & "' " & vbCrLf
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
            Dim tempFile As String = "Delivery Location Master " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                Dim rgAll As ExcelRange = .Cells(2, 1, irow + 2, 12)
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