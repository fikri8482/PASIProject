'Update By Robby
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO
Imports System.Data.OleDb

Public Class UploadAffiliate
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "A02"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim log As String = ""
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            lblInfo.Text = ""
        Else
            lblInfo.Text = ""
            Ext = Server.MapPath("")
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/Master/AffiliateMaster.aspx")
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Uploader.NullText = "Click here to browse files..."

        lblInfo.Text = ""

        Uploader.Enabled = True
        btnSave.Enabled = True
        btnUpload.Enabled = True

        up_GridLoadWhenEventChange()
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        up_Import()
    End Sub

    Private Sub ASPxCallback1_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ASPxCallback1.Callback
        Try
            Dim fi As New FileInfo(Server.MapPath("~\Template\TemplatePO.xlsx"))
            If Not fi.Exists Then
                lblInfo.Text = "[9999] Excel Template Not Found !"
                ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("Template PO.xlsx")

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        End Try

    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If e.GetValue("ErrorCls") = "" Then
        Else
            e.Cell.BackColor = Color.Red
        End If
    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "6021", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    Else
                        Call up_Save()
                    End If
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select   " & vbCrLf &
                  "  	row_number() over (order by a.AffiliateID asc) as NoUrut,  " & vbCrLf &
                  "  	a.[AffiliateID], a.[AffiliateCode], a.[ConsigneeCode], a.[ConsigneeName], a.[BuyerName], " & vbCrLf &
                  " 	a.[BuyerCode], a.[AffiliateName], a.[Address], a.[ConsigneeAddress],  " & vbCrLf &
                  " 	a.[BuyerAddress], a.[City], a.[PostalCode], a.[Phone1], a.[Phone2], " & vbCrLf &
                  " 	a.[Fax], a.[NPWP], a.[KantorPabean], a.[IzinTPB], a.[BCPerson],  " & vbCrLf &
                  " 	a.[PODeliveryBy], a.[FolderOES], a.[OverseasCls], a.[DestinationPort], a.[ErrorCls], " & vbCrLf &
                  " 	xConsigneeCode = b.ConsigneeCode, xConsigneeName = b.ConsigneeName,  " & vbCrLf &
                  " 	xConsigneeAddress = b.ConsigneeAddress, xBuyerCode = b.BuyerCode,  " & vbCrLf &
                  " 	xBuyerName = b.BuyerName, xBuyerAddress = b.BuyerAddress,  " & vbCrLf &
                  " 	xAffiliateName = b.AffiliateName, xAddress = b.Address, xCity = b.City,  "

            ls_SQL = ls_SQL + " 	xPostalCode = b.PostalCode, xPhone1 = b.Phone1, xPhone2 = b.Phone2,  " & vbCrLf & _
                              " 	xFax = b.Fax, xNPWP = b.NPWP, xKantorPabean = b.KantorPabean,  " & vbCrLf & _
                              " 	xIzinTPB = b.IzinTPB, xBCPerson = b.BCPerson, xDestinationPort = b.DestinationPort,  " & vbCrLf & _
                              " 	xFolderOES = b.FolderOES, xPODeliveryBy = b.PODeliveryBy, xOverseasCls = b.OverseasCls   " & vbCrLf & _
                              " from [UploadAffiliate] a  " & vbCrLf & _
                              " left join MS_Affiliate b on a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " order by AffiliateID "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' [AffiliateID],'' [AffiliateCode],'' [ConsigneeCode],'' [BuyerCode],'' [AffiliateID], " & vbCrLf &
                  " '' [AffiliateName],'' [Address],'' [BuyerAddress],'' [ConsigneeAddress],'' [DestinationPort], " & vbCrLf &
                  " ''[OverseasCls],'' [City],'' [PostalCode],'' [Phone1],'' [Phone2],'' [Fax],'' [NPWP], " & vbCrLf &
                  " '' [KantorPabean],'' [IzinTPB],'' [BCPerson],'' [PODeliveryBy],'' [FolderOES],'' [ErrorCls], " & vbCrLf &
                  " '' xConsigneeCode, '' xConsigneeName, '' xConsigneeAddress, '' xBuyerCode, '' xBuyerName,  " & vbCrLf &
                  " '' xBuyerAddress, '' xAffiliateName, '' xAddress, '' xCity, '' xPostalCode, '' xPhone1,  " & vbCrLf &
                  " '' xPhone2, '' xFax, '' xNPWP, '' xKantorPabean, '' xIzinTPB, '' xBCPerson, '' xDestinationPort,  " & vbCrLf &
                  " '' xFolderOES, '' xPODeliveryBy, '' xOverseasCls " & vbCrLf &
                  "  " & vbCrLf &
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

    Private Sub up_Import()
        Dim dt As New System.Data.DataTable
        Dim dtHeader As New System.Data.DataTable
        Dim dtDetail As New System.Data.DataTable
        'Dim tempDate As Date
        Dim ls_MOQ As Double = 0
        Dim ls_sql As String = ""
        Dim ls_SupplierID As String = """"

        Try
            lblInfo.ForeColor = Color.Red
            If (Not Uploader.PostedFile Is Nothing) And (Uploader.PostedFile.ContentLength > 0) Then
                FileName = Path.GetFileName(Uploader.PostedFile.FileName)
                FileExt = Path.GetExtension(Uploader.PostedFile.FileName)
                    FilePath = Ext & "\Import\" & FileName
                    Dim fi As New FileInfo(Server.MapPath("~\Import\" & FileName))
                    If fi.Exists Then
                        fi.Delete()
                        fi = New FileInfo(Server.MapPath("~\Import\" & FileName))
                    End If
                    Uploader.SaveAs(FilePath)

                    Dim connStr As String = ""
                    Select Case FileExt
                        Case ".xls"
                            'Excel 97-03
                            connStr = ConfigurationManager.ConnectionStrings("Excel03ConString").ConnectionString
                        Case ".xlsx"
                            'Excel 07
                            connStr = ConfigurationManager.ConnectionStrings("Excel07ConString").ConnectionString
                    End Select

                    connStr = String.Format(connStr, FilePath, "No")

                    Dim MyConnection As New OleDbConnection(connStr)
                    Dim MyCommand As New OleDbCommand
                    Dim MyAdapter As New OleDbDataAdapter
                    MyCommand.Connection = MyConnection
                    MyConnection.Open()

                    Dim dtSheets As DataTable = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                    Dim listSheet As New List(Of String)
                    Dim drSheet As DataRow

                    For Each drSheet In dtSheets.Rows
                        If InStr("_xlnm#_FilterDatabase", drSheet("TABLE_NAME").ToString(), CompareMethod.Text) = 0 Then
                            If InStr("_xlnm#Print_Titles", drSheet("TABLE_NAME").ToString(), CompareMethod.Text) = 0 Then
                                listSheet.Add(drSheet("TABLE_NAME").ToString())
                            End If
                        End If
                    Next

                    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                        sqlConn.Open()

                        ''==========Table EXCEL Master==========
                        Dim pTableCode As String = listSheet(0)

                        Try

                            'Get Detail Data
                            Dim dtUploadDetailList As New List(Of clsMaster)

                            MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A3:W65536]")
                            MyAdapter.SelectCommand = MyCommand
                            MyAdapter.Fill(dtDetail)

                            If dtDetail.Rows.Count > 0 Then
                                For i = 0 To dtDetail.Rows.Count - 1
                                    If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                    Dim dtUploadDetail As New clsMaster
                                    dtUploadDetail.AffiliateID = dtDetail.Rows(i).Item(0)
                                    dtUploadDetail.AffiliateCode = dtDetail.Rows(i).Item(1)
                                    dtUploadDetail.ConsigneeCode = IIf(IsDBNull(dtDetail.Rows(i).Item(2)), "", dtDetail.Rows(i).Item(2))
                                    dtUploadDetail.BuyerCode = IIf(IsDBNull(dtDetail.Rows(i).Item(3)), "", dtDetail.Rows(i).Item(3))
                                    dtUploadDetail.AffiliateName = IIf(IsDBNull(dtDetail.Rows(i).Item(4)), "", dtDetail.Rows(i).Item(4))
                                    dtUploadDetail.Address = IIf(IsDBNull(dtDetail.Rows(i).Item(5)), "", dtDetail.Rows(i).Item(5))
                                    dtUploadDetail.ConsigneeName = IIf(IsDBNull(dtDetail.Rows(i).Item(6)), "", dtDetail.Rows(i).Item(6))
                                    dtUploadDetail.ConsigneeAddress = IIf(IsDBNull(dtDetail.Rows(i).Item(7)), "", dtDetail.Rows(i).Item(7))
                                    dtUploadDetail.BuyerName = IIf(IsDBNull(dtDetail.Rows(i).Item(8)), "", dtDetail.Rows(i).Item(8))
                                    dtUploadDetail.BuyerAddress = IIf(IsDBNull(dtDetail.Rows(i).Item(9)), "", dtDetail.Rows(i).Item(9))
                                    dtUploadDetail.DestinationPort = IIf(IsDBNull(dtDetail.Rows(i).Item(10)), "", dtDetail.Rows(i).Item(10))

                                    dtUploadDetail.City = IIf(IsDBNull(dtDetail.Rows(i).Item(11)), "", dtDetail.Rows(i).Item(11))
                                    dtUploadDetail.PostalCode = IIf(IsDBNull(dtDetail.Rows(i).Item(12)), "", dtDetail.Rows(i).Item(12))
                                    dtUploadDetail.Phone1 = IIf(IsDBNull(dtDetail.Rows(i).Item(13)), "", dtDetail.Rows(i).Item(13))
                                    dtUploadDetail.Phone2 = IIf(IsDBNull(dtDetail.Rows(i).Item(14)), "", dtDetail.Rows(i).Item(14))
                                    dtUploadDetail.Fax = IIf(IsDBNull(dtDetail.Rows(i).Item(15)), "", dtDetail.Rows(i).Item(15))
                                    dtUploadDetail.NPWP = IIf(IsDBNull(dtDetail.Rows(i).Item(16)), "", dtDetail.Rows(i).Item(16))
                                    dtUploadDetail.KantorPabean = IIf(IsDBNull(dtDetail.Rows(i).Item(17)), "", dtDetail.Rows(i).Item(17))
                                    dtUploadDetail.IzinTPB = IIf(IsDBNull(dtDetail.Rows(i).Item(18)), "", dtDetail.Rows(i).Item(18))
                                    dtUploadDetail.BCPerson = IIf(IsDBNull(dtDetail.Rows(i).Item(19)), "", dtDetail.Rows(i).Item(19))
                                    dtUploadDetail.PODeliveryBy = IIf(IsDBNull(dtDetail.Rows(i).Item(20)), "0", dtDetail.Rows(i).Item(20))
                                    dtUploadDetail.OverseasCls = IIf(IsDBNull(dtDetail.Rows(i).Item(21)), "0", dtDetail.Rows(i).Item(21))
                                    dtUploadDetail.FolderOES = IIf(IsDBNull(dtDetail.Rows(i).Item(22)), "", dtDetail.Rows(i).Item(22))
                                    dtUploadDetailList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                                ''01.01 Delete TempoaryData
                                ls_sql = "delete UploadAffiliate"
                                Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm9.ExecuteNonQuery()
                                sqlComm9.Dispose()


                                ''02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                                For i = 0 To dtUploadDetailList.Count - 1
                                    Dim ls_error As String = ""
                                    Dim PO As clsMaster = dtUploadDetailList(i)

                                If PO.AffiliateID.Trim.Length > 20 Then
                                    If ls_error = "" Then
                                        ls_error = "Affiliate ID max length only 20"
                                    Else
                                        ls_error = ls_error + ",Affiliate ID max length only 20"
                                    End If
                                End If

                                If PO.AffiliateCode.Trim.Length > 4 Then
                                    If ls_error = "" Then
                                        ls_error = "Affiliate Code max length only 4"
                                    Else
                                        ls_error = ls_error + ",Affiliate Code max length only 4"
                                    End If
                                End If

                                If PO.AffiliateName.Trim.Length > 100 Then
                                        If ls_error = "" Then
                                            ls_error = "Affiliate Name max length only 100"
                                        Else
                                            ls_error = ls_error + ",Affiliate Name max length only 100"
                                        End If
                                    End If

                                    If PO.Address.Trim.Length > 250 Then
                                        If ls_error = "" Then
                                            ls_error = "Affiliate Address max length only 250"
                                        Else
                                            ls_error = ls_error + ",Affiliate Address max length only 250"
                                        End If
                                    End If

                                    If PO.ConsigneeCode.Trim.Length > 10 Then
                                        If ls_error = "" Then
                                            ls_error = "Consignee Code max length only 10"
                                        Else
                                            ls_error = ls_error + ",Consignee Code max length only 10"
                                        End If
                                    End If

                                    If PO.ConsigneeName.Trim.Length > 100 Then
                                        If ls_error = "" Then
                                            ls_error = "Consignee Name max length only 100"
                                        Else
                                            ls_error = ls_error + ",Consignee Name max length only 100"
                                        End If
                                    End If

                                    If PO.ConsigneeAddress.Trim.Length > 250 Then
                                        If ls_error = "" Then
                                            ls_error = "Consignee Address max length only 250"
                                        Else
                                            ls_error = ls_error + ",Consignee Address max length only 250"
                                        End If
                                    End If

                                    If PO.BuyerCode.Trim.Length > 10 Then
                                        If ls_error = "" Then
                                            ls_error = "Buyer Code max length only 10"
                                        Else
                                            ls_error = ls_error + ",Buyer Code max length only 10"
                                        End If
                                    End If

                                    If PO.BuyerName.Trim.Length > 100 Then
                                        If ls_error = "" Then
                                            ls_error = "Buyer Name max length only 100"
                                        Else
                                            ls_error = ls_error + ",Buyer Name max length only 100"
                                        End If
                                    End If

                                    If PO.BuyerAddress.Trim.Length > 250 Then
                                        If ls_error = "" Then
                                            ls_error = "Buyer Address max length only 250"
                                        Else
                                            ls_error = ls_error + ",Buyer Address max length only 250"
                                        End If
                                    End If

                                    If PO.City.Trim.Length > 20 Then
                                        If ls_error = "" Then
                                            ls_error = "City max length only 20"
                                        Else
                                            ls_error = ls_error + ",City max length only 20"
                                        End If
                                    End If

                                    If PO.PostalCode.Trim.Length > 15 Then
                                        If ls_error = "" Then
                                            ls_error = "Postal Code max length only 15"
                                        Else
                                            ls_error = ls_error + ",Postal Code max length only 15"
                                        End If
                                    End If

                                    If PO.Phone1.Trim.Length > 20 Then
                                        If ls_error = "" Then
                                            ls_error = "Phone1 max length only 20"
                                        Else
                                            ls_error = ls_error + ",Phone1 max length only 20"
                                        End If
                                    End If

                                    If PO.Phone2.Trim.Length > 20 Then
                                        If ls_error = "" Then
                                            ls_error = "Phone2 max length only 20"
                                        Else
                                            ls_error = ls_error + ",Phone2 max length only 20"
                                        End If
                                    End If

                                    If PO.Fax.Trim.Length > 20 Then
                                        If ls_error = "" Then
                                            ls_error = "Fax max length only 20"
                                        Else
                                            ls_error = ls_error + ",Fax max length only 20"
                                        End If
                                    End If

                                    If PO.NPWP.Trim.Length > 25 Then
                                        If ls_error = "" Then
                                            ls_error = "NPWP max length only 25"
                                        Else
                                            ls_error = ls_error + ",NPWP max length only 25"
                                        End If
                                    End If

                                    If PO.KantorPabean.Trim.Length > 100 Then
                                        If ls_error = "" Then
                                            ls_error = "KantorPabean max length only 100"
                                        Else
                                            ls_error = ls_error + ",KantorPabean max length only 100"
                                        End If
                                    End If

                                    If PO.IzinTPB.Trim.Length > 100 Then
                                        If ls_error = "" Then
                                            ls_error = "IzinTPB max length only 100"
                                        Else
                                            ls_error = ls_error + ",IzinTPB max length only 100"
                                        End If
                                    End If

                                    If PO.BCPerson.Trim.Length > 100 Then
                                        If ls_error = "" Then
                                            ls_error = "BCPerson max length only 100"
                                        Else
                                            ls_error = ls_error + ",BCPerson max length only 100"
                                        End If
                                    End If

                                    If PO.FolderOES.Trim.Length > 300 Then
                                        If ls_error = "" Then
                                            ls_error = "FolderOES max length only 300"
                                        Else
                                            ls_error = ls_error + ",FolderOES max length only 300"
                                        End If
                                    End If

                                    If PO.DestinationPort.Trim.Length > 50 Then
                                        If ls_error = "" Then
                                            ls_error = "Destination Port max length only 50"
                                        Else
                                            ls_error = ls_error + ",Destination Port max length only 50"
                                        End If
                                    End If

                                ls_sql = " INSERT INTO [dbo].[UploadAffiliate] " & vbCrLf &
                                              "            ([AffiliateID], [AffiliateCode], [ConsigneeCode], [BuyerCode], [AffiliateName], [PODeliveryBy], [FolderOES], [Address], [ConsigneeName], [ConsigneeAddress], [BuyerName], [BuyerAddress],[City], [PostalCode], " & vbCrLf &
                                              "             [Phone1],[Phone2],[Fax],[NPWP],[KantorPabean],[IzinTPB],[BCPerson],[ErrorCls], [DestinationPort], [OverseasCls]) " & vbCrLf &
                                              "      VALUES " & vbCrLf &
                                              "            ('" & PO.AffiliateID & "'" & vbCrLf &
                                              "            ,'" & PO.AffiliateCode & "'" & vbCrLf &
                                              "            ,'" & PO.ConsigneeCode & "' " & vbCrLf &
                                              "            ,'" & PO.BuyerCode & "' " & vbCrLf &
                                              "            ,'" & PO.AffiliateName & "' " & vbCrLf &
                                              "            ,'" & PO.PODeliveryBy & "' " & vbCrLf &
                                              "            ,'" & PO.FolderOES & "' " & vbCrLf &
                                              "            ,'" & PO.Address & "' " & vbCrLf &
                                              "            ,'" & PO.ConsigneeName & "' " & vbCrLf &
                                              "            ,'" & PO.ConsigneeAddress & "' " & vbCrLf &
                                              "            ,'" & PO.BuyerName & "' " & vbCrLf &
                                              "            ,'" & PO.BuyerAddress & "' " & vbCrLf &
                                              "            ,'" & PO.City & "' " & vbCrLf &
                                              "            ,'" & PO.PostalCode & "' " & vbCrLf &
                                              "            ,'" & PO.Phone1 & "' " & vbCrLf &
                                              "            ,'" & PO.Phone2 & "' " & vbCrLf &
                                              "            ,'" & PO.Fax & "' " & vbCrLf &
                                              "            ,'" & PO.NPWP & "' " & vbCrLf &
                                              "            ,'" & PO.KantorPabean & "' " & vbCrLf &
                                              "            ,'" & PO.IzinTPB & "' " & vbCrLf &
                                              "            ,'" & PO.BCPerson & "' " & vbCrLf &
                                              "            ,'" & ls_error & "' " & vbCrLf &
                                              "            ,'" & PO.DestinationPort & "' " & vbCrLf &
                                              "            ,'" & PO.OverseasCls & "') "
                                Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm.ExecuteNonQuery()
                                    sqlComm.Dispose()
                                Next
                                sqlTran.Commit()

                                lblInfo.Text = "[7001] Data Checking Done!"
                                lblInfo.ForeColor = Color.Blue
                                grid.JSProperties("cpMessage") = lblInfo.Text

                                Call bindData()
                            End Using
                        Catch ex As Exception
                            lblInfo.Text = ex.Message
                            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                            MyConnection.Close()
                            Exit Sub
                        End Try
                        dt.Reset()
                        dtDetail.Reset()
                        dtHeader.Reset()
                    End Using
                    MyConnection.Close()
                Else
                    If FileName = "" Then
                        lblInfo.Text = "[9999] Please choose the file!"
                        up_GridLoadWhenEventChange()
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Exit Sub
                    End If
                End If
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Protected Sub Uploader_FileUploadComplete(ByVal sender As Object, ByVal e As FileUploadCompleteEventArgs)
        Try
            e.CallbackData = SavePostedFiles(e.UploadedFile)
        Catch ex As Exception
            e.IsValid = False
            lblInfo.Text = ex.Message
        End Try
    End Sub

    Private Function SavePostedFiles(ByVal uploadedFile As UploadedFile) As String
        If (Not uploadedFile.IsValid) Then
            Return String.Empty
        End If

        Ext = Path.Combine(MapPath(""))
        FileName = Uploader.PostedFile.FileName
        FilePath = Ext & "\Import\" & FileName
        uploadedFile.SaveAs(FilePath)

        Return FilePath
    End Function

    Private Sub up_Save()
        Dim i As Integer, j As Integer
        'Dim tampung As String = ""
        Dim ls_Check As Boolean = False
        'Dim ls_PONo As String = ""
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim ls_SupplierID As String = ""
        'Dim ls_Period As Date
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        Dim shostname As String = System.Net.Dns.GetHostName
        Dim ls_Remarks As String = ""
        Dim ls_OverseasCls As String = "", ls_PODel As String = "", ls_xOverseasCls As String = "", ls_xPODel As String = ""
        'Dim ls_DoubleSupplier As Boolean = False
        'Dim ls_TempSupplierID As String = ""
        Try
            '01. Cari ada data yg disubmit
            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "ErrorCls").ToString.Trim <> "" Then
                    ls_Check = True
                    Exit For
                End If
            Next i

            If ls_Check = True Then
                lblInfo.Text = "[9999] Invalid data in this File Upload, please check the file again!"
                Session("YA010IsSubmit") = lblInfo.Text
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
            Dim SqlTran As SqlTransaction

            SqlCon.Open()

            SqlTran = SqlCon.BeginTransaction

            Try
                '2.1 delete data 
                Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                SQLCom.Connection = SqlCon
                SQLCom.Transaction = SqlTran

                '2.2 Insert New Detail Data
                Dim ls_SuppCount As Integer = 1
                For i = 0 To grid.VisibleRowCount - 1
                    If grid.GetRowValues(i, "AffiliateID") <> "" Then
                        ls_Sql = " IF NOT EXISTS (select * from MS_Affiliate where AffiliateID = '" & grid.GetRowValues(i, "AffiliateID") & "')" & vbCrLf &
                                  " BEGIN" & vbCrLf &
                                  " INSERT INTO [dbo].[MS_Affiliate] " & vbCrLf &
                                  "            ([AffiliateID] " & vbCrLf &
                                  "            ,[AffiliateCode] " & vbCrLf &
                                  "            ,[ConsigneeCode] " & vbCrLf &
                                  "            ,[BuyerCode] " & vbCrLf &
                                  "            ,[AffiliateName] " & vbCrLf &
                                  "            ,[Address] " & vbCrLf &
                                  "            ,[ConsigneeName] " & vbCrLf &
                                  "            ,[ConsigneeAddress] " & vbCrLf &
                                  "            ,[BuyerName] " & vbCrLf &
                                  "            ,[BuyerAddress] " & vbCrLf &
                                  "            ,[City] " & vbCrLf &
                                  "            ,[PostalCode] " & vbCrLf &
                                  "            ,[Phone1] " & vbCrLf &
                                  "            ,[Phone2] " & vbCrLf &
                                  "            ,[Fax] " & vbCrLf &
                                  "            ,[NPWP] " & vbCrLf &
                                  "            ,[KantorPabean] " & vbCrLf &
                                  "            ,[IzinTPB] " & vbCrLf &
                                  "            ,[BCPerson] " & vbCrLf &
                                  "            ,[DestinationPort] " & vbCrLf &
                                  "            ,[OverseasCls] " & vbCrLf &
                                  "            ,[PODeliveryBy] "

                        ls_Sql = ls_Sql + "            ,[FolderOES] " & vbCrLf &
                                          "            ,[EntryDate] " & vbCrLf &
                                          "            ,[EntryUser] ) " & vbCrLf &
                                          "      VALUES " & vbCrLf &
                                          "            ('" & grid.GetRowValues(i, "AffiliateID") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "AffiliateCode") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "ConsigneeCode") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "BuyerCode") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "AffiliateName") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "Address") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "ConsigneeName") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "ConsigneeAddress") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "BuyerName") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "BuyerAddress") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "City") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "PostalCode") & "' "

                        ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "Phone1") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "Phone2") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "Fax") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "NPWP") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "KantorPabean") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "IzinTPB") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "BCPerson") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "DestinationPort") & "' " & vbCrLf &
                                          "            ,'" & IIf(grid.GetRowValues(i, "OverseasCls").ToString.ToUpper = "YES", 1, 0) & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "PODeliveryBy") & "' " & vbCrLf &
                                          "            ,'" & grid.GetRowValues(i, "FolderOES") & "' " & vbCrLf &
                                          "            ,getdate() " & vbCrLf &
                                          "            ,'UPLOAD') " & vbCrLf &
                                          " END " & vbCrLf &
                                          " ELSE " & vbCrLf &
                                          " BEGIN" & vbCrLf &
                                          "      UPDATE [dbo].[MS_Affiliate] SET " & vbCrLf &
                                          "       [AffiliateName] = '" & grid.GetRowValues(i, "AffiliateName") & "' " & vbCrLf &
                                          "       ,[AffiliateCode] = '" & grid.GetRowValues(i, "AffiliateCode") & "' " & vbCrLf &
                                          "       ,[ConsigneeCode] = '" & grid.GetRowValues(i, "ConsigneeCode") & "' " & vbCrLf &
                                          "       ,[BuyerCode] = '" & grid.GetRowValues(i, "BuyerCode") & "' " & vbCrLf &
                                          "       ,[Address] = '" & grid.GetRowValues(i, "Address") & "' " & vbCrLf &
                                          "       ,[ConsigneeName] = '" & grid.GetRowValues(i, "ConsigneeName") & "' " & vbCrLf &
                                          "       ,[ConsigneeAddress] = '" & grid.GetRowValues(i, "ConsigneeAddress") & "' " & vbCrLf &
                                          "       ,[BuyerName] = '" & grid.GetRowValues(i, "BuyerName") & "' " & vbCrLf &
                                          "       ,[BuyerAddress] = '" & grid.GetRowValues(i, "BuyerAddress") & "' " & vbCrLf &
                                          "       ,[City] = '" & grid.GetRowValues(i, "City") & "' " & vbCrLf &
                                          "       ,[PostalCode] = '" & grid.GetRowValues(i, "PostalCode") & "' " & vbCrLf &
                                          "       ,[Phone1] = '" & grid.GetRowValues(i, "Phone1") & "' " & vbCrLf &
                                          "       ,[Phone2] = '" & grid.GetRowValues(i, "Phone2") & "' " & vbCrLf &
                                          "       ,[Fax] = '" & grid.GetRowValues(i, "Fax") & "' " & vbCrLf &
                                          "       ,[NPWP] = '" & grid.GetRowValues(i, "NPWP") & "' " & vbCrLf &
                                          "       ,[KantorPabean] = '" & grid.GetRowValues(i, "KantorPabean") & "' " & vbCrLf &
                                          "       ,[IzinTPB] = '" & grid.GetRowValues(i, "IzinTPB") & "' " & vbCrLf &
                                          "       ,[BCPerson] = '" & grid.GetRowValues(i, "BCPerson") & "' " & vbCrLf &
                                          "       ,[PODeliveryBy] = '" & grid.GetRowValues(i, "PODeliveryBy") & "' " & vbCrLf &
                                          "       ,[FolderOES] = '" & grid.GetRowValues(i, "FolderOES") & "' " & vbCrLf &
                                          "       ,[DestinationPort] = '" & grid.GetRowValues(i, "DestinationPort") & "' " & vbCrLf &
                                          "       ,[OverseasCls] = '" & IIf(grid.GetRowValues(i, "OverseasCls").ToString.Trim.ToUpper = "YES", 1, 0) & "' " & vbCrLf &
                                          "      WHERE [AffiliateID] = '" & grid.GetRowValues(i, "AffiliateID") & "' " & vbCrLf &
                                          " END"

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                    End If

                    ls_OverseasCls = grid.GetRowValues(i, "OverseasCls").ToString
                    ls_xOverseasCls = grid.GetRowValues(i, "xOverseasCls").ToString
                    ls_PODel = grid.GetRowValues(i, "PODeliveryBy").ToString
                    ls_xPODel = grid.GetRowValues(i, "xPODeliveryBy").ToString

                    If (Not IsDBNull(grid.GetRowValues(i, "AffiliateName")) And Not IsDBNull(grid.GetRowValues(i, "xAffiliateName"))) And (grid.GetRowValues(i, "AffiliateName").ToString <> "" And grid.GetRowValues(i, "xAffiliateName").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "AffiliateName").ToString) <> Trim(grid.GetRowValues(i, "xAffiliateName").ToString)) Then
                            ls_Remarks = ls_Remarks + "AffiliateName " + Trim(grid.GetRowValues(i, "xAffiliateName").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "AffiliateName").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "Address")) And Not IsDBNull(grid.GetRowValues(i, "xAddress"))) And (grid.GetRowValues(i, "Address").ToString <> "" And grid.GetRowValues(i, "xAddress").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "Address").ToString) <> Trim(grid.GetRowValues(i, "xAddress").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Address " + Trim(grid.GetRowValues(i, "xAddress").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "Address").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "City")) And Not IsDBNull(grid.GetRowValues(i, "xCity"))) And (grid.GetRowValues(i, "City").ToString <> "" And grid.GetRowValues(i, "xCity").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "City").ToString) <> Trim(grid.GetRowValues(i, "xCity").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "City " + Trim(grid.GetRowValues(i, "xCity").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "City").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "PostalCode")) And Not IsDBNull(grid.GetRowValues(i, "xPostalCode"))) And (grid.GetRowValues(i, "PostalCode").ToString <> "" And grid.GetRowValues(i, "xPostalCode").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "PostalCode").ToString) <> Trim(grid.GetRowValues(i, "xPostalCode").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "PostalCode " + Trim(grid.GetRowValues(i, "xPostalCode").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "PostalCode").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "Phone1")) And Not IsDBNull(grid.GetRowValues(i, "xPhone1"))) And (grid.GetRowValues(i, "Phone1").ToString <> "" And grid.GetRowValues(i, "xPhone1").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "Phone1").ToString) <> Trim(grid.GetRowValues(i, "xPhone1").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Phone1 " + Trim(grid.GetRowValues(i, "xPhone1").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "Phone1").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "Phone2")) And Not IsDBNull(grid.GetRowValues(i, "xPhone2"))) And (grid.GetRowValues(i, "Phone2").ToString <> "" And grid.GetRowValues(i, "xPhone2").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "Phone2").ToString) <> Trim(grid.GetRowValues(i, "xPhone2").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Phone2 " + Trim(grid.GetRowValues(i, "xPhone2").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "Phone2").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "Fax")) And Not IsDBNull(grid.GetRowValues(i, "xFax"))) And (grid.GetRowValues(i, "Fax").ToString <> "" And grid.GetRowValues(i, "xFax").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "Fax").ToString) <> Trim(grid.GetRowValues(i, "xFax").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Fax " + Trim(grid.GetRowValues(i, "xFax").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "Fax").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "NPWP")) And Not IsDBNull(grid.GetRowValues(i, "xNPWP"))) And (grid.GetRowValues(i, "NPWP").ToString <> "" And grid.GetRowValues(i, "xNPWP").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "NPWP").ToString) <> Trim(grid.GetRowValues(i, "xNPWP").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "NPWP " + Trim(grid.GetRowValues(i, "xNPWP").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "NPWP").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "KantorPabean")) And Not IsDBNull(grid.GetRowValues(i, "xKantorPabean"))) And (grid.GetRowValues(i, "KantorPabean").ToString <> "" And grid.GetRowValues(i, "xKantorPabean").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "KantorPabean").ToString) <> Trim(grid.GetRowValues(i, "xKantorPabean").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "KantorPabean " + Trim(grid.GetRowValues(i, "xKantorPabean").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "KantorPabean").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "IzinTPB")) And Not IsDBNull(grid.GetRowValues(i, "xIzinTPB"))) And (grid.GetRowValues(i, "IzinTPB").ToString <> "" And grid.GetRowValues(i, "xIzinTPB").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "IzinTPB").ToString) <> Trim(grid.GetRowValues(i, "xIzinTPB").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "IzinTPB " + Trim(grid.GetRowValues(i, "xIzinTPB").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "IzinTPB").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "BCPerson")) And Not IsDBNull(grid.GetRowValues(i, "xBCPerson"))) And (grid.GetRowValues(i, "BCPerson").ToString <> "" And grid.GetRowValues(i, "xBCPerson").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "BCPerson").ToString) <> Trim(grid.GetRowValues(i, "xBCPerson").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "BCPerson " + Trim(grid.GetRowValues(i, "xBCPerson").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "BCPerson").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "PODeliveryBy")) And Not IsDBNull(grid.GetRowValues(i, "xPODeliveryBy"))) And (grid.GetRowValues(i, "PODeliveryBy").ToString <> "" And grid.GetRowValues(i, "xPODeliveryBy").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "PODeliveryBy").ToString) <> Trim(grid.GetRowValues(i, "xPODeliveryBy").ToString)) Then
                            If grid.GetRowValues(i, "xPODeliveryBy").ToString.Trim = "1" Then ls_xPODel = "PASI" Else ls_xPODel = "Supplier"
                            If grid.GetRowValues(i, "PODeliveryBy").ToString.Trim = "1" Then ls_PODel = "PASI" Else ls_PODel = "Supplier"
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "PODeliveryBy " + ls_xPODel & " " & "->" & " " & ls_PODel & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "OverseasCls")) And Not IsDBNull(grid.GetRowValues(i, "xOverseasCls"))) And (grid.GetRowValues(i, "OverseasCls").ToString <> "" And grid.GetRowValues(i, "xOverseasCls").ToString <> "") Then
                        If (IIf(Trim(grid.GetRowValues(i, "OverseasCls").ToString.Trim).ToUpper = "YES", "1", "0") <> Trim(grid.GetRowValues(i, "xOverseasCls").ToString)) Then
                            If grid.GetRowValues(i, "xOverseasCls").ToString.Trim = "1" Then ls_xOverseasCls = "YES" Else ls_xOverseasCls = "NO"
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "OverseasCls " + ls_xOverseasCls & " " & "->" & " " & grid.GetRowValues(i, "OverseasCls").ToString.Trim & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "ConsigneeCode")) And Not IsDBNull(grid.GetRowValues(i, "xConsigneeCode"))) And (grid.GetRowValues(i, "ConsigneeCode").ToString <> "" And grid.GetRowValues(i, "xConsigneeCode").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "ConsigneeCode").ToString) <> Trim(grid.GetRowValues(i, "xConsigneeCode").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "ConsigneeCode " + Trim(grid.GetRowValues(i, "xConsigneeCode").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "ConsigneeCode").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "ConsigneeName")) And Not IsDBNull(grid.GetRowValues(i, "xConsigneeName"))) And (grid.GetRowValues(i, "ConsigneeName").ToString <> "" And grid.GetRowValues(i, "xConsigneeName").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "ConsigneeName").ToString) <> Trim(grid.GetRowValues(i, "xConsigneeName").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "ConsigneeName " + Trim(grid.GetRowValues(i, "xConsigneeName").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "ConsigneeName").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "ConsigneeAddress")) And Not IsDBNull(grid.GetRowValues(i, "xConsigneeAddress"))) And (grid.GetRowValues(i, "ConsigneeAddress").ToString <> "" And grid.GetRowValues(i, "xConsigneeAddress").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "ConsigneeAddress").ToString) <> Trim(grid.GetRowValues(i, "xConsigneeAddress").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "ConsigneeAddress " + Trim(grid.GetRowValues(i, "xConsigneeAddress").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "ConsigneeAddress").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "BuyerCode")) And Not IsDBNull(grid.GetRowValues(i, "xBuyerCode"))) And (grid.GetRowValues(i, "BuyerCode").ToString <> "" And grid.GetRowValues(i, "xBuyerCode").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "BuyerCode").ToString) <> Trim(grid.GetRowValues(i, "xBuyerCode").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "BuyerCode " + Trim(grid.GetRowValues(i, "xBuyerCode").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "BuyerCode").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "BuyerName")) And Not IsDBNull(grid.GetRowValues(i, "xBuyerName"))) And (grid.GetRowValues(i, "BuyerName").ToString <> "" And grid.GetRowValues(i, "xBuyerName").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "BuyerName").ToString) <> Trim(grid.GetRowValues(i, "xBuyerName").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "BuyerName " + Trim(grid.GetRowValues(i, "xBuyerName").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "BuyerName").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "BuyerAddress")) And Not IsDBNull(grid.GetRowValues(i, "xBuyerAddress"))) And (grid.GetRowValues(i, "BuyerAddress").ToString <> "" And grid.GetRowValues(i, "xBuyerAddress").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "BuyerAddress").ToString) <> Trim(grid.GetRowValues(i, "xBuyerAddress").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "BuyerAddress " + Trim(grid.GetRowValues(i, "xBuyerAddress").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "BuyerAddress").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "DestinationPort")) And Not IsDBNull(grid.GetRowValues(i, "xDestinationPort"))) And (grid.GetRowValues(i, "DestinationPort").ToString <> "" And grid.GetRowValues(i, "xDestinationPort").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "DestinationPort").ToString) <> Trim(grid.GetRowValues(i, "xDestinationPort").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "DestinationPort " + Trim(grid.GetRowValues(i, "xDestinationPort").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "DestinationPort").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "FolderOES")) And Not IsDBNull(grid.GetRowValues(i, "xFolderOES"))) And (grid.GetRowValues(i, "FolderOES").ToString <> "" And grid.GetRowValues(i, "xFolderOES").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "FolderOES").ToString) <> Trim(grid.GetRowValues(i, "xFolderOES").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "FolderOES " + Trim(grid.GetRowValues(i, "xFolderOES").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "FolderOES").ToString) & ""
                        End If
                    End If

                    If ls_Remarks <> "" Then
                        'insert into history
                        ls_Sql = " INSERT INTO MS_History (PCName, MenuID, OperationID, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                                 "VALUES ('" & shostname & "','" & menuID & "','U', 'Update [" & ls_Remarks & "]', " & vbCrLf & _
                                 "GETDATE(), '" & Session("UserID") & "')  "
                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                        ls_Remarks = ""
                    End If
                Next i

                '2.3.1 Habis save semua,.. delete tada di tempolary table
                ls_Sql = "delete UploadAffiliate "

                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                '2.3.3 Commit transaction
                ls_MsgID = "1001"
                SqlTran.Commit()
                Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                Session("YA010IsSubmit") = lblInfo.Text
            Catch ex As Exception
                Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                Session("YA010IsSubmit") = lblInfo.Text
                SqlTran.Rollback()
                SqlCon.Close()
                Exit Sub
            End Try

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("YA010IsSubmit") = lblInfo.Text
        End Try
    End Sub

#End Region
End Class