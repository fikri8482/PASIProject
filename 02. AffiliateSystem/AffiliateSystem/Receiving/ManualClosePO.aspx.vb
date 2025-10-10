Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing


Public Class ManualClosePO
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = ""

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "E03"
#End Region

#Region "Procedures"
    Private Sub up_Initialize()
        Dim script As String = _
            "var a = new Date(); " & vbCrLf & _
            "dtSupplierPeriod.SetDate(a); " & vbCrLf & _
            "txtPONo.SetText(''); " & vbCrLf & _
            "cboSupplier.SetValue('==ALL=='); " & vbCrLf & _
            "txtSupplierName.SetText('==ALL=='); " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(chkSupplierPeriod, chkSupplierPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""

            ls_SQL = ls_SQL + " SELECT ColNo = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PONo)), " & vbCrLf & _
                              " 	   MCP.* " & vbCrLf & _
                              "   FROM ( SELECT DISTINCT " & vbCrLf & _
                              "  				Period = SUBSTRING(CONVERT(CHAR,POM.Period,106),4,9),  " & vbCrLf & _
                              "  				PONo = POD.PONo,   " & vbCrLf & _
                              "  				--DeliveryLocationCode = ISNULL(KM.DeliveryLocationCode,''),  " & vbCrLf & _
                              " 				--DeliveryLocationName = ISNULL(MD.DeliveryLocationName,''),  " & vbCrLf & _
                              "  				SupplierCode = POM.SupplierID,  " & vbCrLf & _
                              " 				SupplierName = MS.SupplierName,  " & vbCrLf & _
                              "  				PartNo = RTRIM(POD.PartNo),  " & vbCrLf & _
                              " 				PartName = RTRIM(MP.PartName),  " & vbCrLf & _
                              "  				CloseCls = ISNULL(POD.CloseCls,'0'),  " & vbCrLf & _
                              "  				CloseDate = ISNULL(CONVERT(CHAR,POD.CloseDate,106),''),  " & vbCrLf
            ls_SQL = ls_SQL + "  				SupplierPIC = ISNULL(RTRIM(POD.CloseSupplierPIC),'') " & vbCrLf & _
                              " 		   FROM PO_DETAIL POD " & vbCrLf & _
                              "  				LEFT JOIN PO_Master POM ON POM.AffiliateID =POD.AffiliateID  " & vbCrLf & _
                              "  					AND POM.SupplierID =POD.SupplierID  " & vbCrLf & _
                              "  					AND POM.PONO =POD.PONO  " & vbCrLf & _
                              "  				LEFT JOIN Kanban_Detail KD ON KD.AffiliateID =POD.AffiliateID  " & vbCrLf & _
                              "  					AND KD.SupplierID =POD.SupplierID  " & vbCrLf & _
                              "  					AND KD.PONO =POD.PONO  " & vbCrLf & _
                              "  					AND KD.PartNo =POD.PartNo  " & vbCrLf & _
                              "  				LEFT JOIN Kanban_Master KM ON KD.AffiliateID =KM.AffiliateID  " & vbCrLf & _
                              "  					AND KD.SupplierID =KM.SupplierID  " & vbCrLf
            ls_SQL = ls_SQL + "  					AND KD.KanbanNo =KM.KanbanNo  " & vbCrLf & _
                              "  				LEFT JOIN DOSupplier_Detail DSD ON KD.AffiliateID =DSD.AffiliateID  " & vbCrLf & _
                              "  					AND KD.SupplierID =DSD.SupplierID  " & vbCrLf & _
                              "  					AND KD.PONO =DSD.PONO  " & vbCrLf & _
                              "  					AND KD.PartNo =DSD.PartNo  " & vbCrLf & _
                              "  					AND KD.KanbanNo =DSD.KanbanNo  " & vbCrLf & _
                              "  				LEFT JOIN DOSupplier_Master DSM ON DSM.AffiliateID =DSD.AffiliateID  " & vbCrLf & _
                              "  					AND DSM.SupplierID =DSD.SupplierID  " & vbCrLf & _
                              "  					AND DSM.SuratJalanNo =DSD.SuratJalanNo  " & vbCrLf & _
                              "  				LEFT JOIN DOPASI_Detail DPD ON KD.AffiliateID =DPD.AffiliateID  " & vbCrLf & _
                              "  					AND KD.SupplierID =DPD.SupplierID  " & vbCrLf
            ls_SQL = ls_SQL + "  					AND KD.PONO =DPD.PONO  " & vbCrLf & _
                              "  					AND KD.PartNo =DPD.PartNo  " & vbCrLf & _
                              "  					AND KD.KanbanNo =DPD.KanbanNo  " & vbCrLf & _
                              "  				LEFT JOIN DOPASI_Master DPM ON DPM.AffiliateID =DPD.AffiliateID  " & vbCrLf & _
                              "  					AND DPM.SupplierID =DPD.SupplierID  " & vbCrLf & _
                              "  					AND DPM.SuratJalanNo =DPD.SuratJalanNo  " & vbCrLf & _
                              "  				LEFT JOIN ReceivePASI_Detail RPD ON RPD.AffiliateID = DPM.AffiliateID  " & vbCrLf & _
                              "  					AND RPD.SupplierID = DPM.SupplierID  " & vbCrLf & _
                              "  					AND RPD.PONo = POD.PONo  " & vbCrLf & _
                              "  					AND RPD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  					AND RPD.KanbanNo = KD.KanbanNo  " & vbCrLf
            ls_SQL = ls_SQL + "  				LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = RPD.AffiliateID  " & vbCrLf & _
                              "  					AND RAM.SupplierID = RPD.SupplierID  " & vbCrLf & _
                              "  				LEFT JOIN ReceiveAffiliate_Detail RAD ON RAD.AffiliateID = RPD.AffiliateID  " & vbCrLf & _
                              "  					AND RAD.SupplierID = RPD.SupplierID  " & vbCrLf & _
                              "  					AND RAD.SuratJalanNo = RAM.SuratJalanNo  " & vbCrLf & _
                              "  				LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                              "  				LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls  " & vbCrLf & _
                              " 				LEFT JOIN MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                              " 				LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
                              " 		  WHERE POD.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf
            'SUPPLIER PLAN DELIVERY DATE (UNTIL)
            If chkSupplierPeriod.Checked = True Then
                ls_SQL = ls_SQL + _
                              "              AND CONVERT(DATETIME,KM.KanbanDate) <= CONVERT(DATETIME,'" & Format(dtSupplierPeriod.Value, "yyyy-MM-dd") & "') " & vbCrLf
            End If
            'SUPPLIER CODE/NAME
            If Trim(cboSupplier.Text) <> "==ALL==" And Trim(cboSupplier.Text) <> "" Then
                ls_SQL = ls_SQL + _
                              "              AND POM.SupplierID = '" & Trim(cboSupplier.Text) & "' " & vbCrLf
            End If
            'PO NO.
            If Trim(txtPONo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                              "              AND ISNULL(POM.PONo,'') = '" & Trim(txtPONo.Text) & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + "        ) MCP " & vbCrLf 


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "SELECT TOP 0  " & vbCrLf & _
                     " 		 ColNo = 0, Period = '', PONo = '', /*DeliveryLocationCode = '', DeliveryLocationName = '',*/ SupplierCode = '', SupplierName = '', PartNo = '', PartName = '', " & vbCrLf & _
                     " 		 CloseCls = '', CloseDate = '', SupplierPIC = '' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        'Combo Supplier
        With cboSupplier
            ls_SQL = "SELECT SupplierID = '==ALL==', SupplierName = '==ALL=='" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     "SELECT SupplierID, SupplierName FROM dbo.MS_Supplier " & vbCrLf
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierID")
                .Columns(0).Width = 90
                .Columns.Add("SupplierName")
                .Columns(1).Width = 240

                .TextField = "SupplierID"
                .DataBind()
            End Using
        End With
    End Sub
#End Region

#Region "Form Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_FillCombo()
                Call up_GridLoadWhenEventChange()
                Call up_Initialize()
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("E03Msg")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_BatchUpdate(sender As Object, e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Try
            Dim ls_CloseCls As String = "0", ls_SupplierPIC As String = "", ls_CloseDate As String, _
                ls_PONo As String = "", ls_AffilateID As String = Session("AffiliateID"), ls_SupplierID As String = "", _
                ls_PartNo As String = ""
            Dim iLoop As Integer = 0

            Dim sqlCmd As New SqlCommand()

            Using scope As New TransactionScope

                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                    sqlConn.Open()

                    For iLoop = 0 To e.UpdateValues.Count - 1
                        ls_PartNo = e.UpdateValues(iLoop).NewValues("PartNo").ToString()
                        ls_CloseCls = e.UpdateValues(iLoop).NewValues("CloseCls").ToString()
                        If ls_CloseCls = "1" Then ls_CloseDate = "GETDATE() " Else ls_CloseDate = "NULL "
                        ls_PONo = e.UpdateValues(iLoop).NewValues("PONo").ToString().Trim()
                        ls_SupplierID = e.UpdateValues(iLoop).NewValues("SupplierCode").ToString().Trim()
                        ls_SupplierPIC = e.UpdateValues(iLoop).NewValues("SupplierPIC").ToString().Trim()

                        ls_SQL = "UPDATE dbo.PO_Detail " & vbCrLf & _
                                 "   SET CloseCls = '" & ls_CloseCls & "', CloseDate = " & ls_CloseDate & ", CloseSupplierPIC = '" & ls_SupplierPIC & "' " & vbCrLf & _
                                 " WHERE PONo = '" & ls_PONo & "'" & vbCrLf & _
                                 "   AND SupplierID = '" & ls_SupplierID & "'" & vbCrLf & _
                                 "   AND AffiliateID = '" & ls_AffilateID & "'" & vbCrLf & _
                                 "   AND PartNo = '" & ls_PartNo & "'"
                        sqlCmd = New SqlCommand(ls_SQL, sqlConn)
                        sqlCmd.ExecuteNonQuery()
                        sqlCmd.Dispose()
                    Next iLoop

                    Call clsMsg.DisplayMessage(lblInfo, "1002", clsMessage.MsgType.InformationMessage)
                    Session("E03Msg") = lblInfo.Text

                End Using
                scope.Complete()
            End Using

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("E03Msg") = lblInfo.Text
        End Try
    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        Session("E03Msg") = lblInfo.Text
                    Else
                        grid.PageIndex = 0
                    End If
                Case "clear"
                    Call up_GridLoadWhenEventChange()
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("E03Msg") = lblInfo.Text
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
        If (Not IsNothing(Session("E03Msg"))) Then grid.JSProperties("cpMessage") = Session("E03Msg") : Session.Remove("E03Msg")
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If Not (e.DataColumn.FieldName = "CloseCls" Or _
            e.DataColumn.FieldName = "SupplierPIC") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        Else
            e.Cell.BackColor = Color.White
        End If
    End Sub

    Private Sub grid_PageIndexChanged(sender As Object, e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region

End Class