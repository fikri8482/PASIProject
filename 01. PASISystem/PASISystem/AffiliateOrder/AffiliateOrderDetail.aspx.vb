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


Public Class AffiliateOrderDetail
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "B02"
    Dim pSearch As Boolean = False
    Dim remain As Double
    Dim TotQtyAff As Double
    Dim TotQtyPASI As Double
    Dim ls_POqty As Double

    Dim ls_DeliveryD1Old As Double : Dim ls_DeliveryD2Old As Double : Dim ls_DeliveryD3Old As Double : Dim ls_DeliveryD4Old As Double : Dim ls_DeliveryD5Old As Double
    Dim ls_DeliveryD6Old As Double : Dim ls_DeliveryD7Old As Double : Dim ls_DeliveryD8Old As Double : Dim ls_DeliveryD9Old As Double : Dim ls_DeliveryD10Old As Double
    Dim ls_DeliveryD11Old As Double : Dim ls_DeliveryD12Old As Double : Dim ls_DeliveryD13Old As Double : Dim ls_DeliveryD14Old As Double : Dim ls_DeliveryD15Old As Double
    Dim ls_DeliveryD16Old As Double : Dim ls_DeliveryD17Old As Double : Dim ls_DeliveryD18Old As Double : Dim ls_DeliveryD19Old As Double : Dim ls_DeliveryD20Old As Double
    Dim ls_DeliveryD21Old As Double : Dim ls_DeliveryD22Old As Double : Dim ls_DeliveryD23Old As Double : Dim ls_DeliveryD24Old As Double : Dim ls_DeliveryD25Old As Double
    Dim ls_DeliveryD26Old As Double : Dim ls_DeliveryD27Old As Double : Dim ls_DeliveryD28Old As Double : Dim ls_DeliveryD29Old As Double : Dim ls_DeliveryD30Old As Double
    Dim ls_DeliveryD31Old As Double

    Dim errorBatch As Boolean
    Dim UpdateSend As Boolean = False

#End Region

#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

            Dim tmpPO As String, tmpAffiliateID As String, tmpSupplierID As String
            Dim tmpDate As Date

            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Session("M01Url") <> "" Then
                    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                        Session("MenuDesc") = "AFFILIATE ORDER DETAIL ENTRY"
                        up_Fillcombo()

                        tmpPO = (Request.QueryString("id"))
                        tmpAffiliateID = (Request.QueryString("t1"))
                        tmpSupplierID = (Request.QueryString("t2"))
                        tmpDate = (Request.QueryString("t3"))

                        pSearch = False
                        bindDataHeader(tmpAffiliateID, tmpPO, tmpSupplierID)
                        bindDataDetail(tmpDate, tmpAffiliateID, tmpPO, tmpSupplierID)
                        Call SaveDataMaster(tmpAffiliateID, tmpPO, tmpSupplierID)
                        Call SaveDataDetail(tmpAffiliateID, tmpPO, tmpSupplierID)
                        SaveDeliveryCls()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        pSearch = True
                        If txtSuppAppDate.Text <> "" Or txtSuppPendDate.Text <> "" Or txtSuppUnpDate.Text <> "" Or txtPASIAppDate.Text <> "" Or txtAffFinalAppDate.Text <> "" Then
                            btnSubmit.Enabled = False
                            btnSendSupplier.Enabled = False
                        End If                        
                    ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                        Session("MenuDesc") = "AFFILIATE ORDER DETAIL ENTRY"
                        up_Fillcombo()

                        tmpPO = clsNotification.DecryptURL(Request.QueryString("id2"))
                        tmpAffiliateID = clsNotification.DecryptURL(Request.QueryString("t1"))
                        tmpSupplierID = clsNotification.DecryptURL(Request.QueryString("t2"))
                        tmpDate = clsNotification.DecryptURL(Request.QueryString("t3"))

                        pSearch = False
                        bindDataHeader(tmpAffiliateID, tmpPO, tmpSupplierID)
                        bindDataDetail(tmpDate, tmpAffiliateID, tmpPO, tmpSupplierID)
                        Call SaveDataMaster(tmpAffiliateID, tmpPO, tmpSupplierID)
                        Call SaveDataDetail(tmpAffiliateID, tmpPO, tmpSupplierID)
                        SaveDeliveryCls()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        pSearch = True
                        If txtSuppAppDate.Text <> "" Or txtSuppPendDate.Text <> "" Or txtSuppUnpDate.Text <> "" Or txtPASIAppDate.Text <> "" Or txtAffFinalAppDate.Text <> "" Then
                            btnSubmit.Enabled = False
                            btnSendSupplier.Enabled = False
                        End If
                    Else
                        Session("MenuDesc") = "AFFILIATE ORDER DETAIL ENTRY"
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        btnClear.Visible = True
                    End If
                Else
                    dtPeriod.Value = Now
                    btnClear.Visible = True
                End If
            End If

            If ls_AllowUpdate = False Then btnSubmit.Enabled = False

            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MenuID As String = "", ls_MsgID As String = ""
        Dim iLoop As Long = 0, jLoop As Long = 0
        Dim ls_UserID As String = ""
        'Cek Sudah Approve atau belum

        'If getApp(Trim(cboPONo.Text)) = True Then
        '    ls_MsgID = "6029"
        '    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
        '    Session("ZZ010Msg") = lblInfo.Text
        '    Exit Sub
        'End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("item_code")

                ls_SQL = "Update PO_Master set DeliveryByPASICls = '" & IIf(rblDelivery.Value = 0, 0, 1) & "' where PONo = '" & cboPONo.Text & "' and AffiliateID = '" & cboAffiliateCode.Text & "'"

                Dim SqlComm6 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm6.ExecuteNonQuery()
                SqlComm6.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("UserMenu")

                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    Session("ZZ010Msg") = lblInfo.Text
                    Exit Sub
                End If

                For iLoop = 0 To e.UpdateValues.Count - 1
                    If (e.UpdateValues(iLoop).NewValues("POQty") Mod e.UpdateValues(iLoop).NewValues("MinOrderQty")) <> 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "5005", clsMessage.MsgType.ErrorMessage)
                        Session("YA010IsSubmit") = lblInfo.Text
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        errorBatch = True
                        Exit Sub
                    End If

                    Dim ls_PartNo As String = (e.UpdateValues(iLoop).OldValues("PartNos").ToString())
                    Dim ls_Kanban As String = (e.UpdateValues(iLoop).OldValues("KanbanCls").ToString())
                    If ls_Kanban = "YES" Then ls_Kanban = "1" Else ls_Kanban = "0"

                    ls_POqty = e.UpdateValues(iLoop).NewValues("POQty")

                    Dim ls_ForeCast1 As Double = e.UpdateValues(iLoop).OldValues("ForecastN1")
                    Dim ls_ForeCast2 As Double = e.UpdateValues(iLoop).OldValues("ForecastN2")
                    Dim ls_ForeCast3 As Double = e.UpdateValues(iLoop).OldValues("ForecastN3")
                    Dim ls_DeliveryD1 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD1")
                    Dim ls_DeliveryD2 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD2")
                    Dim ls_DeliveryD3 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD3")
                    Dim ls_DeliveryD4 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD4")
                    Dim ls_DeliveryD5 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD5")
                    Dim ls_DeliveryD6 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD6")
                    Dim ls_DeliveryD7 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD7")
                    Dim ls_DeliveryD8 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD8")
                    Dim ls_DeliveryD9 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD9")
                    Dim ls_DeliveryD10 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD10")
                    Dim ls_DeliveryD11 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD11")
                    Dim ls_DeliveryD12 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD12")
                    Dim ls_DeliveryD13 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD13")
                    Dim ls_DeliveryD14 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD14")
                    Dim ls_DeliveryD15 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD15")
                    Dim ls_DeliveryD16 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD16")
                    Dim ls_DeliveryD17 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD17")
                    Dim ls_DeliveryD18 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD18")
                    Dim ls_DeliveryD19 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD19")
                    Dim ls_DeliveryD20 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD20")
                    Dim ls_DeliveryD21 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD21")
                    Dim ls_DeliveryD22 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD22")
                    Dim ls_DeliveryD23 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD23")
                    Dim ls_DeliveryD24 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD24")
                    Dim ls_DeliveryD25 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD25")
                    Dim ls_DeliveryD26 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD26")
                    Dim ls_DeliveryD27 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD27")
                    Dim ls_DeliveryD28 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD28")
                    Dim ls_DeliveryD29 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD29")
                    Dim ls_DeliveryD30 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD30")
                    Dim ls_DeliveryD31 As Double = e.UpdateValues(iLoop).NewValues("DeliveryD31")


                    ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.Affiliate_Detail WHERE PONo='" & Trim(cboPONo.Text) & "' AND AffiliateID='" & Trim(cboAffiliateCode.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' AND PartNo='" & ls_PartNo & "')  " & vbCrLf & _
                              " BEGIN  " & vbCrLf & _
                              " 	INSERT INTO dbo.Affiliate_Detail " & vbCrLf & _
                              "         ( PONo , " & vbCrLf & _
                              "           AffiliateID , " & vbCrLf & _
                              "           SupplierID , " & vbCrLf & _
                              "           PartNo , " & vbCrLf & _                              
                              "           --KanbanCls , " & vbCrLf & _                              
                              "           POQty , " & vbCrLf

                    ls_SQL = ls_SQL + "           --POQtyOld , " & vbCrLf & _                                     
                                      "           DeliveryD1 , " & vbCrLf & _
                                      "           --DeliveryD1Old , " & vbCrLf & _
                                      "           DeliveryD2 , " & vbCrLf & _
                                      "           --DeliveryD2Old , " & vbCrLf & _
                                      "           DeliveryD3 , " & vbCrLf & _
                                      "           --DeliveryD3Old , " & vbCrLf & _
                                      "           DeliveryD4 , " & vbCrLf

                    ls_SQL = ls_SQL + "           --DeliveryD4Old , " & vbCrLf & _
                                      "           DeliveryD5 , " & vbCrLf & _
                                      "           --DeliveryD5Old , " & vbCrLf & _
                                      "           DeliveryD6 , " & vbCrLf & _
                                      "           --DeliveryD6Old , " & vbCrLf & _
                                      "           DeliveryD7 , " & vbCrLf & _
                                      "           --DeliveryD7Old , " & vbCrLf & _
                                      "           DeliveryD8 , " & vbCrLf & _
                                      "           --DeliveryD8Old , " & vbCrLf & _
                                      "           DeliveryD9 , " & vbCrLf & _
                                      "           --DeliveryD9Old , " & vbCrLf

                    ls_SQL = ls_SQL + "           DeliveryD10 , " & vbCrLf & _
                                      "           --DeliveryD10Old , " & vbCrLf & _
                                      "           DeliveryD11 , " & vbCrLf & _
                                      "           --DeliveryD11Old , " & vbCrLf & _
                                      "           DeliveryD12 , " & vbCrLf & _
                                      "           --DeliveryD12Old , " & vbCrLf & _
                                      "           DeliveryD13 , " & vbCrLf & _
                                      "           --DeliveryD13Old , " & vbCrLf & _
                                      "           DeliveryD14 , " & vbCrLf & _
                                      "           --DeliveryD14Old , " & vbCrLf & _
                                      "           DeliveryD15 , " & vbCrLf

                    ls_SQL = ls_SQL + "           --DeliveryD15Old , " & vbCrLf & _
                                      "           DeliveryD16 , " & vbCrLf & _
                                      "           --DeliveryD16Old , " & vbCrLf & _
                                      "           DeliveryD17 , " & vbCrLf & _
                                      "           --DeliveryD17Old , " & vbCrLf & _
                                      "           DeliveryD18 , " & vbCrLf & _
                                      "           --DeliveryD18Old , " & vbCrLf & _
                                      "           DeliveryD19 , " & vbCrLf & _
                                      "           --DeliveryD19Old , " & vbCrLf & _
                                      "           DeliveryD20 , " & vbCrLf & _
                                      "           --DeliveryD20Old , " & vbCrLf

                    ls_SQL = ls_SQL + "           DeliveryD21 , " & vbCrLf & _
                                      "           --DeliveryD21Old , " & vbCrLf & _
                                      "           DeliveryD22 , " & vbCrLf & _
                                      "           --DeliveryD22Old , " & vbCrLf & _
                                      "           DeliveryD23 , " & vbCrLf & _
                                      "           --DeliveryD23Old , " & vbCrLf & _
                                      "           DeliveryD24 , " & vbCrLf & _
                                      "           --DeliveryD24Old , " & vbCrLf & _
                                      "           DeliveryD25 , " & vbCrLf & _
                                      "           --DeliveryD25Old , " & vbCrLf & _
                                      "           DeliveryD26 , " & vbCrLf

                    ls_SQL = ls_SQL + "           --DeliveryD26Old , " & vbCrLf & _
                                      "           DeliveryD27 , " & vbCrLf & _
                                      "           --DeliveryD27Old , " & vbCrLf & _
                                      "           DeliveryD28 , " & vbCrLf & _
                                      "           --DeliveryD28Old , " & vbCrLf & _
                                      "           DeliveryD29 , " & vbCrLf & _
                                      "           --DeliveryD29Old , " & vbCrLf & _
                                      "           DeliveryD30 , " & vbCrLf & _
                                      "           --DeliveryD30Old , " & vbCrLf & _
                                      "           DeliveryD31 , " & vbCrLf & _
                                      "           --DeliveryD31Old , " & vbCrLf

                    ls_SQL = ls_SQL + "           EntryDate , " & vbCrLf & _
                                      "           EntryUser , " & vbCrLf & _
                                      "           UpdateDate , " & vbCrLf & _
                                      "           UpdateUser " & vbCrLf & _
                                      "         ) " & vbCrLf & _
                                      " 	VALUES  ( '" & Trim(cboPONo.Text) & "' , -- PONo - char(20) " & vbCrLf & _
                                      "           '" & Trim(cboAffiliateCode.Text) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                      "           '" & Trim(txtSupplierCode.Text) & "' , -- SupplierID - char(20) " & vbCrLf & _
                                      "           '" & ls_PartNo & "' , -- PartNo - char(25) " & vbCrLf & _
                                      "           --'" & ls_Kanban & "' , -- KanbanCls - char(1) " & vbCrLf

                    ls_SQL = ls_SQL + "           " & ls_POqty & " , -- POQty - numeric " & vbCrLf & _
                                      "           --" & ls_POqty & " , -- POQtyOld - numeric " & vbCrLf & _                                     
                                      "           " & ls_DeliveryD1 & " , -- DeliveryD1 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD1") & " , -- DeliveryD1Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD2 & " , -- DeliveryD2 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD2") & " , -- DeliveryD2Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD3 & " , -- DeliveryD3 - numeric " & vbCrLf

                    ls_SQL = ls_SQL + "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD3") & " , -- DeliveryD3Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD4 & " , -- DeliveryD4 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD4") & " , -- DeliveryD4Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD5 & " , -- DeliveryD5 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD5") & " , -- DeliveryD5Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD6 & " , -- DeliveryD6 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD6") & " , -- DeliveryD6Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD7 & " , -- DeliveryD7 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD7") & " , -- DeliveryD7Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD8 & " , -- DeliveryD8 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD8") & " , -- DeliveryD8Old - numeric " & vbCrLf

                    ls_SQL = ls_SQL + "           " & ls_DeliveryD9 & " , -- DeliveryD9 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD9") & " , -- DeliveryD9Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD10 & " , -- DeliveryD10 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD10") & " , -- DeliveryD10Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD11 & " , -- DeliveryD11 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD11") & " , -- DeliveryD11Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD12 & " , -- DeliveryD12 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD12") & " , -- DeliveryD12Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD13 & " , -- DeliveryD13 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD13") & " , -- DeliveryD13Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD14 & " , -- DeliveryD14 - numeric " & vbCrLf

                    ls_SQL = ls_SQL + "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD14") & " , -- DeliveryD14Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD15 & " , -- DeliveryD15 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD15") & " , -- DeliveryD15Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD16 & " , -- DeliveryD16 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD16") & " , -- DeliveryD16Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD17 & " , -- DeliveryD17 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD17") & " , -- DeliveryD17Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD18 & " , -- DeliveryD18 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD18") & " , -- DeliveryD18Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD19 & " , -- DeliveryD19 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD19") & " , -- DeliveryD19Old - numeric " & vbCrLf

                    ls_SQL = ls_SQL + "           " & ls_DeliveryD20 & " , -- DeliveryD20 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD20") & " , -- DeliveryD20Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD21 & " , -- DeliveryD21 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD21") & " , -- DeliveryD21Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD22 & " , -- DeliveryD22 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD22") & " , -- DeliveryD22Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD23 & " , -- DeliveryD23 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD23") & " , -- DeliveryD23Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD24 & " , -- DeliveryD24 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD24") & " , -- DeliveryD24Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD25 & " , -- DeliveryD25 - numeric " & vbCrLf

                    ls_SQL = ls_SQL + "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD25") & " , -- DeliveryD25Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD26 & " , -- DeliveryD26 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD26") & " , -- DeliveryD26Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD27 & " , -- DeliveryD27 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD27") & " , -- DeliveryD27Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD28 & " , -- DeliveryD28 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD28") & " , -- DeliveryD28Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD29 & " , -- DeliveryD29 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD29") & " , -- DeliveryD29Old - numeric " & vbCrLf & _
                                      "           " & ls_DeliveryD30 & " , -- DeliveryD30 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD30") & " , -- DeliveryD30Old - numeric " & vbCrLf

                    ls_SQL = ls_SQL + "           " & ls_DeliveryD31 & " , -- DeliveryD31 - numeric " & vbCrLf & _
                                      "           --" & e.UpdateValues(iLoop).OldValues("DeliveryD31") & " , -- DeliveryD31Old - numeric " & vbCrLf & _
                                      "           getdate() , -- EntryDate - datetime " & vbCrLf & _
                                      "           '" & Session("UserID") & "' , -- EntryUser - char(15) " & vbCrLf & _
                                      "           getdate() , -- UpdateDate - datetime " & vbCrLf & _
                                      "           '" & Session("UserID") & "'  -- UpdateUser - char(15) " & vbCrLf & _
                                      "         ) " & vbCrLf & _
                                      "         END	 " & vbCrLf & _
                                      "         ELSE	 " & vbCrLf & _
                                      "         BEGIN  " & vbCrLf & _
                                      "            UPDATE [dbo].[Affiliate_Detail] " & vbCrLf

                    ls_SQL = ls_SQL + " 		   SET [POQty] = " & ls_POqty & " " & vbCrLf & _
                                      " 			  --,[POQtyOld] = " & ls_POqty & " " & vbCrLf & _                                      
                                      " 			  ,[DeliveryD1] = " & ls_DeliveryD1 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD1Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD1") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD2] = " & ls_DeliveryD2 & " " & vbCrLf

                    ls_SQL = ls_SQL + " 			  --,[DeliveryD2Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD2") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD3] = " & ls_DeliveryD3 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD3Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD3") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD4] = " & ls_DeliveryD4 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD4Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD4") & "" & vbCrLf & _
                                      " 			  ,[DeliveryD5] = " & ls_DeliveryD5 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD5Old] =" & e.UpdateValues(iLoop).OldValues("DeliveryD5") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD6] =  " & ls_DeliveryD6 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD6Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD6") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD7] =  " & ls_DeliveryD7 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD7Old] =" & e.UpdateValues(iLoop).OldValues("DeliveryD7") & " " & vbCrLf

                    ls_SQL = ls_SQL + " 			  ,[DeliveryD8] =  " & ls_DeliveryD8 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD8Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD8") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD9] =  " & ls_DeliveryD9 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD9Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD9") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD10] = " & ls_DeliveryD10 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD10Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD10") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD11] = " & ls_DeliveryD11 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD11Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD11") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD12] =  " & ls_DeliveryD12 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD12Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD12") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD13] = " & ls_DeliveryD13 & " " & vbCrLf

                    ls_SQL = ls_SQL + " 			  --,[DeliveryD13Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD13") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD14] =  " & ls_DeliveryD14 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD14Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD14") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD15] =  " & ls_DeliveryD15 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD15Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD15") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD16] =  " & ls_DeliveryD16 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD16Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD16") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD17] = " & ls_DeliveryD17 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD17Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD17") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD18] =  " & ls_DeliveryD18 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD18Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD18") & " " & vbCrLf

                    ls_SQL = ls_SQL + " 			  ,[DeliveryD19] =  " & ls_DeliveryD19 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD19Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD19") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD20] =  " & ls_DeliveryD20 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD20Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD20") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD21] =  " & ls_DeliveryD21 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD21Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD21") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD22] = " & ls_DeliveryD22 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD22Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD22") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD23] =  " & ls_DeliveryD23 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD23Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD23") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD24] =  " & ls_DeliveryD24 & " " & vbCrLf

                    ls_SQL = ls_SQL + " 			  --,[DeliveryD24Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD24") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD25] =  " & ls_DeliveryD25 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD25Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD25") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD26] =  " & ls_DeliveryD26 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD26Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD26") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD27] =  " & ls_DeliveryD27 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD27Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD27") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD28] =  " & ls_DeliveryD28 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD28Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD28") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD29] =  " & ls_DeliveryD29 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD29Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD29") & " " & vbCrLf

                    ls_SQL = ls_SQL + " 			  ,[DeliveryD30] =  " & ls_DeliveryD30 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD30Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD30") & " " & vbCrLf & _
                                      " 			  ,[DeliveryD31] =  " & ls_DeliveryD31 & " " & vbCrLf & _
                                      " 			  --,[DeliveryD31Old] = " & e.UpdateValues(iLoop).OldValues("DeliveryD31") & "  " & vbCrLf & _
                                      " 			  ,[UpdateDate] = getdate() " & vbCrLf & _
                                      " 			  ,[UpdateUser] = '" & Session("UserID") & "' " & vbCrLf & _
                                      " 			WHERE [PONo] = '" & Trim(cboPONo.Text) & "' " & vbCrLf & _
                                      " 			  AND [AffiliateID] ='" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf & _
                                      " 			  AND [SupplierID] = '" & Trim(txtSupplierCode.Text) & "'" & vbCrLf

                    ls_SQL = ls_SQL + " 			  AND [PartNo] = '" & ls_PartNo & "' " & vbCrLf & _
                                      " 		 END  "



                    ls_MsgID = "1002"

                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                Next iLoop


                sqlTran.Commit()
                Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
                If lblInfo.Text = "[] " Then lblInfo.Text = ""
                grid.JSProperties("cpMessage") = lblInfo.Text
                Session("YA010IsSubmit") = lblInfo.Text
            End Using

            sqlConn.Close()
        End Using
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")
            Dim ls_MsgID As String
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Dim pDate As Date = Split(e.Parameters, "|")(1)
            Dim pAffCode As String = Split(e.Parameters, "|")(2)
            Dim pPONo As String = Split(e.Parameters, "|")(3)
            Dim pSuppCode As String = Split(e.Parameters, "|")(4)
            Dim pComm As String = Split(e.Parameters, "|")(5)
            Dim pDelBy As String = Split(e.Parameters, "|")(6)
            Dim pKanban As String = Split(e.Parameters, "|")(7)
            Dim pShipBy As String = Split(e.Parameters, "|")(8)
            Select Case pAction
                Case "load"
                    pSearch = True
                    Call bindDataHeader(pAffCode, pPONo, pSuppCode)
                    Call bindDataDetail(pDate, pAffCode, pPONo, pSuppCode)
                    'pSearch = False
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Session("YA010IsSubmit") = lblInfo.Text
                    End If
                Case "kosong"
                    'Call up_GridLoadWhenEventChange()
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Call SaveDataMaster(pAffCode, pPONo, Trim(pSuppCode))
                    Call SaveDataDetail(pAffCode, pPONo, Trim(pSuppCode))
                    Call SaveDeliveryCls()
                    Call bindDataHeader(pAffCode, pPONo, pSuppCode)
                    Call bindDataDetail(pDate, pAffCode, pPONo, pSuppCode)
                    ls_MsgID = "1001"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Session("YA010IsSubmit") = lblInfo.Text
                Case "send"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    UpdateSend = True
                    Call SaveDeliveryCls()
                    Call UpdatePO(pAffCode, pPONo, pSuppCode)
                    Call bindDataHeader(pAffCode, pPONo, pSuppCode)
                    Call bindDataDetail(pDate, pAffCode, pPONo, pSuppCode)
                    Call UpdateExcel(pAffCode, pPONo, pSuppCode)
                    'Call Excel()
                    UpdateSend = False
            End Select
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        Dim x As Integer = CInt(e.VisibleIndex.ToString())

        If x > grid.VisibleRowCount Then Exit Sub
        If e.GetValue("BYWHAT") = "BY AFFILIATE" Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
            e.Cell.BackColor = Color.AliceBlue
            If e.DataColumn.FieldName = "MonthlyProductionCapacity" Then
                remain = If(IsDBNull(e.GetValue("MonthlyProductionCapacity")), 0, (e.GetValue("MonthlyProductionCapacity")))
                TotQtyAff = If(IsDBNull(e.GetValue("POQty")), 0, e.GetValue("POQty"))
            End If
        End If

        With grid
            If .VisibleRowCount > 0 Then
                'Dim Remaining As Double = CDbl(IIf(IsDBNull(e.GetValue("MonthlyProductionCapacity")) Or e.GetValue("MonthlyProductionCapacity") = "", 0, e.GetValue("MonthlyProductionCapacity")))
                Dim TotalQty As Double = If(IsDBNull(e.GetValue("POQty")), 0, e.GetValue("POQty"))
                If remain < TotalQty Then
                    If e.DataColumn.FieldName = "MonthlyProductionCapacity" Then
                        e.Cell.BackColor = Color.HotPink
                    End If
                End If
                If e.GetValue("BYWHAT") = "BY AFFILIATE" Then
                    TotQtyPASI = If(IsDBNull(e.GetValue("POQty")), 0, e.GetValue("POQty"))
                    If TotQtyAff <> TotQtyPASI Then
                        If e.DataColumn.FieldName = "POQty" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    ls_DeliveryD1Old = If(IsDBNull(e.GetValue("DeliveryD1")), 0, e.GetValue("DeliveryD1"))
                    ls_DeliveryD2Old = If(IsDBNull(e.GetValue("DeliveryD2")), 0, e.GetValue("DeliveryD2"))
                    ls_DeliveryD3Old = If(IsDBNull(e.GetValue("DeliveryD3")), 0, e.GetValue("DeliveryD3"))
                    ls_DeliveryD4Old = If(IsDBNull(e.GetValue("DeliveryD4")), 0, e.GetValue("DeliveryD4"))
                    ls_DeliveryD5Old = If(IsDBNull(e.GetValue("DeliveryD5")), 0, e.GetValue("DeliveryD5"))
                    ls_DeliveryD6Old = If(IsDBNull(e.GetValue("DeliveryD6")), 0, e.GetValue("DeliveryD6"))
                    ls_DeliveryD7Old = If(IsDBNull(e.GetValue("DeliveryD7")), 0, e.GetValue("DeliveryD7"))
                    ls_DeliveryD8Old = If(IsDBNull(e.GetValue("DeliveryD8")), 0, e.GetValue("DeliveryD8"))
                    ls_DeliveryD9Old = If(IsDBNull(e.GetValue("DeliveryD9")), 0, e.GetValue("DeliveryD9"))
                    ls_DeliveryD10Old = If(IsDBNull(e.GetValue("DeliveryD10")), 0, e.GetValue("DeliveryD10"))
                    ls_DeliveryD11Old = If(IsDBNull(e.GetValue("DeliveryD11")), 0, e.GetValue("DeliveryD11"))
                    ls_DeliveryD12Old = If(IsDBNull(e.GetValue("DeliveryD12")), 0, e.GetValue("DeliveryD12"))
                    ls_DeliveryD13Old = If(IsDBNull(e.GetValue("DeliveryD13")), 0, e.GetValue("DeliveryD13"))
                    ls_DeliveryD14Old = If(IsDBNull(e.GetValue("DeliveryD14")), 0, e.GetValue("DeliveryD14"))
                    ls_DeliveryD15Old = If(IsDBNull(e.GetValue("DeliveryD15")), 0, e.GetValue("DeliveryD15"))
                    ls_DeliveryD16Old = If(IsDBNull(e.GetValue("DeliveryD16")), 0, e.GetValue("DeliveryD16"))
                    ls_DeliveryD17Old = If(IsDBNull(e.GetValue("DeliveryD17")), 0, e.GetValue("DeliveryD17"))
                    ls_DeliveryD18Old = If(IsDBNull(e.GetValue("DeliveryD18")), 0, e.GetValue("DeliveryD18"))
                    ls_DeliveryD19Old = If(IsDBNull(e.GetValue("DeliveryD19")), 0, e.GetValue("DeliveryD19"))
                    ls_DeliveryD20Old = If(IsDBNull(e.GetValue("DeliveryD20")), 0, e.GetValue("DeliveryD20"))
                    ls_DeliveryD21Old = If(IsDBNull(e.GetValue("DeliveryD21")), 0, e.GetValue("DeliveryD21"))
                    ls_DeliveryD22Old = If(IsDBNull(e.GetValue("DeliveryD22")), 0, e.GetValue("DeliveryD22"))
                    ls_DeliveryD23Old = If(IsDBNull(e.GetValue("DeliveryD23")), 0, e.GetValue("DeliveryD23"))
                    ls_DeliveryD24Old = If(IsDBNull(e.GetValue("DeliveryD24")), 0, e.GetValue("DeliveryD24"))
                    ls_DeliveryD25Old = If(IsDBNull(e.GetValue("DeliveryD25")), 0, e.GetValue("DeliveryD25"))
                    ls_DeliveryD26Old = If(IsDBNull(e.GetValue("DeliveryD26")), 0, e.GetValue("DeliveryD26"))
                    ls_DeliveryD27Old = If(IsDBNull(e.GetValue("DeliveryD27")), 0, e.GetValue("DeliveryD27"))
                    ls_DeliveryD28Old = If(IsDBNull(e.GetValue("DeliveryD28")), 0, e.GetValue("DeliveryD28"))
                    ls_DeliveryD29Old = If(IsDBNull(e.GetValue("DeliveryD29")), 0, e.GetValue("DeliveryD29"))
                    ls_DeliveryD30Old = If(IsDBNull(e.GetValue("DeliveryD30")), 0, e.GetValue("DeliveryD30"))
                    ls_DeliveryD31Old = If(IsDBNull(e.GetValue("DeliveryD31")), 0, e.GetValue("DeliveryD31"))
                End If
                If e.GetValue("BYWHAT") = "BY PASI" Then
                    If e.DataColumn.FieldName = "MonthlyProductionCapacity" Or e.DataColumn.FieldName = "ForecastN1" Or e.DataColumn.FieldName = "ForecastN2" Or e.DataColumn.FieldName = "ForecastN3" Then
                        e.Cell.Text = ""
                    End If
                    If e.DataColumn.FieldName = "DeliveryD1" Or e.DataColumn.FieldName = "DeliveryD2" Or e.DataColumn.FieldName = "DeliveryD3" Or e.DataColumn.FieldName = "DeliveryD4" Or e.DataColumn.FieldName = "DeliveryD5" _
                        Or e.DataColumn.FieldName = "DeliveryD6" Or e.DataColumn.FieldName = "DeliveryD7" Or e.DataColumn.FieldName = "DeliveryD8" Or e.DataColumn.FieldName = "DeliveryD9" Or e.DataColumn.FieldName = "DeliveryD10" Or e.DataColumn.FieldName = "DeliveryD11" _
                        Or e.DataColumn.FieldName = "DeliveryD12" Or e.DataColumn.FieldName = "DeliveryD13" Or e.DataColumn.FieldName = "DeliveryD14" Or e.DataColumn.FieldName = "DeliveryD15" Or e.DataColumn.FieldName = "DeliveryD16" Or e.DataColumn.FieldName = "DeliveryD17" _
                        Or e.DataColumn.FieldName = "DeliveryD18" Or e.DataColumn.FieldName = "DeliveryD19" Or e.DataColumn.FieldName = "DeliveryD20" Or e.DataColumn.FieldName = "DeliveryD21" Or e.DataColumn.FieldName = "DeliveryD22" Or e.DataColumn.FieldName = "DeliveryD23" _
                        Or e.DataColumn.FieldName = "DeliveryD24" Or e.DataColumn.FieldName = "DeliveryD25" Or e.DataColumn.FieldName = "DeliveryD26" Or e.DataColumn.FieldName = "DeliveryD27" Or e.DataColumn.FieldName = "DeliveryD28" Or e.DataColumn.FieldName = "DeliveryD29" _
                        Or e.DataColumn.FieldName = "DeliveryD30" Or e.DataColumn.FieldName = "DeliveryD31" Then
                        e.Cell.BackColor = Color.White
                    End If

                    If If(IsDBNull(e.GetValue("POQty")), 0, e.GetValue("POQty")) <> TotQtyPASI Then
                        If e.DataColumn.FieldName = "POQty" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD1")), 0, e.GetValue("DeliveryD1")) <> ls_DeliveryD1Old Then
                        If e.DataColumn.FieldName = "DeliveryD1" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD2")), 0, e.GetValue("DeliveryD2")) <> ls_DeliveryD2Old Then
                        If e.DataColumn.FieldName = "DeliveryD2" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD3")), 0, e.GetValue("DeliveryD3")) <> ls_DeliveryD3Old Then
                        If e.DataColumn.FieldName = "DeliveryD3" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD4")), 0, e.GetValue("DeliveryD4")) <> ls_DeliveryD4Old Then
                        If e.DataColumn.FieldName = "DeliveryD4" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD5")), 0, e.GetValue("DeliveryD5")) <> ls_DeliveryD5Old Then
                        If e.DataColumn.FieldName = "DeliveryD5" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD6")), 0, e.GetValue("DeliveryD6")) <> ls_DeliveryD6Old Then
                        If e.DataColumn.FieldName = "DeliveryD6" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD7")), 0, e.GetValue("DeliveryD7")) <> ls_DeliveryD7Old Then
                        If e.DataColumn.FieldName = "DeliveryD7" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD8")), 0, e.GetValue("DeliveryD8")) <> ls_DeliveryD8Old Then
                        If e.DataColumn.FieldName = "DeliveryD8" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD9")), 0, e.GetValue("DeliveryD9")) <> ls_DeliveryD9Old Then
                        If e.DataColumn.FieldName = "DeliveryD9" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD10")), 0, e.GetValue("DeliveryD10")) <> ls_DeliveryD10Old Then
                        If e.DataColumn.FieldName = "DeliveryD10" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD11")), 0, e.GetValue("DeliveryD11")) <> ls_DeliveryD11Old Then
                        If e.DataColumn.FieldName = "DeliveryD11" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD12")), 0, e.GetValue("DeliveryD12")) <> ls_DeliveryD12Old Then
                        If e.DataColumn.FieldName = "DeliveryD12" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD13")), 0, e.GetValue("DeliveryD13")) <> ls_DeliveryD13Old Then
                        If e.DataColumn.FieldName = "DeliveryD13" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD14")), 0, e.GetValue("DeliveryD14")) <> ls_DeliveryD14Old Then
                        If e.DataColumn.FieldName = "DeliveryD14" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD15")), 0, e.GetValue("DeliveryD15")) <> ls_DeliveryD15Old Then
                        If e.DataColumn.FieldName = "DeliveryD15" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD16")), 0, e.GetValue("DeliveryD16")) <> ls_DeliveryD16Old Then
                        If e.DataColumn.FieldName = "DeliveryD16" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD17")), 0, e.GetValue("DeliveryD17")) <> ls_DeliveryD17Old Then
                        If e.DataColumn.FieldName = "DeliveryD17" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD18")), 0, e.GetValue("DeliveryD18")) <> ls_DeliveryD18Old Then
                        If e.DataColumn.FieldName = "DeliveryD18" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD19")), 0, e.GetValue("DeliveryD19")) <> ls_DeliveryD19Old Then
                        If e.DataColumn.FieldName = "DeliveryD19" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD20")), 0, e.GetValue("DeliveryD20")) <> ls_DeliveryD20Old Then
                        If e.DataColumn.FieldName = "DeliveryD20" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD21")), 0, e.GetValue("DeliveryD21")) <> ls_DeliveryD21Old Then
                        If e.DataColumn.FieldName = "DeliveryD21" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD22")), 0, e.GetValue("DeliveryD22")) <> ls_DeliveryD22Old Then
                        If e.DataColumn.FieldName = "DeliveryD22" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD23")), 0, e.GetValue("DeliveryD23")) <> ls_DeliveryD23Old Then
                        If e.DataColumn.FieldName = "DeliveryD23" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD24")), 0, e.GetValue("DeliveryD24")) <> ls_DeliveryD24Old Then
                        If e.DataColumn.FieldName = "DeliveryD24" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD25")), 0, e.GetValue("DeliveryD25")) <> ls_DeliveryD25Old Then
                        If e.DataColumn.FieldName = "DeliveryD25" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD26")), 0, e.GetValue("DeliveryD26")) <> ls_DeliveryD26Old Then
                        If e.DataColumn.FieldName = "DeliveryD26" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD27")), 0, e.GetValue("DeliveryD27")) <> ls_DeliveryD27Old Then
                        If e.DataColumn.FieldName = "DeliveryD27" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD28")), 0, e.GetValue("DeliveryD28")) <> ls_DeliveryD28Old Then
                        If e.DataColumn.FieldName = "DeliveryD28" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD29")), 0, e.GetValue("DeliveryD29")) <> ls_DeliveryD29Old Then
                        If e.DataColumn.FieldName = "DeliveryD29" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD30")), 0, e.GetValue("DeliveryD30")) <> ls_DeliveryD30Old Then
                        If e.DataColumn.FieldName = "DeliveryD30" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If If(IsDBNull(e.GetValue("DeliveryD31")), 0, e.GetValue("DeliveryD31")) <> ls_DeliveryD31Old Then
                        If e.DataColumn.FieldName = "DeliveryD31" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                End If
            End If
        End With
    End Sub

    Protected Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubMenu.Click
        If Session("M01Url") <> "" Then
            'Session.Remove("M01Url")
            Response.Redirect("~/AffiliateOrder/AffiliateOrderList.aspx")
        Else
            'Session.Remove("M01Url")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Private Sub cboPONo_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboPONo.Callback
        If String.IsNullOrEmpty(e.Parameter) Then
            Return
        End If

        Dim ls_value As String = Split(e.Parameter, "|")(0)
        Dim ls_sql As String = ""

        ls_sql = "SELECT '" & clsGlobal.gs_All & "' PONo UNION ALL SELECT RTRIM(PONo)PONo FROM dbo.PO_Master WHERE AffiliateID='" & ls_value & "' " & vbCrLf
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
                .Columns(0).Width = 50

                .TextField = "PONo"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub cbPONo_Callback(source As Object, e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbPONo.Callback

        Dim ls_sql As String = ""

        Dim pAction As String = Split(e.Parameter, "|")(0)
        Dim pDate As Date = Split(e.Parameter, "|")(1)
        Dim pPONo As String = Split(e.Parameter, "|")(2)
        Dim pAffCode As String = Split(e.Parameter, "|")(3)


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_sql = "   SELECT DISTINCT Period,POD.AffiliateID,AffiliateName,POM.PONo  " & vbCrLf & _
                  "   ,CASE WHEN CommercialCls = '0' THEN 'NO' ELSE 'YES' END CommercialCls  " & vbCrLf & _
                  "   ,POD.SupplierID,SupplierName,ShipCls   " & vbCrLf & _
                  "   ,PODeliveryBy   " & vbCrLf & _
                  "   ,POD.KanbanCls   " & vbCrLf & _
                  "   FROM dbo.PO_Master POM    " & vbCrLf & _
                  "   LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID  AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                  "   LEFT JOIN dbo.MS_Affiliate MA ON POD.AffiliateID = MA.AffiliateID  " & vbCrLf & _
                  "   LEFT JOIN dbo.MS_Parts MP ON POD.PartNo = MP.PartNo   " & vbCrLf & _
                  "   LEFT JOIN dbo.MS_Supplier MS ON POD.SupplierID = MS.SupplierID   " & vbCrLf & _
                  "  WHERE YEAR(Period) = YEAR('" & pDate & "') AND MONTH(Period) = MONTH('" & pDate & "')  "

            ls_sql = ls_sql + "  AND POM.PONo = '" & Trim(pPONo) & "'    " & vbCrLf & _
                              "  AND POM.AffiliateID='" & pAffCode & "' "


            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With ds.Tables(0)
                If ds.Tables(0).Rows.Count > 0 Then
                    cbPONo.JSProperties("cpCommercialCls") = .Rows(0).Item("CommercialCls")
                    cbPONo.JSProperties("cpSupplierID") = .Rows(0).Item("SupplierID")
                    cbPONo.JSProperties("cpSupplierName") = .Rows(0).Item("SupplierName")
                    cbPONo.JSProperties("cpShipCls") = .Rows(0).Item("ShipCls")
                    cbPONo.JSProperties("cpPODeliveryBy") = .Rows(0).Item("PODeliveryBy")
                    cbPONo.JSProperties("cpKanbanCls") = .Rows(0).Item("KanbanCls")
                Else
                    cbPONo.JSProperties("cpCommercialCls") = ""
                    cbPONo.JSProperties("cpSupplierID") = ""
                    cbPONo.JSProperties("cpSupplierName") = ""
                    cbPONo.JSProperties("cpShipCls") = ""
                    cbPONo.JSProperties("cpPODeliveryBy") = 0
                    cbPONo.JSProperties("cpKanbanCls") = 2
                End If
            End With

            sqlConn.Close()
        End Using
    End Sub

#End Region

#Region "PROCEDURE"
    Private Sub up_Fillcombo()
        Dim ls_SQL As String = ""
        'Combo Affiliate
        ls_SQL = "SELECT RTRIM(AffiliateID) AffiliateID,AffiliateName FROM dbo.MS_Affiliate" & vbCrLf
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
        'Combo Affiliate
        ls_SQL = "SELECT RTRIM(PONo)PONo FROM dbo.PO_Master" & vbCrLf
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
                .Columns(0).Width = 50

                .TextField = "PONo"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataHeader(ByVal pAffCode As String, ByVal pPONo As String, ByVal pSupplierID As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT DISTINCT " & vbCrLf & _
                      " 	PM.Period " & vbCrLf & _
                      " 	, PM.AffiliateID " & vbCrLf & _
                      " 	, MA.AffiliateName " & vbCrLf & _
                      " 	, PM.SupplierID " & vbCrLf & _
                      " 	, MS.SupplierName " & vbCrLf & _
                      " 	, PM.PONo " & vbCrLf & _
                      " 	, CASE WHEN PM.CommercialCls = '0' THEN 'NO' ELSE 'YES' END CommercialCls " & vbCrLf & _
                      " 	, PM.ShipCls " & vbCrLf & _
                      " 	, ISNULL(PM.DeliveryByPASICls,0)PODeliveryBy  " & vbCrLf & _
                      " 	, ISNULL(PD.KanbanCls,0)KanbanCls "

            ls_SQL = ls_SQL + " 	,CONVERT(DATETIME,PM.EntryDate,120)EntryDate,ISNULL(PM.EntryUser,'')EntryUser --1   " & vbCrLf & _
                              " 	,CONVERT(DATETIME,PM.AffiliateApproveDate,120)AffiliateApproveDate,ISNULL(PM.AffiliateApproveUser,'')AffiliateApproveUser --2   " & vbCrLf & _
                              " 	,CONVERT(DATETIME,PM.PASISendAffiliateDate,120)PASISendAffiliateDate,ISNULL(PM.PASISendAffiliateUser,'')PASISendAffiliateUser --3   " & vbCrLf & _
                              " 	,CONVERT(DATETIME,PM.SupplierApproveDate,120)SupplierApproveDate,ISNULL(PM.SupplierApproveUser,'')SupplierApproveUser --4   " & vbCrLf & _
                              " 	,CONVERT(DATETIME,PM.SupplierApprovePendingDate,120)SupplierApprovePendingDate,ISNULL(PM.SupplierApprovePendingUser,'')SupplierApprovePendingUser --5   " & vbCrLf & _
                              " 	,CONVERT(DATETIME,PM.SupplierUnApproveDate,120)SupplierUnApproveDate,ISNULL(PM.SupplierUnApproveUser,'')SupplierUnApproveUser --6   " & vbCrLf & _
                              " 	,CONVERT(DATETIME,PM.PASIApproveDate,120)PASIApproveDate,ISNULL(PM.PASIApproveUser ,'')PASIApproveUser --7   " & vbCrLf & _
                              " 	,CONVERT(DATETIME,PM.FinalApproveDate,120)FinalApproveDate,ISNULL(PM.FinalApproveUser,'')FinalApproveUser --8   " & vbCrLf & _
                              " FROM PO_Master PM " & vbCrLf & _
                              " INNER JOIN PO_Detail PD ON PM.AffiliateID = PD.AffiliateID and PM.SupplierID = PD.SupplierID and PM.PONo = PD.PONo " & vbCrLf & _
                              " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PM.AffiliateID "

            ls_SQL = ls_SQL + " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PM.SupplierID " & vbCrLf & _
                              " WHERE PM.AffiliateID = '" & pAffCode & "' AND PM.SupplierID = '" & pSupplierID & "' AND PM.PONo = '" & pPONo & "' " & vbCrLf & _
                              "  " & vbCrLf & _
                              "  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                dtPeriod.Value = ds.Tables(0).Rows(0)("Period")
                cboAffiliateCode.Text = ds.Tables(0).Rows(0)("AffiliateID")
                txtAffiliateName.Text = ds.Tables(0).Rows(0)("AffiliateName")
                cboPONo.Text = ds.Tables(0).Rows(0)("PONo")
                txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                txtSupplierCode.Text = ds.Tables(0).Rows(0)("SupplierID")
                txtSupplierName.Text = ds.Tables(0).Rows(0)("SupplierName")
                txtShipBy.Text = ds.Tables(0).Rows(0)("ShipCls")
                rblDelivery.Value = ds.Tables(0).Rows(0)("PODeliveryBy")
                rblPOKanban.Value = ds.Tables(0).Rows(0)("KanbanCls")
                txtEntryDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("EntryDate")), "", Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd hh:mm:ss"))
                txtEntryUser.Text = ds.Tables(0).Rows(0)("EntryUser")
                txtAffAppDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")), "", Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                txtAffAppUser.Text = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                txtSendDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")), "", Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd hh:mm:ss"))
                txtSendUser.Text = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                txtSuppAppDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                txtSuppAppUser.Text = ds.Tables(0).Rows(0)("SupplierApproveUser")
                txtSuppPendDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd hh:mm:ss"))
                txtSuppPendUser.Text = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                txtSuppUnpDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                txtSuppUnpUser.Text = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                txtPASIAppDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")), "", Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                txtPASIAppUser.Text = ds.Tables(0).Rows(0)("PASIApproveUser")
                txtAffFinalAppDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")), "", Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                txtAffFinalAppUser.Text = ds.Tables(0).Rows(0)("FinalApproveUser")

                grid.JSProperties("cpKanban") = ds.Tables(0).Rows(0)("KanbanCls")
                grid.JSProperties("cpEntryDate") = If(IsDBNull(ds.Tables(0).Rows(0)("EntryDate")), "", Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpEntryUser") = ds.Tables(0).Rows(0)("EntryUser")
                grid.JSProperties("cpAffAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")), "", Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpAffAppUser") = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                grid.JSProperties("cpSendDate") = If(IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")), "", Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpSendUser") = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                grid.JSProperties("cpSuppAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpSuppAppUser") = ds.Tables(0).Rows(0)("SupplierApproveUser")
                grid.JSProperties("cpSuppAppPendingDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpSuppAppPendingUser") = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                grid.JSProperties("cpSuppUnApproveDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpSuppUnApproveUser") = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                grid.JSProperties("cpPASIAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")), "", Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpPASIAppUser") = ds.Tables(0).Rows(0)("PASIApproveUser")
                grid.JSProperties("cpFinalAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")), "", Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpFinalAppUser") = ds.Tables(0).Rows(0)("FinalApproveUser")
                If UpdateSend = True Then
                    Call clsMsg.DisplayMessage(lblInfo, "1008", clsMessage.MsgType.InformationMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Session("YA010IsSubmit") = lblInfo.Text
                End If
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataDetail(ByVal pDate As Date, ByVal pAffCode As String, ByVal pPONo As String, ByVal pSupplierCode As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.Affiliate_Detail WHERE PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSupplierCode) & "')  " & vbCrLf & _
                  "    BEGIN   " & vbCrLf & _
                  " 	   SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName, KanbanCls = KanbanCls ,Description    " & vbCrLf & _
                  " 			,MOQ , MinOrderQty, QtyBox , Maker   " & vbCrLf & _
                  " 			,ISNULL(MonthlyProductionCapacity,0)MonthlyProductionCapacity " & vbCrLf & _
                  " 			,BYWHAT    " & vbCrLf & _
                  " 			,POQty " & vbCrLf & _
                  " 			,ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                  " 			,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                  " 			,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                  " 			,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15   "

            ls_SQL = ls_SQL + " 			,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              " 		FROM (    " & vbCrLf & _
                              "  			SELECT  " & vbCrLf & _
                              " 				CONVERT(CHAR,row_number() over (order by POD.PONo)) as NoUrut,POD.PartNo,POD.PartNo PartNos,PartName " & vbCrLf & _
                              " 				,CASE WHEN POD.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls " & vbCrLf & _
                              "      			,MU.Description,MOQ =CONVERT(CHAR,MOQ),MinOrderQty = MOQ, QtyBox = CONVERT(CHAR,QtyBox),ISNULL(MPART.Maker,'')Maker " & vbCrLf & _
                              " 				,MonthlyProductionCapacity = (SELECT ISNULL(QtyRemaining, MonthlyInjectionCapacity) from MS_SupplierCapacity A  " & vbCrLf & _
                              " 												LEFT JOIN RemainingCapacity B ON A.PartNo = B.PartNo AND A.SupplierID = B.SupplierID AND POD.PartNo = B.PartNo  AND POD.SupplierID = A.SupplierID " & vbCrLf & _
                              " 												WHERE B.Period = '" & Format(pDate, "yyyMM") & "') "

            ls_SQL = ls_SQL + " 				,'BY AFFILIATE' BYWHAT      " & vbCrLf & _
                              "      			,SUM(POQty)POQty  " & vbCrLf & _
                              "   				,ISNULL(ForecastN1,0)ForecastN1  " & vbCrLf & _
                              "    				,ISNULL(ForecastN2,0)ForecastN2  " & vbCrLf & _
                              "    				,ISNULL(ForecastN3,0)ForecastN3  " & vbCrLf & _
                              "      			,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10  " & vbCrLf & _
                              "      			,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20  " & vbCrLf & _
                              "      			,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31  " & vbCrLf & _
                              "  				,row_number() over (order by POD.PONo) as Sort  " & vbCrLf & _
                              " 			FROM dbo.PO_Master POM  " & vbCrLf & _
                              " 			LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID and POM.SupplierID = POD.SupplierID and POM.PONo = POD.PONo "

            ls_SQL = ls_SQL + " 			LEFT JOIN dbo.MS_PartMapping  MPP ON MPP.AffiliateID = POD.AffiliateID and MPP.SupplierID = POD.SupplierID and MPP.PartNo = POD.PartNo " & vbCrLf & _
                              " 			LEFT JOIN dbo.MS_Parts MPART ON POD.PartNo = MPART.PartNo " & vbCrLf & _
                              " 			LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls                  " & vbCrLf & _
                              " 			WHERE POM.PONo = '" & Trim(pPONo) & "'  AND POM.SupplierID='" & Trim(pSupplierCode) & "' AND POM.AffiliateID='" & Trim(pAffCode) & "'  " & vbCrLf & _
                              " 			GROUP BY POD.PONo,POD.PartNo,PartName,POD.KanbanCls,MU.Description,MOQ,QtyBox,MPART.Maker, POD.SupplierID " & vbCrLf & _
                              " 				,Period,POD.PartNo,POM.AffiliateID, POD.ForecastN1, POD.ForecastN2, POD.ForecastN3    " & vbCrLf & _
                              " 				,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10          " & vbCrLf & _
                              " 				,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20        		    " & vbCrLf & _
                              " 				,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31     " & vbCrLf & _
                              "  		) Detail1    " & vbCrLf & _
                              " 		UNION ALL    "

            ls_SQL = ls_SQL + " 		SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              " 			,MOQ = MOQ ,MinOrderQty, QtyBox = QtyBox ,Maker ,MonthlyProductionCapacity ,BYWHAT    " & vbCrLf & _
                              " 			,POQty  " & vbCrLf & _
                              " 			,ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                              " 			,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              " 		FROM (  "

            ls_SQL = ls_SQL + " 			SELECT '' as NoUrut,'' PartNo,POD.PartNo PartNos,''PartName,'' KanbanCls,''Description,MOQ = '',MinOrderQty = MOQ " & vbCrLf & _
                              "  				,'' QtyBox,ISNULL(MPART.Maker,'')Maker   " & vbCrLf & _
                              "      			,0 MonthlyProductionCapacity,'BY PASI' BYWHAT,SUM(POQty)POQty  " & vbCrLf & _
                              "   				,ISNULL(ForecastN1,0)ForecastN1  " & vbCrLf & _
                              "    				,ISNULL(ForecastN2,0)ForecastN2  " & vbCrLf & _
                              "    				,ISNULL(ForecastN3,0)ForecastN3  " & vbCrLf & _
                              "      			,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10  " & vbCrLf & _
                              "      			,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20  " & vbCrLf & _
                              "      			,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31  " & vbCrLf & _
                              "  				,row_number() over (order by POD.PONo) as Sort      " & vbCrLf & _
                              "      		FROM dbo.PO_Master POM  "

            ls_SQL = ls_SQL + "      		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID and POM.SupplierID = POD.SupplierID and POM.PONo = POD.PONo  " & vbCrLf & _
                              " 			LEFT JOIN dbo.MS_PartMapping  MPP ON MPP.AffiliateID = POD.AffiliateID and MPP.SupplierID = POD.SupplierID and MPP.PartNo = POD.PartNo     " & vbCrLf & _
                              "  			LEFT JOIN dbo.MS_Parts MPART ON POD.PartNo = MPART.PartNo              " & vbCrLf & _
                              "  			LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls     " & vbCrLf & _
                              "         WHERE POM.PONo = '" & Trim(pPONo) & "'  AND POM.SupplierID='" & Trim(pSupplierCode) & "' AND POM.AffiliateID='" & Trim(pAffCode) & "'  " & vbCrLf & _
                              "  				GROUP BY POD.PONo,POD.PartNo,PartName,POD.KanbanCls,MU.Description,MOQ,QtyBox,MPART.Maker,POD.SupplierID        " & vbCrLf & _
                              "      			,Period,POD.PartNo,POM.AffiliateID, POD.ForecastN1, POD.ForecastN2, POD.ForecastN3   " & vbCrLf & _
                              "      			,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10        " & vbCrLf & _
                              "      			,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20          " & vbCrLf & _
                              "      			,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31   " & vbCrLf & _
                              "  		)detail2    "

            ls_SQL = ls_SQL + "  	ORDER BY sort, PartNos, NoUrut DESC    " & vbCrLf & _
                              " 	END  " & vbCrLf & _
                              " ELSE  " & vbCrLf & _
                              " 	BEGIN  " & vbCrLf & _
                              " 		SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              " 			,MOQ, MinOrderQty, QtyBox , Maker , ISNULL(MonthlyProductionCapacity,0) MonthlyProductionCapacity, BYWHAT    " & vbCrLf & _
                              " 			,POQty , ForecastN1 ,ForecastN2 ,ForecastN3   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   "

            ls_SQL = ls_SQL + " 			,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              " 			,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              " 		FROM (    " & vbCrLf & _
                              " 			SELECT  " & vbCrLf & _
                              " 				row_number() over (order by AD.PONo) as Sort, CONVERT(CHAR,row_number() over (order by AD.PONo)) as NoUrut,  " & vbCrLf & _
                              " 				AD.PartNo as PartNo, AD.PartNo AS PartNos, PartName, CASE WHEN PD.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls, MU.Description,  " & vbCrLf & _
                              " 				MOQ = CONVERT(CHAR,MOQ), MinOrderQty = MOQ, QtyBox = CONVERT(CHAR,QtyBox), MPART.Maker,   " & vbCrLf & _
                              " 				MonthlyProductionCapacity = (SELECT ISNULL(QtyRemaining, MonthlyInjectionCapacity) from MS_SupplierCapacity A  " & vbCrLf & _
                              " 											 LEFT JOIN RemainingCapacity B ON A.PartNo = B.PartNo AND A.SupplierID = B.SupplierID AND AD.PartNo = B.PartNo AND AD.SupplierID = B.SupplierID " & vbCrLf & _
                              " 											 WHERE B.Period = '" & Format(pDate, "yyyyMM") & "')   " & vbCrLf & _
                              " 				,'BY AFFILIATE' BYWHAT    "

            ls_SQL = ls_SQL + "   				,POQtyOld POqty  " & vbCrLf & _
                              "   				,ISNULL(ForecastN1,0)ForecastN1  " & vbCrLf & _
                              "    				,ISNULL(ForecastN2,0)ForecastN2  " & vbCrLf & _
                              "    				,ISNULL(ForecastN3,0)ForecastN3  " & vbCrLf & _
                              "      			,DeliveryD1Old DeliveryD1,DeliveryD2Old DeliveryD2,DeliveryD3Old DeliveryD3,DeliveryD4Old DeliveryD4,DeliveryD5Old DeliveryD5   " & vbCrLf & _
                              "   				,DeliveryD6Old DeliveryD6,DeliveryD7Old DeliveryD7,DeliveryD8Old DeliveryD8,DeliveryD9Old DeliveryD9,DeliveryD10Old DeliveryD10   " & vbCrLf & _
                              "   				,DeliveryD11Old DeliveryD11,DeliveryD12Old DeliveryD12,DeliveryD13Old DeliveryD13,DeliveryD14Old DeliveryD14,DeliveryD15Old DeliveryD15   " & vbCrLf & _
                              "   				,DeliveryD16Old DeliveryD16,DeliveryD17Old DeliveryD17,DeliveryD18Old DeliveryD18,DeliveryD19Old DeliveryD19,DeliveryD20Old DeliveryD20   " & vbCrLf & _
                              "   				,DeliveryD21Old DeliveryD21,DeliveryD22Old DeliveryD22,DeliveryD23Old DeliveryD23,DeliveryD24Old DeliveryD24,DeliveryD25Old DeliveryD25   " & vbCrLf & _
                              "   				,DeliveryD26Old DeliveryD26,DeliveryD27Old DeliveryD27,DeliveryD28Old DeliveryD28,DeliveryD29Old DeliveryD29,DeliveryD30Old DeliveryD30 " & vbCrLf & _
                              " 				,DeliveryD31Old DeliveryD31   "

            ls_SQL = ls_SQL + "  			FROM dbo.Affiliate_Detail AD  " & vbCrLf & _
                              "  			LEFT JOIN dbo.PO_Detail PD ON AD.PONo = PD.PONo and AD.SupplierID = PD.SupplierID and AD.AffiliateID = PD.AffiliateID and AD.PartNo = PD.PartNo " & vbCrLf & _
                              " 			LEFT JOIN dbo.MS_PartMapping  MPP ON MPP.AffiliateID = AD.AffiliateID and MPP.SupplierID = AD.SupplierID and MPP.PartNo = AD.PartNo " & vbCrLf & _
                              "  			LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo       " & vbCrLf & _
                              "  			LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls         " & vbCrLf & _
                              "  			WHERE AD.PONo='" & Trim(pPONo) & "' AND AD.SupplierID='" & Trim(pSupplierCode) & "' AND AD.AffiliateID='" & Trim(pAffCode) & "'  " & vbCrLf & _
                              "   			GROUP BY AD.PONo,AD.PartNo,PartName,PD.KanbanCls,POQtyOld,MU.Description,MOQ,QtyBox,MPART.Maker,AD.SupplierID " & vbCrLf & _
                              "      				,AD.PartNo,AD.AffiliateID, PD.ForecastN1, PD.ForecastN2, PD.ForecastN3    " & vbCrLf & _
                              "      				,DeliveryD1Old,DeliveryD2Old,DeliveryD3Old,DeliveryD4Old,DeliveryD5Old   " & vbCrLf & _
                              "   					,DeliveryD6Old,DeliveryD7Old,DeliveryD8Old,DeliveryD9Old,DeliveryD10Old   " & vbCrLf & _
                              "   					,DeliveryD11Old,DeliveryD12Old,DeliveryD13Old,DeliveryD14Old,DeliveryD15Old   "

            ls_SQL = ls_SQL + "   					,DeliveryD16Old,DeliveryD17Old,DeliveryD18Old,DeliveryD19Old,DeliveryD20Old   " & vbCrLf & _
                              "   					,DeliveryD21Old,DeliveryD22Old,DeliveryD23Old,DeliveryD24Old,DeliveryD25Old   " & vbCrLf & _
                              "   					,DeliveryD26Old,DeliveryD27Old,DeliveryD28Old,DeliveryD29Old,DeliveryD30Old,DeliveryD31Old " & vbCrLf & _
                              " 		)detail1 " & vbCrLf & _
                              " 		UNION ALL  " & vbCrLf

            ls_SQL = ls_SQL + "         SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              "  			,MOQ = MOQ ,MinOrderQty, QtyBox = QtyBox ,Maker ,MonthlyProductionCapacity ,BYWHAT    " & vbCrLf & _
                              "  			,POqty  " & vbCrLf & _
                              "  			,ForecastN1 ,ForecastN2 ,ForecastN3   " & vbCrLf & _
                              "  			,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              "  			,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   			,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15   " & vbCrLf & _
                              "  			,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              "  			,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              "  			,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31  " & vbCrLf & _
                              "  		FROM (     " & vbCrLf & _
                              "  			SELECT   "

            ls_SQL = ls_SQL + "  				row_number() over (order by AD.PONo) as Sort,'' as NoUrut,'' PartNo,AD.PartNo PartNos,''PartName,'' KanbanCls,''Description,MOQ = ''      " & vbCrLf & _
                              "  				,MinOrderQty = MOQ, '' QtyBox,ISNULL(MPART.Maker,'')Maker , 0 MonthlyProductionCapacity,'BY PASI' BYWHAT   " & vbCrLf & _
                              "  				,AD.POQty  " & vbCrLf & _
                              "  				,ISNULL(ForecastN1,0)ForecastN1   " & vbCrLf & _
                              "  				,ISNULL(ForecastN2,0)ForecastN2   				,ISNULL(ForecastN3,0)ForecastN3   " & vbCrLf & _
                              "  				,AD.DeliveryD1 ,AD.DeliveryD2 ,AD.DeliveryD3 ,AD.DeliveryD4 ,AD.DeliveryD5 ,AD.DeliveryD6 ,AD.DeliveryD7 ,AD.DeliveryD8 ,AD.DeliveryD9 ,AD.DeliveryD10     " & vbCrLf & _
                              "  				,AD.DeliveryD11 ,AD.DeliveryD12 ,AD.DeliveryD13 ,AD.DeliveryD14 ,AD.DeliveryD15 ,AD.DeliveryD16 ,AD.DeliveryD17 ,AD.DeliveryD18 ,AD.DeliveryD19 ,AD.DeliveryD20    " & vbCrLf & _
                              "  				,AD.DeliveryD21 ,AD.DeliveryD22 ,AD.DeliveryD23 ,AD.DeliveryD24 ,AD.DeliveryD25 ,AD.DeliveryD26 ,AD.DeliveryD27 ,AD.DeliveryD28 ,AD.DeliveryD29 ,AD.DeliveryD30 ,AD.DeliveryD31   " & vbCrLf & _
                              "  			FROM dbo.Affiliate_Detail AD   " & vbCrLf & _
                              "  			LEFT JOIN dbo.PO_Detail PD ON AD.PONo = PD.PONo and AD.SupplierID = PD.SupplierID and AD.AffiliateID = PD.AffiliateID and AD.PartNo = PD.PartNo   " & vbCrLf & _
                              "  			LEFT JOIN dbo.MS_PartMapping  MPP ON MPP.AffiliateID = AD.AffiliateID and MPP.SupplierID = AD.SupplierID and MPP.PartNo = AD.PartNo  "

            ls_SQL = ls_SQL + "  			LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo  					  " & vbCrLf & _
                              "  			LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls    " & vbCrLf & _
                              "  			WHERE  AD.PONo='" & Trim(pPONo) & "' AND AD.SupplierID='" & Trim(pSupplierCode) & "' AND AD.AffiliateID='" & Trim(pAffCode) & "'  " & vbCrLf & _
                              "  			GROUP BY AD.PONo,AD.PartNo,PartName,PD.KanbanCls,AD.POQty,MU.Description,MOQ,QtyBox,MPART.Maker,AD.PartNo,AD.AffiliateID, PD.ForecastN1, PD.ForecastN2, PD.ForecastN3     " & vbCrLf & _
                              "  					,AD.DeliveryD1,AD.DeliveryD2,AD.DeliveryD3,AD.DeliveryD4,AD.DeliveryD5    " & vbCrLf & _
                              "  					,AD.DeliveryD6,AD.DeliveryD7,AD.DeliveryD8,AD.DeliveryD9,AD.DeliveryD10    " & vbCrLf & _
                              "  					,AD.DeliveryD11,AD.DeliveryD12,AD.DeliveryD13,AD.DeliveryD14,AD.DeliveryD15    " & vbCrLf & _
                              "  					,AD.DeliveryD16,AD.DeliveryD17,AD.DeliveryD18,AD.DeliveryD19,AD.DeliveryD20    " & vbCrLf & _
                              "  					,AD.DeliveryD21,AD.DeliveryD22,AD.DeliveryD23,AD.DeliveryD24,AD.DeliveryD25    " & vbCrLf & _
                              "  					,AD.DeliveryD26,AD.DeliveryD27,AD.DeliveryD28,AD.DeliveryD29,AD.DeliveryD30,AD.DeliveryD31  " & vbCrLf & _
                              "  		)detail2 " & vbCrLf & _
                              " 		ORDER BY sort, PartNos, NoUrut DESC  " & vbCrLf & _
                              " 	 END "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Select Case Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, pDate)))
                    Case 28
                        grid.Columns("DeliveryD29").Visible = False
                        grid.Columns("DeliveryD30").Visible = False
                        grid.Columns("DeliveryD31").Visible = False

                    Case 29
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = False
                        grid.Columns("DeliveryD31").Visible = False

                    Case 30
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = True
                        grid.Columns("DeliveryD31").Visible = False

                    Case 31
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = True
                        grid.Columns("DeliveryD31").Visible = True
                End Select                
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub SaveDataMaster(ByVal pAffCode As String, _
                                ByVal pPONo As String, _
                                ByVal pSuppCode As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("PO")
                    Dim sqlComm As New SqlCommand()
                    ls_SQL = "  IF NOT EXISTS (SELECT * FROM dbo.Affiliate_Master WHERE PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSuppCode) & "')  " & vbCrLf & _
                              "  BEGIN  " & vbCrLf & _
                              "     INSERT INTO dbo.Affiliate_Master " & vbCrLf & _
                              "          ( PONo ,AffiliateID ,SupplierID ,EntryDate ,EntryUser ,UpdateDate ,UpdateUSer) " & vbCrLf & _
                              "     VALUES  ( '" & Trim(pPONo) & "' , '" & Trim(pAffCode) & "' , '" & Trim(pSuppCode) & "' , GETDATE(), '" & Session("UserID") & "' , getdate() ,  '" & Session("UserID") & "')  " & vbCrLf & _
                              "  END  " & vbCrLf & _
                              " ELSE  " & vbCrLf & _
                              "  BEGIN  " & vbCrLf & _
                              "     UPDATE dbo.Affiliate_Master  " & vbCrLf & _
                              "     SET UpdateDate = GETDATE() " & vbCrLf & _
                              "     ,UpdateUSer= '" & Session("UserID") & "' "

                    ls_SQL = ls_SQL + "  WHERE PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSuppCode) & "' " & vbCrLf & _
                                      " END "

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
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

    Private Sub SaveDataDetail(ByVal pAffCode As String, _
                               ByVal pPONo As String, _
                               ByVal pSuppCode As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_DeliveryD1 As Double = 0 : Dim ls_DeliveryD2 As Double = 0 : Dim ls_DeliveryD3 As Double = 0 : Dim ls_DeliveryD4 As Double = 0 : Dim ls_DeliveryD5 As Double = 0
        Dim ls_DeliveryD6 As Double = 0 : Dim ls_DeliveryD7 As Double = 0 : Dim ls_DeliveryD8 As Double = 0 : Dim ls_DeliveryD9 As Double = 0 : Dim ls_DeliveryD10 As Double = 0
        Dim ls_DeliveryD11 As Double = 0 : Dim ls_DeliveryD12 As Double = 0 : Dim ls_DeliveryD13 As Double = 0 : Dim ls_DeliveryD14 As Double = 0 : Dim ls_DeliveryD15 As Double = 0
        Dim ls_DeliveryD16 As Double = 0 : Dim ls_DeliveryD17 As Double = 0 : Dim ls_DeliveryD18 As Double = 0 : Dim ls_DeliveryD19 As Double = 0 : Dim ls_DeliveryD20 As Double = 0
        Dim ls_DeliveryD21 As Double = 0 : Dim ls_DeliveryD22 As Double = 0 : Dim ls_DeliveryD23 As Double = 0 : Dim ls_DeliveryD24 As Double = 0 : Dim ls_DeliveryD25 As Double = 0
        Dim ls_DeliveryD26 As Double = 0 : Dim ls_DeliveryD27 As Double = 0 : Dim ls_DeliveryD28 As Double = 0 : Dim ls_DeliveryD29 As Double = 0 : Dim ls_DeliveryD30 As Double = 0
        Dim ls_DeliveryD31 As Double = 0

        Dim ls_POQty As Double = 0
        Dim ls_POQtyOld As Double = 0

        Dim admin As String = Session("UserID").ToString

        Try
            Dim iLoop As Long = 0, jLoop As Long = 0

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("SaveDetail")
                    If grid.VisibleRowCount = 0 Then
                        ls_MsgID = "6011"
                        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                        Session("ZZ010Msg") = lblInfo.Text
                        Exit Sub
                    End If
                    With grid
                        For iLoop = 0 To grid.VisibleRowCount - 1
                            Dim ls_Kanban As String = .GetRowValues(iLoop, "KanbanCls").ToString()
                            ls_DeliveryD1 = .GetRowValues(iLoop, "DeliveryD1")
                            ls_DeliveryD2 = .GetRowValues(iLoop, "DeliveryD2")
                            ls_DeliveryD3 = .GetRowValues(iLoop, "DeliveryD3")
                            ls_DeliveryD4 = .GetRowValues(iLoop, "DeliveryD4")
                            ls_DeliveryD5 = .GetRowValues(iLoop, "DeliveryD5")
                            ls_DeliveryD6 = .GetRowValues(iLoop, "DeliveryD6")
                            ls_DeliveryD7 = .GetRowValues(iLoop, "DeliveryD7")
                            ls_DeliveryD8 = .GetRowValues(iLoop, "DeliveryD8")
                            ls_DeliveryD9 = .GetRowValues(iLoop, "DeliveryD9")
                            ls_DeliveryD10 = .GetRowValues(iLoop, "DeliveryD10")
                            ls_DeliveryD11 = .GetRowValues(iLoop, "DeliveryD11")
                            ls_DeliveryD12 = .GetRowValues(iLoop, "DeliveryD12")
                            ls_DeliveryD13 = .GetRowValues(iLoop, "DeliveryD13")
                            ls_DeliveryD14 = .GetRowValues(iLoop, "DeliveryD14")
                            ls_DeliveryD15 = .GetRowValues(iLoop, "DeliveryD15")
                            ls_DeliveryD16 = .GetRowValues(iLoop, "DeliveryD16")
                            ls_DeliveryD17 = .GetRowValues(iLoop, "DeliveryD17")
                            ls_DeliveryD18 = .GetRowValues(iLoop, "DeliveryD18")
                            ls_DeliveryD19 = .GetRowValues(iLoop, "DeliveryD19")
                            ls_DeliveryD20 = .GetRowValues(iLoop, "DeliveryD20")
                            ls_DeliveryD21 = .GetRowValues(iLoop, "DeliveryD21")
                            ls_DeliveryD22 = .GetRowValues(iLoop, "DeliveryD22")
                            ls_DeliveryD23 = .GetRowValues(iLoop, "DeliveryD23")
                            ls_DeliveryD24 = .GetRowValues(iLoop, "DeliveryD24")
                            ls_DeliveryD25 = .GetRowValues(iLoop, "DeliveryD25")
                            ls_DeliveryD26 = .GetRowValues(iLoop, "DeliveryD26")
                            ls_DeliveryD27 = .GetRowValues(iLoop, "DeliveryD27")
                            ls_DeliveryD28 = .GetRowValues(iLoop, "DeliveryD28")
                            ls_DeliveryD29 = .GetRowValues(iLoop, "DeliveryD29")
                            ls_DeliveryD30 = .GetRowValues(iLoop, "DeliveryD30")
                            ls_DeliveryD31 = .GetRowValues(iLoop, "DeliveryD31")


                            If ls_Kanban = "YES" Then ls_Kanban = "1" Else ls_Kanban = "0"
                            Dim byWhat As String = .GetRowValues(iLoop, "BYWHAT")
                            If byWhat = "BY AFFILIATE" Then 'OLD
                                ls_POQtyOld = .GetRowValues(iLoop, "POQty")
                                ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.Affiliate_Detail WHERE PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSuppCode) & "' AND PartNo='" & .GetRowValues(iLoop, "PartNos").ToString & "')  " & vbCrLf & _
                                          " BEGIN  " & vbCrLf & _
                                          " 	INSERT INTO dbo.Affiliate_Detail " & vbCrLf & _
                                          "         ( PONo , " & vbCrLf & _
                                          "           AffiliateID , " & vbCrLf & _
                                          "           SupplierID , " & vbCrLf & _
                                          "           PartNo , " & vbCrLf & _                                          
                                          "           POQtyOld , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD1Old , " & vbCrLf & _
                                                  "           DeliveryD2Old , " & vbCrLf & _
                                                  "           DeliveryD3Old , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD4Old , " & vbCrLf & _
                                                  "           DeliveryD5Old , " & vbCrLf & _
                                                  "           DeliveryD6Old , " & vbCrLf & _
                                                  "           DeliveryD7Old , " & vbCrLf & _
                                                  "           DeliveryD8Old , " & vbCrLf & _
                                                  "           DeliveryD9Old , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD10Old , " & vbCrLf & _
                                                  "           DeliveryD11Old , " & vbCrLf & _
                                                  "           DeliveryD12Old , " & vbCrLf & _
                                                  "           DeliveryD13Old , " & vbCrLf & _
                                                  "           DeliveryD14Old , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD15Old , " & vbCrLf & _
                                                  "           DeliveryD16Old , " & vbCrLf & _
                                                  "           DeliveryD17Old , " & vbCrLf & _
                                                  "           DeliveryD18Old , " & vbCrLf & _
                                                  "           DeliveryD19Old , " & vbCrLf & _
                                                  "           DeliveryD20Old , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD21Old , " & vbCrLf & _
                                                  "           DeliveryD22Old , " & vbCrLf & _
                                                  "           DeliveryD23Old , " & vbCrLf & _
                                                  "           DeliveryD24Old , " & vbCrLf & _
                                                  "           DeliveryD25Old , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD26Old , " & vbCrLf & _
                                                  "           DeliveryD27Old , " & vbCrLf & _
                                                  "           DeliveryD28Old , " & vbCrLf & _
                                                  "           DeliveryD29Old , " & vbCrLf & _
                                                  "           DeliveryD30Old , " & vbCrLf & _
                                                  "           DeliveryD31Old , " & vbCrLf

                                ls_SQL = ls_SQL + "           EntryDate , " & vbCrLf & _
                                                  "           EntryUser , " & vbCrLf & _
                                                  "           UpdateDate , " & vbCrLf & _
                                                  "           UpdateUser " & vbCrLf & _
                                                  "         ) " & vbCrLf & _
                                                  " 	VALUES  ( '" & Trim(pPONo) & "' , -- PONo - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pAffCode) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pSuppCode) & "' , -- SupplierID - char(20) " & vbCrLf & _
                                                  "           '" & .GetRowValues(iLoop, "PartNos").ToString & "' , -- PartNo - char(25) " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_POQtyOld & " , -- POQtyOld - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD1 & " , -- DeliveryD1Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD2 & " , -- DeliveryD2Old - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD3 & " , -- DeliveryD3Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD4 & " , -- DeliveryD4Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD5 & " , -- DeliveryD5Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD6 & " , -- DeliveryD6Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD7 & " , -- DeliveryD7Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD8 & " , -- DeliveryD8Old - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD9 & " , -- DeliveryD9Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD10 & " , -- DeliveryD10Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD11 & " , -- DeliveryD11Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD12 & " , -- DeliveryD12Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD13 & " , -- DeliveryD13Old - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD14 & " , -- DeliveryD14Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD15 & " , -- DeliveryD15Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD16 & " , -- DeliveryD16Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD17 & " , -- DeliveryD17Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD18 & " , -- DeliveryD18Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD19 & " , -- DeliveryD19Old - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD20 & " , -- DeliveryD20Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD21 & " , -- DeliveryD21Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD22 & " , -- DeliveryD22Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD23 & " , -- DeliveryD23Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD24 & " , -- DeliveryD24Old - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD25 & " , -- DeliveryD25Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD26 & " , -- DeliveryD26Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD27 & " , -- DeliveryD27Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD28 & " , -- DeliveryD28Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD29 & " , -- DeliveryD29Old - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD30 & " , -- DeliveryD30Old - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD31 & " , -- DeliveryD31Old - numeric " & vbCrLf & _
                                                  "           getdate() , -- EntryDate - datetime " & vbCrLf & _
                                                  "           '" & Session("UserID") & "' , -- EntryUser - char(15) " & vbCrLf & _
                                                  "           getdate() , -- UpdateDate - datetime " & vbCrLf & _
                                                  "           '" & Session("UserID") & "'  -- UpdateUser - char(15) " & vbCrLf & _
                                                  "         ) " & vbCrLf & _
                                                  "         END	 " & vbCrLf & _
                                                  "         ELSE	 " & vbCrLf & _
                                                  "         BEGIN  " & vbCrLf & _
                                                  "            UPDATE [dbo].[Affiliate_Detail] " & vbCrLf

                                ls_SQL = ls_SQL + " 		   SET [POQtyOld] = " & ls_POQtyOld & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD1Old] = " & ls_DeliveryD1 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD2Old] = " & ls_DeliveryD2 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD3Old] = " & ls_DeliveryD3 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD4Old] = " & ls_DeliveryD4 & "" & vbCrLf & _
                                                  " 			  ,[DeliveryD5Old] =" & ls_DeliveryD5 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD6Old] = " & ls_DeliveryD6 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD7Old] =" & ls_DeliveryD7 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD8Old] = " & ls_DeliveryD8 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD9Old] = " & ls_DeliveryD9 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD10Old] = " & ls_DeliveryD10 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD11Old] = " & ls_DeliveryD11 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD12Old] = " & ls_DeliveryD12 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD13Old] = " & ls_DeliveryD13 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD14Old] = " & ls_DeliveryD14 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD15Old] = " & ls_DeliveryD15 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD16Old] = " & ls_DeliveryD16 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD17Old] = " & ls_DeliveryD17 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD18Old] = " & ls_DeliveryD18 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD19Old] = " & ls_DeliveryD19 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD20Old] = " & ls_DeliveryD20 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD21Old] = " & ls_DeliveryD21 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD22Old] = " & ls_DeliveryD22 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD23Old] = " & ls_DeliveryD23 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD24Old] = " & ls_DeliveryD24 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD25Old] = " & ls_DeliveryD25 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD26Old] = " & ls_DeliveryD26 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD27Old] = " & ls_DeliveryD27 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD28Old] = " & ls_DeliveryD28 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD29Old] = " & ls_DeliveryD29 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD30Old] = " & ls_DeliveryD30 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD31Old] = " & ls_DeliveryD31 & "  " & vbCrLf & _
                                                  " 			  ,[UpdateDate] = getdate() " & vbCrLf & _
                                                  " 			  ,[UpdateUser] = '" & Session("UserID") & "' " & vbCrLf & _
                                                  " 			WHERE [PONo] = '" & Trim(pPONo) & "' " & vbCrLf & _
                                                  " 			  AND [AffiliateID] ='" & Trim(pAffCode) & "' " & vbCrLf & _
                                                  " 			  AND [SupplierID] = '" & Trim(pSuppCode) & "'" & vbCrLf

                                ls_SQL = ls_SQL + " 			  AND [PartNo] = '" & .GetRowValues(iLoop, "PartNos") & "' " & vbCrLf & _
                                                  " 		 END  "


                                ls_MsgID = "1002"

                                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()
                            Else
                                'BY PASI New
                                ls_POQty = .GetRowValues(iLoop, "POQty")
                                ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.Affiliate_Detail WHERE PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSuppCode) & "' AND PartNo='" & .GetRowValues(iLoop, "PartNos").ToString & "')  " & vbCrLf & _
                                          " BEGIN  " & vbCrLf & _
                                          " 	INSERT INTO dbo.Affiliate_Detail " & vbCrLf & _
                                          "         ( PONo , " & vbCrLf & _
                                          "           AffiliateID , " & vbCrLf & _
                                          "           SupplierID , " & vbCrLf & _
                                          "           PartNo , " & vbCrLf

                                ls_SQL = ls_SQL + "           POQty , " & vbCrLf & _
                                                  "           DeliveryD1 , " & vbCrLf & _
                                                  "           DeliveryD2 , " & vbCrLf & _
                                                  "           DeliveryD3 , " & vbCrLf & _
                                                  "           DeliveryD4 , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD5 , " & vbCrLf & _
                                                  "           DeliveryD6 , " & vbCrLf & _
                                                  "           DeliveryD7 , " & vbCrLf & _
                                                  "           DeliveryD8 , " & vbCrLf & _
                                                  "           DeliveryD9 , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD10 , " & vbCrLf & _
                                                  "           DeliveryD11 , " & vbCrLf & _
                                                  "           DeliveryD12 , " & vbCrLf & _
                                                  "           DeliveryD13 , " & vbCrLf & _
                                                  "           DeliveryD14 , " & vbCrLf & _
                                                  "           DeliveryD15 , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD16 , " & vbCrLf & _
                                                  "           DeliveryD17 , " & vbCrLf & _
                                                  "           DeliveryD18 , " & vbCrLf & _
                                                  "           DeliveryD19 , " & vbCrLf & _
                                                  "           DeliveryD20 , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD21 , " & vbCrLf & _
                                                  "           DeliveryD22 , " & vbCrLf & _
                                                  "           DeliveryD23 , " & vbCrLf & _
                                                  "           DeliveryD24 , " & vbCrLf & _
                                                  "           DeliveryD25 , " & vbCrLf & _
                                                  "           DeliveryD26 , " & vbCrLf

                                ls_SQL = ls_SQL + "           DeliveryD27 , " & vbCrLf & _
                                                  "           DeliveryD28 , " & vbCrLf & _
                                                  "           DeliveryD29 , " & vbCrLf & _
                                                  "           DeliveryD30 , " & vbCrLf & _
                                                  "           DeliveryD31 , " & vbCrLf

                                ls_SQL = ls_SQL + "           EntryDate , " & vbCrLf & _
                                                  "           EntryUser , " & vbCrLf & _
                                                  "           UpdateDate , " & vbCrLf & _
                                                  "           UpdateUser " & vbCrLf & _
                                                  "         ) " & vbCrLf & _
                                                  " 	VALUES  ( '" & Trim(pPONo) & "' , -- PONo - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pAffCode) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pSuppCode) & "' , -- SupplierID - char(20) " & vbCrLf & _
                                                  "           '" & .GetRowValues(iLoop, "PartNos").ToString & "' , -- PartNo - char(25) " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_POQty & " , -- POQty - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD1 & " , -- DeliveryD1 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD2 & " , -- DeliveryD2 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD3 & " , -- DeliveryD3 - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD4 & " , -- DeliveryD4 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD5 & " , -- DeliveryD5 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD6 & " , -- DeliveryD6 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD7 & " , -- DeliveryD7 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD8 & " , -- DeliveryD8 - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD9 & " , -- DeliveryD9 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD10 & " , -- DeliveryD10 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD11 & " , -- DeliveryD11 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD12 & " , -- DeliveryD12 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD13 & " , -- DeliveryD13 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD14 & " , -- DeliveryD14 - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD15 & " , -- DeliveryD15 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD16 & " , -- DeliveryD16 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD17 & " , -- DeliveryD17 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD18 & " , -- DeliveryD18 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD19 & " , -- DeliveryD19 - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD20 & " , -- DeliveryD20 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD21 & " , -- DeliveryD21 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD22 & " , -- DeliveryD22 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD23 & " , -- DeliveryD23 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD24 & " , -- DeliveryD24 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD25 & " , -- DeliveryD25 - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD26 & " , -- DeliveryD26 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD27 & " , -- DeliveryD27 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD28 & " , -- DeliveryD28 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD29 & " , -- DeliveryD29 - numeric " & vbCrLf & _
                                                  "           " & ls_DeliveryD30 & " , -- DeliveryD30 - numeric " & vbCrLf

                                ls_SQL = ls_SQL + "           " & ls_DeliveryD31 & " , -- DeliveryD31 - numeric " & vbCrLf & _
                                                  "           getdate() , -- EntryDate - datetime " & vbCrLf & _
                                                  "           '" & Session("UserID") & "' , -- EntryUser - char(15) " & vbCrLf & _
                                                  "           getdate() , -- UpdateDate - datetime " & vbCrLf & _
                                                  "           '" & Session("UserID") & "'  -- UpdateUser - char(15) " & vbCrLf & _
                                                  "         ) " & vbCrLf & _
                                                  "         END	 " & vbCrLf & _
                                                  "         ELSE	 " & vbCrLf & _
                                                  "         BEGIN  " & vbCrLf & _
                                                  "            UPDATE [dbo].[Affiliate_Detail] " & vbCrLf

                                ls_SQL = ls_SQL + " 		   SET [POQty] = " & ls_POQty & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD1] = " & ls_DeliveryD1 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD2] = " & ls_DeliveryD2 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD3] = " & ls_DeliveryD3 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD4] = " & ls_DeliveryD4 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD5] = " & ls_DeliveryD5 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD6] =  " & ls_DeliveryD6 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD7] =  " & ls_DeliveryD7 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD8] =  " & ls_DeliveryD8 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD9] =  " & ls_DeliveryD9 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD10] = " & ls_DeliveryD10 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD11] = " & ls_DeliveryD11 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD12] = " & ls_DeliveryD12 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD13] = " & ls_DeliveryD13 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD14] = " & ls_DeliveryD14 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD15] = " & ls_DeliveryD15 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD16] = " & ls_DeliveryD16 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD17] = " & ls_DeliveryD17 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD18] = " & ls_DeliveryD18 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD19] = " & ls_DeliveryD19 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD20] = " & ls_DeliveryD20 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD21] = " & ls_DeliveryD21 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD22] = " & ls_DeliveryD22 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD23] = " & ls_DeliveryD23 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD24] = " & ls_DeliveryD24 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD25] = " & ls_DeliveryD25 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD26] = " & ls_DeliveryD26 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD27] = " & ls_DeliveryD27 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD28] = " & ls_DeliveryD28 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD29] = " & ls_DeliveryD29 & " " & vbCrLf

                                ls_SQL = ls_SQL + " 			  ,[DeliveryD30] = " & ls_DeliveryD30 & " " & vbCrLf & _
                                                  " 			  ,[DeliveryD31] = " & ls_DeliveryD31 & " " & vbCrLf & _
                                                  " 			  ,[UpdateDate] = getdate() " & vbCrLf & _
                                                  " 			  ,[UpdateUser] = '" & Session("UserID") & "' " & vbCrLf & _
                                                  " 			WHERE [PONo] = '" & Trim(pPONo) & "' " & vbCrLf & _
                                                  " 			  AND [AffiliateID] ='" & Trim(pAffCode) & "' " & vbCrLf & _
                                                  " 			  AND [SupplierID] = '" & Trim(pSuppCode) & "'" & vbCrLf

                                ls_SQL = ls_SQL + " 			  AND [PartNo] = '" & .GetRowValues(iLoop, "PartNos") & "' " & vbCrLf & _
                                                  " 		 END  "


                                ls_MsgID = "1002"

                                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()
                            End If
EndNext:
                        Next iLoop


                        sqlTran.Commit()
                        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
                        If lblInfo.Text = "[] " Then lblInfo.Text = ""
                        Session("ZZ010Msg") = lblInfo.Text
                    End With
                End Using

                sqlConn.Close()

            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub SaveDeliveryCls()
        Dim ls_sql
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("item_code")

                ls_sql = "Update PO_Master set DeliveryByPASICls = '" & IIf(rblDelivery.Value = 0, 0, 1) & "' where PONo = '" & cboPONo.Text & "' and AffiliateID = '" & cboAffiliateCode.Text & "' and SupplierID = '" & txtSupplierCode.Text & "'"

                Dim SqlComm6 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                SqlComm6.ExecuteNonQuery()
                SqlComm6.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using
    End Sub

    Private Sub UpdatePO(ByVal pAffCode As String, ByVal pPONo As String, ByVal pSuppCode As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " UPDATE dbo.PO_Master " & vbCrLf & _
                          " SET PASISendAffiliateUser='" & admin & "' " & vbCrLf & _
                          " ,PASISendAffiliateDate=getdate() " & vbCrLf & _
                          " WHERE PONo='" & pPONo & "'  " & vbCrLf & _
                          " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                          " AND SupplierID='" & pSuppCode & "' "

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub UpdateExcel(ByVal pAffCode As String, _
                            ByVal pPONo As String, _
                            ByVal pSuppCode As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " UPDATE dbo.Affiliate_Master " & vbCrLf & _
                          " SET ExcelCls='1'" & vbCrLf & _
                          " WHERE PONo='" & pPONo & "'  " & vbCrLf & _
                          " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                          " AND SupplierID='" & pSuppCode & "' "

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

#End Region

End Class