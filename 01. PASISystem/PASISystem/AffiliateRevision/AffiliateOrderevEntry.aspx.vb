Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports Microsoft.Office.Interop
Imports System.Net
Imports System.Net.Mail


Public Class AffiliateOrderevEntry
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True
    Dim pub_Period As Date
    Dim pub_PORev As String
    Dim pub_PO As String
    Dim pub_Commercial As String
    Dim pub_AffiliateID As String
    Dim pub_AffiliateName As String
    Dim pub_SupplierID As String
    Dim pub_SupplierName As String
    Dim pub_ShipBy As String
    Dim pub_Kanban As String
    Dim pub_SeqNo As String
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
    Dim FlagGrid As Integer
    Dim UpdateSend As Boolean = False
#End Region

#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                Session("M01Url") = Request.QueryString("Session")
                flag = False
            Else
                flag = True
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Session("M01Url") <> "" Then
                    If (Not String.IsNullOrEmpty(Request.QueryString("id")))  Then
                        Session("MenuDesc") = "AFFILIATE ORDER REVISION DETAIL"
                        'up_Fillcombo()
                        pub_Period = Request.QueryString("t1")
                        pub_PORev = Request.QueryString("t2")
                        pub_PO = Request.QueryString("t3")
                        pub_Commercial = Request.QueryString("t4")
                        pub_AffiliateID = Request.QueryString("t5")
                        pub_AffiliateName = Request.QueryString("t6")
                        pub_SupplierID = Request.QueryString("t7")
                        pub_SupplierName = Request.QueryString("t8")
                        pub_Kanban = Request.QueryString("t9")
                        pub_ShipBy = Request.QueryString("t10")
                        pub_SeqNo = Request.QueryString("t11")
                        'tabIndex()
                        pSearch = False
                        bindDataHeader(pub_Period, pub_PORev, pub_PO, pub_AffiliateID, pub_SupplierID, pub_Kanban)
                        bindDataDetail(pub_Period, pub_PORev, pub_PO, pub_AffiliateID, pub_SupplierID, pub_Kanban)
                        Call SaveDataMaster(ValidasiInput(pub_AffiliateID), pub_Period, pub_PORev, pub_PO, Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), pub_Commercial, Trim(rblPOKanban.Value), Trim(txtShipBy.Text), pub_SeqNo)
                        Call SaveDataDetail(ValidasiInput(pub_AffiliateID), pub_Period, pub_PORev, pub_PO, Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), pub_Commercial, Trim(rblPOKanban.Value), Trim(txtShipBy.Text))
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        pSearch = True
                        If txtSuppAppDate.Text <> "" Or txtSuppPendDate.Text <> "" Or txtSuppUnpDate.Text <> "" Or txtPASIAppDate.Text <> "" Or txtAffFinalAppDate.Text <> "" Then
                            btnSubmit.Enabled = False
                            btnSendSupplier.Enabled = False
                        End If
                        'btnClear.Visible = False
                        'ScriptManager.RegisterStartupScript(AffiliateSubmit, AffiliateSubmit.GetType(), "scriptKey", "txtAffiliateID.GetInputElement().setAttribute('style', 'background:#FFFFFF;');", True)
                    ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                        Session("MenuDesc") = "AFFILIATE ORDER REVISION DETAIL"

                        pub_Period = clsNotification.DecryptURL(Request.QueryString("t1"))
                        pub_PORev = clsNotification.DecryptURL(Request.QueryString("t2"))
                        pub_PO = clsNotification.DecryptURL(Request.QueryString("id2"))
                        pub_AffiliateID = clsNotification.DecryptURL(Request.QueryString("t5"))
                        pub_SupplierID = clsNotification.DecryptURL(Request.QueryString("t7"))
                        pub_Kanban = clsNotification.DecryptURL(Request.QueryString("t9"))
                        
                        pSearch = False
                        bindDataHeader(pub_Period, pub_PORev, pub_PO, pub_AffiliateID, pub_SupplierID, pub_Kanban)
                        bindDataDetail(pub_Period, pub_PORev, pub_PO, pub_AffiliateID, pub_SupplierID, pub_Kanban)
                        Call SaveDataMaster(ValidasiInput(pub_AffiliateID), pub_Period, pub_PORev, pub_PO, Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), pub_Commercial, Trim(rblPOKanban.Value), Trim(txtShipBy.Text), pub_SeqNo)
                        Call SaveDataDetail(ValidasiInput(pub_AffiliateID), pub_Period, pub_PORev, pub_PO, Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), pub_Commercial, Trim(rblPOKanban.Value), Trim(txtShipBy.Text))
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        pSearch = True
                        If txtSuppAppDate.Text <> "" Or txtSuppPendDate.Text <> "" Or txtSuppUnpDate.Text <> "" Or txtPASIAppDate.Text <> "" Or txtAffFinalAppDate.Text <> "" Then
                            btnSubmit.Enabled = False
                            btnSendSupplier.Enabled = False
                        End If
                    Else
                        Session("MenuDesc") = "AFFILIATE ORDER REVISION DETAIL"
                        'tabIndex()
                        'clear()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        btnClear.Visible = True
                    End If
                Else
                    btnClear.Visible = True
                    'txtAffiliateID.Focus()
                    'tabIndex()
                    'clear()
                End If
            End If

            'If ls_AllowDelete = False Then btnDelete.Enabled = False
            If ls_AllowUpdate = False Then btnSubmit.Enabled = False

            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MenuID As String = "", ls_MsgID As String = ""
        Dim iLoop As Long = 0, jLoop As Long = 0
        'Dim ls_PartNo As String = "", ls_Kanban As String = "", ls_Maker As String = "", ls_POqty As String = ""
        Dim ls_UserID As String = ""
        If getApp(Trim(txtPORev.Text), Trim(txtPONo.Text)) = True Then
            ls_MsgID = "6029"
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            Session("ZZ010Msg") = lblInfo.Text
            Exit Sub
        End If

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
                    If e.UpdateValues(iLoop).NewValues("POQty") = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "5001", clsMessage.MsgType.ErrorMessage)
                        Session("YA010IsSubmit") = lblInfo.Text
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        errorBatch = True
                        Exit Sub
                    End If
                    If (e.UpdateValues(iLoop).NewValues("POQty") Mod e.UpdateValues(iLoop).NewValues("MinOrderQty")) <> 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "5005", clsMessage.MsgType.ErrorMessage)
                        Session("YA010IsSubmit") = lblInfo.Text
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        errorBatch = True
                        Exit Sub
                    End If

                    Dim ls_PartNo As String = (e.UpdateValues(iLoop).OldValues("PartNos").ToString())
                    Dim ls_SeqNo As String = (e.UpdateValues(iLoop).OldValues("SeqNo").ToString())
                    Dim ls_Kanban As String = (e.UpdateValues(iLoop).OldValues("KanbanCls").ToString())
                    If ls_Kanban = "YES" Then ls_Kanban = "1" Else ls_Kanban = "0"
                    Dim ls_Maker As String = (e.UpdateValues(iLoop).OldValues("Maker").ToString())
                    ls_POqty = e.UpdateValues(iLoop).NewValues("POQty")
                    'Dim ls_POqtyOld As Double = e.UpdateValues(iLoop).OldValues("POQty")
                    Dim ls_DiffCls As String = ""
                    'If ls_POqty = ls_POqtyOld Then
                    '    ls_DiffCls = "0"
                    'Else
                    '    ls_DiffCls = "1"
                    'End If
                    ' FlagGrid = e.UpdateValues(iLoop).NewValues("Flag")
                    'Dim ls_CurrAff As String = e.UpdateValues(iLoop).OldValues("CurrCodeAff").ToString()
                    'Dim ls_PriceAff As Double = (e.UpdateValues(iLoop).OldValues("PriceAff").ToString())
                    'Dim ls_AmountAff As Double = e.UpdateValues(iLoop).OldValues("AmountAff")
                    Dim ls_CurrAff As String = ""
                    Dim ls_PriceAff As Double = 0
                    Dim ls_AmountAff As Double = 0
                    'Dim ls_CurrSupp As String = If(IsDBNull(e.UpdateValues(iLoop).OldValues("CurrCodeSupp").ToString()), "", e.UpdateValues(iLoop).OldValues("CurrCodeSupp").ToString())
                    'Dim ls_PriceSupp As Double = If(IsDBNull(e.UpdateValues(iLoop).OldValues("PriceSupp")), 0, e.UpdateValues(iLoop).OldValues("PriceSupp"))
                    'Dim ls_AmountSupp As Double = If(IsDBNull(e.UpdateValues(iLoop).OldValues("AmountSupp")), 0, e.UpdateValues(iLoop).OldValues("AmountSupp"))
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


                    'ls_MenuID = Trim(e.UpdateValues(iLoop).NewValues("MenuID").ToString())
                    'ls_SQL = " IF EXISTS (SELECT * FROM dbo.Affiliate_Detail WHERE PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' AND PartNo='" & ls_PartNo & "') " & vbCrLf & _
                    '  " BEGIN " & vbCrLf & _
                    '  " INSERT INTO dbo.Affiliate_Detail " & vbCrLf & _
                    '  "         ( PONo ,AffiliateID ,SupplierID ,PartNo ,KanbanCls ,Maker ,POQty ,CurrCls , " & vbCrLf & _
                    '  "           Price ,Amount ,ForecastN1 ,ForecastN2 ,ForecastN3 , " & vbCrLf & _
                    '  "           DeliveryD1 ,DeliveryD2 ,DeliveryD3 ,DeliveryD4 ,DeliveryD5 ,DeliveryD6 ,DeliveryD7 ,DeliveryD8 ,DeliveryD9 ,DeliveryD10 , " & vbCrLf & _
                    '  "           DeliveryD11 ,DeliveryD12 ,DeliveryD13 ,DeliveryD14 ,DeliveryD15 ,DeliveryD16 ,DeliveryD17 ,DeliveryD18 ,DeliveryD19 ,DeliveryD20 , " & vbCrLf & _
                    '  "           DeliveryD21 ,DeliveryD22 ,DeliveryD23 ,DeliveryD24 ,DeliveryD25 ,DeliveryD26 ,DeliveryD27 ,DeliveryD28 ,DeliveryD29 ,DeliveryD30 ,DeliveryD31 , " & vbCrLf & _
                    '  "           EntryDate ,EntryUser) " & vbCrLf & _
                    '  " VALUES  ( '" & Trim(txtPONo.Text) & "' , '" & Trim(txtAffiliateID.Text) & "' , '" & Trim(txtSupplierCode.Text) & "' , '" & ls_PartNo & "' , '" & ls_Kanban & "' , '" & ls_Maker & "' , " & ls_POqty & " , '" & ls_CurrSupp & "' ,  " & vbCrLf & _
                    '  "           " & ls_PriceAff & " , " & ls_AmountAff & " , " & ls_ForeCast1 & " , " & ls_ForeCast2 & " , " & ls_ForeCast3 & " ,  "

                    'ls_SQL = ls_SQL + "           " & ls_DeliveryD1 & " , " & ls_DeliveryD2 & " , " & ls_DeliveryD3 & " , " & ls_DeliveryD4 & " , " & ls_DeliveryD5 & " , " & ls_DeliveryD6 & " , " & ls_DeliveryD7 & " , " & ls_DeliveryD8 & " , " & ls_DeliveryD9 & " , " & ls_DeliveryD10 & " ,  " & vbCrLf & _
                    '                  "           " & ls_DeliveryD11 & " , " & ls_DeliveryD12 & " , " & ls_DeliveryD13 & " , " & ls_DeliveryD14 & " , " & ls_DeliveryD15 & " , " & ls_DeliveryD16 & " , " & ls_DeliveryD17 & " , " & ls_DeliveryD18 & " , " & ls_DeliveryD19 & " , " & ls_DeliveryD20 & " ,  " & vbCrLf & _
                    '                  "           " & ls_DeliveryD21 & " , " & ls_DeliveryD22 & " , " & ls_DeliveryD23 & " , " & ls_DeliveryD24 & " , " & ls_DeliveryD25 & " , " & ls_DeliveryD26 & " , " & ls_DeliveryD27 & " , " & ls_DeliveryD28 & " , " & ls_DeliveryD29 & " , " & ls_DeliveryD30 & " , " & ls_DeliveryD31 & " ,  " & vbCrLf & _
                    '                  "           getdate() , '" & Session("UserID") & "'  ) " & vbCrLf & _
                    '                  "           END " & vbCrLf & _
                    '                  "           ELSE " & vbCrLf & _
                    '                  "           BEGIN " & vbCrLf & _
                    '                  "           UPDATE dbo.Affiliate_Detail " & vbCrLf & _
                    '                  "           SET  " & vbCrLf & _
                    '                  "           KanbanCls ='" & ls_Kanban & "',Maker ='" & ls_Maker & "',POQty =" & ls_POqty & ",CurrCls = '" & ls_CurrSupp & "', " & vbCrLf & _
                    '                  "           Price =" & ls_PriceAff & ",Amount =" & ls_AmountAff & ",ForecastN1 =" & ls_ForeCast1 & ",ForecastN2 =" & ls_ForeCast2 & ",ForecastN3 =" & ls_ForeCast3 & ", "

                    'ls_SQL = ls_SQL + "           DeliveryD1 = " & ls_DeliveryD1 & ",DeliveryD2 = " & ls_DeliveryD1 & ",DeliveryD3 = " & ls_DeliveryD1 & ",DeliveryD4 = " & ls_DeliveryD1 & ",DeliveryD5 = " & ls_DeliveryD1 & ",DeliveryD6 = " & ls_DeliveryD1 & ",DeliveryD7 = " & ls_DeliveryD1 & ",DeliveryD8 = " & ls_DeliveryD1 & ",DeliveryD9 = " & ls_DeliveryD1 & ",DeliveryD10 = " & ls_DeliveryD1 & ", " & vbCrLf & _
                    '                  "           DeliveryD11 = " & ls_DeliveryD1 & ",DeliveryD12 = " & ls_DeliveryD1 & ",DeliveryD13 = " & ls_DeliveryD1 & ",DeliveryD14 = " & ls_DeliveryD1 & ",DeliveryD15 = " & ls_DeliveryD1 & ",DeliveryD16 = " & ls_DeliveryD1 & ",DeliveryD17 = " & ls_DeliveryD1 & ",DeliveryD18 = " & ls_DeliveryD1 & ",DeliveryD19 = " & ls_DeliveryD1 & ",DeliveryD20 = " & ls_DeliveryD1 & ", " & vbCrLf & _
                    '                  "           DeliveryD21 = " & ls_DeliveryD1 & ",DeliveryD22 = " & ls_DeliveryD1 & ",DeliveryD23 = " & ls_DeliveryD1 & ",DeliveryD24 = " & ls_DeliveryD1 & ",DeliveryD25 = " & ls_DeliveryD1 & ",DeliveryD26 = " & ls_DeliveryD1 & ",DeliveryD27 = " & ls_DeliveryD1 & ",DeliveryD28 = " & ls_DeliveryD1 & ",DeliveryD29 = " & ls_DeliveryD1 & ",DeliveryD30 = " & ls_DeliveryD1 & ",DeliveryD31 = " & ls_DeliveryD1 & " " & vbCrLf & _
                    '                  "           WHERE PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' AND PartNo='" & ls_PartNo & "' " & vbCrLf & _
                    '                  "            " & vbCrLf & _
                    '                  "           END "

                    ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(txtPORev.Text) & "' AND PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' AND PartNo='" & ls_PartNo & "')  " & vbCrLf & _
                  " BEGIN  " & vbCrLf & _
                  " 	INSERT INTO dbo.AffiliateRev_Detail " & vbCrLf & _
                  "         ( PORevNo, " & vbCrLf & _
                  "           PONo , " & vbCrLf & _
                  "           AffiliateID , " & vbCrLf & _
                  "           SupplierID , " & vbCrLf & _
                  "           PartNo , " & vbCrLf & _
                  "           SeqNo , " & vbCrLf & _
                  "           DifferenceCls , " & vbCrLf & _
                  "           --KanbanCls , " & vbCrLf & _
                  "           Maker , " & vbCrLf & _
                  "           POQty , " & vbCrLf

                    ls_SQL = ls_SQL + "           --POQtyOld , " & vbCrLf & _
                                      "           CurrCls , " & vbCrLf & _
                                      "           Price , " & vbCrLf & _
                                      "           Amount , " & vbCrLf & _
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
                                      " 	VALUES  ( '" & Trim(txtPORev.Text) & "' , -- PORevNo - char(20) " & vbCrLf & _
                                      "           '" & Trim(txtPONo.Text) & "' , -- PONo - char(20) " & vbCrLf & _
                                      "           '" & Trim(txtAffiliateID.Text) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                      "           '" & Trim(txtSupplierCode.Text) & "' , -- SupplierID - char(20) " & vbCrLf & _
                                      "           '" & ls_PartNo & "' , -- PartNo - char(25) " & vbCrLf & _
                                      "           '" & ls_SeqNo & "' , -- SeqNo - char(25) " & vbCrLf & _
                                      "           '" & ls_DiffCls & "' , -- DifferenceCls - char(1) " & vbCrLf & _
                                      "           --'" & ls_Kanban & "' , -- KanbanCls - char(1) " & vbCrLf

                    ls_SQL = ls_SQL + "           '" & ls_Maker & "', -- Maker - char(20) " & vbCrLf & _
                                      "           " & ls_POqty & " , -- POQty - numeric " & vbCrLf & _
                                      "           --" & ls_POqty & " , -- POQtyOld - numeric " & vbCrLf & _
                                      "           '" & ls_CurrAff & "' , -- CurrCls - char(2) " & vbCrLf & _
                                      "           " & ls_PriceAff & " , -- Price - numeric " & vbCrLf & _
                                      "           " & ls_AmountAff & " , -- Amount - numeric " & vbCrLf & _
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
                                      "            UPDATE [dbo].[AffiliateRev_Detail] " & vbCrLf

                    ls_SQL = ls_SQL + " 		   SET [SeqNo] = '" & ls_SeqNo & "' " & vbCrLf & _
                                      " 			  ,[Maker] = '" & ls_Maker & "' " & vbCrLf & _
                                      " 			  ,[POQty] = " & ls_POqty & " " & vbCrLf & _
                                      " 			  --,[POQtyOld] = " & ls_POqty & " " & vbCrLf & _
                                      " 			  ,[CurrCls] = '" & ls_CurrAff & "' " & vbCrLf & _
                                      " 			  ,[Price] = " & ls_PriceAff & " " & vbCrLf & _
                                      " 			  ,[Amount] = " & ls_AmountAff & " " & vbCrLf & _
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
                                      " 			WHERE PORevNo='" & Trim(txtPORev.Text) & "' " & vbCrLf & _
                                      "               AND [PONo] = '" & Trim(txtPONo.Text) & "' " & vbCrLf & _
                                      " 			  AND [AffiliateID] ='" & Trim(txtAffiliateID.Text) & "' " & vbCrLf & _
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
            'grid.JSProperties("cpMessage") = Session("AA220Msg")
            Dim ls_MsgID As String
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Dim pDate As Date = Split(e.Parameters, "|")(1)
            Dim pPORevNo As String = Split(e.Parameters, "|")(2)
            Dim pPONo As String = Split(e.Parameters, "|")(3)
            Dim pAffCode As String = Split(e.Parameters, "|")(4)
            Dim pSuppCode As String = Split(e.Parameters, "|")(5)
            Dim pKanban As String = Split(e.Parameters, "|")(6)
            'Dim pShipBy As String = Split(e.Parameters, "|")(7)
            Select Case pAction
                Case "load"
                    pSearch = True
                    Call bindDataHeader(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)
                    Call bindDataDetail(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)
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
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliateID)
                    'Call SaveDataDetail(lb_IsUpdate, pDate, pAffCode, pPONo, pSuppCode, pShipBy, pDelBy, pKanban)
                    'Call SaveDataMaster(lb_IsUpdate, pDate, pAffCode, pPONo, pSuppCode, pShipBy, pDelBy, pKanban)
                    Call bindDataHeader(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)
                    Call bindDataDetail(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)
                    ls_MsgID = "1001"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Session("YA010IsSubmit") = lblInfo.Text
                Case "send"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliateID)
                    UpdateSend = True
                    Call UpdatePO(lb_IsUpdate, pAffCode, pPORevNo, pPONo, pSuppCode)
                    Call bindDataHeader(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)
                    Call bindDataDetail(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)
                    Call UpdateExcel(lb_IsUpdate, pAffCode, pPORevNo, pPONo, pSuppCode)
                    'Call Excel()
                    UpdateSend = False
            End Select

EndProcedure:
            Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        Dim x As Integer = CInt(e.VisibleIndex.ToString())

        If x > grid.VisibleRowCount Then Exit Sub
        If e.GetValue("BYWHAT") = "REV. BY AFFILIATE" Then
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
                If e.GetValue("BYWHAT") = "REV. BY AFFILIATE" Then
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
                If e.GetValue("BYWHAT") = "REV. BY PASI" Then
                    'If e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" Then
                    '    e.Cell.Text = ""
                    'End If
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
            Response.Redirect("~/AffiliateRevision/AffiliateOrderRevList.aspx")
        Else
            'Session.Remove("M01Url")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    'Private Sub txtPONo_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles txtPONo.Callback
    '    If String.IsNullOrEmpty(e.Parameter) Then
    '        Return
    '    End If

    '    Dim ls_value As String = Split(e.Parameter, "|")(0)
    '    Dim ls_sql As String = ""

    '    ls_sql = "SELECT '" & clsGlobal.gs_All & "' PONo UNION ALL SELECT RTRIM(PONo)PONo FROM dbo.PO_Master WHERE AffiliateID='" & ls_value & "' " & vbCrLf
    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)

    '        With txtPONo
    '            .Items.Clear()
    '            .Columns.Clear()
    '            .DataSource = ds.Tables(0)
    '            .Columns.Add("PONo")
    '            .Columns(0).Width = 50

    '            .TextField = "PONo"
    '            .DataBind()
    '            .SelectedIndex = 0
    '        End With

    '        sqlConn.Close()
    '    End Using
    'End Sub

    'Private Sub cbPONo_Callback(source As Object, e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbPONo.Callback

    '    Dim ls_sql As String = ""

    '    Dim pAction As String = Split(e.Parameter, "|")(0)
    '    Dim pDate As Date = Split(e.Parameter, "|")(1)
    '    Dim pPONo As String = Split(e.Parameter, "|")(2)
    '    Dim pAffCode As String = Split(e.Parameter, "|")(3)


    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        ls_sql = "   SELECT DISTINCT Period,POD.AffiliateID,AffiliateName,POM.PONo  " & vbCrLf & _
    '              "   ,CASE WHEN CommercialCls = '0' THEN 'NO' ELSE 'YES' END CommercialCls  " & vbCrLf & _
    '              "   ,POD.SupplierID,SupplierName,ShipCls   " & vbCrLf & _
    '              "   ,PODeliveryBy   " & vbCrLf & _
    '              "   ,MP.KanbanCls   " & vbCrLf & _
    '              "   FROM dbo.PO_Master POM    " & vbCrLf & _
    '              "   LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID  AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
    '              "   LEFT JOIN dbo.MS_Affiliate MA ON POD.AffiliateID = MA.AffiliateID  " & vbCrLf & _
    '              "   LEFT JOIN dbo.MS_Parts MP ON POD.PartNo = MP.PartNo   " & vbCrLf & _
    '              "   LEFT JOIN dbo.MS_Supplier MS ON POD.SupplierID = MS.SupplierID   " & vbCrLf & _
    '              "  WHERE YEAR(Period) = YEAR('" & pDate & "') AND MONTH(Period) = MONTH('" & pDate & "')  "

    '        ls_sql = ls_sql + "  AND POM.PONo = '" & pPONo & "'    " & vbCrLf & _
    '                          "  AND POM.AffiliateID='" & pAffCode & "' "


    '        Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)

    '        With ds.Tables(0)
    '            If ds.Tables(0).Rows.Count > 0 Then
    '                cbPONo.JSProperties("cpCommercialCls") = .Rows(0).Item("CommercialCls")
    '                cbPONo.JSProperties("cpSupplierID") = .Rows(0).Item("SupplierID")
    '                cbPONo.JSProperties("cpSupplierName") = .Rows(0).Item("SupplierName")
    '                cbPONo.JSProperties("cpShipCls") = .Rows(0).Item("ShipCls")
    '                cbPONo.JSProperties("cpPODeliveryBy") = .Rows(0).Item("PODeliveryBy")
    '                cbPONo.JSProperties("cpKanbanCls") = .Rows(0).Item("KanbanCls")
    '            Else
    '                cbPONo.JSProperties("cpCommercialCls") = ""
    '                cbPONo.JSProperties("cpSupplierID") = ""
    '                cbPONo.JSProperties("cpSupplierName") = ""
    '                cbPONo.JSProperties("cpShipCls") = ""
    '                cbPONo.JSProperties("cpPODeliveryBy") = 0
    '                cbPONo.JSProperties("cpKanbanCls") = 2
    '            End If
    '        End With

    '        sqlConn.Close()
    '    End Using
    'End Sub

    Private Sub grid_PageIndexChanged(sender As Object, e As System.EventArgs) Handles grid.PageIndexChanged
        Call bindDataDetail(Trim(txtPeriod.Text), Trim(txtPORev.Text), Trim(txtPONo.Text), Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), rblPOKanban.Value)
    End Sub
#End Region

#Region "PROCEDURE"

    Private Sub bindDataHeader(ByVal pDate As Date, ByVal pPORevNo As String, ByVal pPONo As String, ByVal pAffCode As String, ByVal pSupplierID As String, ByVal pKanban As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "   SELECT DISTINCT PORM.Period,PORD.AffiliateID,AffiliateName,PORM.PORevNo,PORM.PONo  " & vbCrLf & _
                  "   ,CASE WHEN POM.CommercialCls = '0' THEN 'NO' ELSE 'YES' END CommercialCls  " & vbCrLf & _
                  "   ,PORD.SupplierID,SupplierName,POM.ShipCls   " & vbCrLf & _
                  "   ,PODeliveryBy   " & vbCrLf & _
                  "   ,MP.KanbanCls   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.EntryDate,120)EntryDate,ISNULL(PORM.EntryUser,'')EntryUser --1   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.AffiliateApproveDate,120)AffiliateApproveDate,ISNULL(PORM.AffiliateApproveUser,'')AffiliateApproveUser --2   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.PASISendAffiliateDate,120)PASISendAffiliateDate,ISNULL(PORM.PASISendAffiliateUser,'')PASISendAffiliateUser --3   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.SupplierApproveDate,120)SupplierApproveDate,ISNULL(PORM.SupplierApproveUser,'')SupplierApproveUser --4   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.SupplierApprovePendingDate,120)SupplierApprovePendingDate,ISNULL(PORM.SupplierApprovePendingUser,'')SupplierApprovePendingUser --5   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.SupplierUnApproveDate,120)SupplierUnApproveDate,ISNULL(PORM.SupplierUnApproveUser,'')SupplierUnApproveUser --6   " & vbCrLf

            ls_SQL = ls_SQL + "   ,CONVERT(DATETIME,PORM.PASIApproveDate,120)PASIApproveDate,ISNULL(PORM.PASIApproveUser ,'')PASIApproveUser --7   " & vbCrLf & _
                              "   ,CONVERT(DATETIME,PORM.FinalApproveDate,120)FinalApproveDate,ISNULL(PORM.FinalApproveUser,'')FinalApproveUser --8    " & vbCrLf & _
                              "   ,PORD.PartNo,PartName,CASE WHEN MP.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,UnitCls,MOQ,QtyBox,CASE WHEN MP.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,UnitCls,MOQ,QtyBox   " & vbCrLf & _                              
                              "   ,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & pDate & "'))),0) " & vbCrLf & _
                              "   ,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & pDate & "'))),0) " & vbCrLf & _
                              "   ,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & pDate & "'))),0) " & vbCrLf & _
                              "   ,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10   " & vbCrLf & _
                              "   ,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20   " & vbCrLf & _
                              "   ,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31   " & vbCrLf & _
                              "   FROM dbo.PORev_Master PORM    " & vbCrLf

            ls_SQL = ls_SQL + "   LEFT JOIN dbo.PORev_Detail PORD ON PORM.PORevNo = PORD.PORevNo AND PORM.PONo = PORD.PONo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID " & vbCrLf & _
                              "   LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID  " & vbCrLf & _
                              "   LEFT JOIN dbo.PO_Detail POD ON PORM.PONo = POD.PONo AND PORM.AffiliateID = POD.AffiliateID AND PORM.SupplierID = POD.SupplierID  " & vbCrLf & _
                              "   LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID  " & vbCrLf & _
                              "   LEFT JOIN dbo.MS_Parts MP ON PORD.PartNo = MP.PartNo   " & vbCrLf & _
                              "   LEFT JOIN MS_PartMapping MPM ON MPM.SupplierID = PORD.SupplierID and MPM.PartNo = PORD.PartNo and MPM.AffiliateID = PORD.AffiliateID  " & vbCrLf & _
                              "   LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID   " & vbCrLf & _
                              "   LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID  " & vbCrLf

            ls_SQL = ls_SQL + " WHERE MONTH(PORM.Period) = MONTH('" & pDate & "') AND YEAR(PORM.Period) = YEAR('" & pDate & "')" & vbCrLf & _
                              " AND PORM.PORevNo = '" & pPORevNo & "' AND PORM.PONo='" & pPONo & "' AND PORM.AffiliateID='" & pAffCode & "' AND PORM.SupplierID='" & pSupplierID & "'   " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND POD.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtPeriod.Value = Format(ds.Tables(0).Rows(0)("Period"), "MMM yyyy")
                txtAffiliateID.Text = ds.Tables(0).Rows(0)("AffiliateID")
                txtAffiliateName.Text = ds.Tables(0).Rows(0)("AffiliateName")
                txtPORev.Text = ds.Tables(0).Rows(0)("PORevNo")
                txtPONo.Text = ds.Tables(0).Rows(0)("PONo")
                txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                txtSupplierCode.Text = ds.Tables(0).Rows(0)("SupplierID")
                txtSupplierName.Text = ds.Tables(0).Rows(0)("SupplierName")
                txtShipBy.Text = ds.Tables(0).Rows(0)("ShipCls")
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


    Private Sub bindDataDetail(ByVal pDate As Date, ByVal pPORevNo As String, ByVal pPONo As String, ByVal pAffCode As String, ByVal pSupplierID As String, ByVal pKanban As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "    IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(pPORevNo) & "' AND PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSupplierID) & "')   " & vbCrLf & _
                  "    BEGIN   " & vbCrLf & _
                  "    SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description    " & vbCrLf & _
                  "       ,MOQ ,MinOrderQty,SeqNo, QtyBox ,Maker  ,ISNULL(MonthlyProductionCapacity,0)MonthlyProductionCapacity   " & vbCrLf & _
                  "       ,BYWHAT,POQty ,ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                  "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                  "       FROM (    " & vbCrLf

            ls_SQL = ls_SQL + "  		SELECT CONVERT(CHAR,row_number() over (order by PORD.PONo)) as NoUrut,PORD.PartNo,PORD.PartNo PartNos,PartName       " & vbCrLf & _
                              "        	,CASE WHEN MPART.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,MU.DESCRIPTION  " & vbCrLf & _
                              "        	,MOQ =CONVERT(CHAR,MOQ),MinOrderQty = MOQ,PORM.SeqNo,QtyBox = CONVERT(CHAR,QtyBox),ISNULL(MPART.Maker,'')Maker       " & vbCrLf & _
                              "          ,(SELECT ISNULL(QtyRemaining, MonthlyProductionCapacity) from MS_SupplierCapacity A  " & vbCrLf & _
                              "             LEFT JOIN RemainingCapacity B ON A.PartNo = B.PartNo AND A.SupplierID = B.SupplierID AND PORD.PartNo = B.PartNo  " & vbCrLf & _
                              "             WHERE B.Period = '" & Format(pDate, "yyyyMM") & "' AND A.SupplierID='" & pSupplierID.Trim & "') MonthlyProductionCapacity   " & vbCrLf & _
                              "         ,'REV. BY AFFILIATE' BYWHAT,PORD.POQty ,hari = CEILING((SUM(PORD.POQty)/QtyBox))  " & vbCrLf & _
                              "       	,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,PORM.Period))),0)    " & vbCrLf & _
                              "    	    ,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,PORM.Period))),0)    " & vbCrLf & _
                              "    		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,PORM.Period))),0)                              " & vbCrLf & _
                              "      	,PORD.DeliveryD1 ,PORD.DeliveryD2 ,PORD.DeliveryD3 ,PORD.DeliveryD4 ,PORD.DeliveryD5 ,PORD.DeliveryD6 ,PORD.DeliveryD7 ,PORD.DeliveryD8 ,PORD.DeliveryD9 ,PORD.DeliveryD10  " & vbCrLf

            ls_SQL = ls_SQL + "      	,PORD.DeliveryD11 ,PORD.DeliveryD12 ,PORD.DeliveryD13 ,PORD.DeliveryD14 ,PORD.DeliveryD15 ,PORD.DeliveryD16 ,PORD.DeliveryD17 ,PORD.DeliveryD18 ,PORD.DeliveryD19 ,PORD.DeliveryD20  " & vbCrLf & _
                              "      	,PORD.DeliveryD21 ,PORD.DeliveryD22 ,PORD.DeliveryD23 ,PORD.DeliveryD24 ,PORD.DeliveryD25 ,PORD.DeliveryD26 ,PORD.DeliveryD27 ,PORD.DeliveryD28 ,PORD.DeliveryD29 ,PORD.DeliveryD30 ,PORD.DeliveryD31  " & vbCrLf & _
                              "      	,row_number() over (order by PORD.PONo) as Sort      " & vbCrLf & _
                              "      	FROM dbo.PORev_Master PORM      " & vbCrLf & _
                              "  		INNER JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.MS_Parts MPART ON PORD.PartNo = MPART.PartNo         " & vbCrLf & _
                              "         LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = PORD.PartNo and MPM.SupplierID = PORD.SupplierID and MPM.AffiliateID = PORD.AffiliateID  " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID          " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls                 " & vbCrLf & _                              
                              "         WHERE MONTH(PORM.Period) = MONTH('" & pDate & "') AND YEAR(PORM.Period) = YEAR('" & pDate & "')  " & vbCrLf & _
                              "         AND PORM.PORevNo='" & Trim(pPORevNo) & "' AND PORM.PONo='" & pPONo.Trim & "' AND PORM.SupplierID='" & Trim(pSupplierID) & "'  " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND MPART.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If

            ls_SQL = ls_SQL + "            GROUP BY PORD.PONo,PORD.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ " & vbCrLf & _
                              "            ,PORM.SeqNo,QtyBox,PORD.poqty,MPART.Maker,MonthlyProductionCapacity,PORM.Period,MSC.PartNo,PORM.AffiliateID    " & vbCrLf & _
                              "      	   ,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10          " & vbCrLf & _
                              "      	   ,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20        		    " & vbCrLf & _
                              "      	   ,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31     " & vbCrLf & _
                              " 	) Detail1    " & vbCrLf & _
                              "  	UNION ALL    " & vbCrLf

            ls_SQL = ls_SQL + "      SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              "       ,MOQ = MOQ ,MinOrderQty,SeqNo,QtyBox = QtyBox ,Maker ,MonthlyProductionCapacity ,BYWHAT    " & vbCrLf & _
                              "       ,POQty, ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                              "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              "       FROM (    " & vbCrLf & _
                              "      		SELECT '' as NoUrut,'' PartNo,PORD.PartNo PartNos,''PartName,'' KanbanCls,''Description,MOQ = '',MinOrderQty = MOQ ,PORM.SeqNo    " & vbCrLf

            ls_SQL = ls_SQL + "  			,'' QtyBox,ISNULL(MPART.Maker,'')Maker,0 MonthlyProductionCapacity  " & vbCrLf & _
                              "  			,'REV. BY PASI' BYWHAT,PORD.POQty  " & vbCrLf & _
                              "      		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,PORM.Period))),0)    " & vbCrLf & _
                              "    			,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,PORM.Period))),0)    " & vbCrLf & _
                              "    			,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,PORM.Period))),0)     " & vbCrLf & _
                              "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6 ,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10 " & vbCrLf & _
                              "      		,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20 " & vbCrLf & _
                              "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31 " & vbCrLf & _
                              "      		,row_number() over (order by PORD.PONo) as Sort      " & vbCrLf & _
                              "      	FROM dbo.PORev_Master PORM      " & vbCrLf & _
                              "  		 LEFT JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID   " & vbCrLf

            ls_SQL = ls_SQL + "  		 LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  		 LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Parts MPART ON PORD.PartNo = MPART.PartNo         " & vbCrLf & _
                              "         LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = PORD.PartNo and MPM.SupplierID = PORD.SupplierID and MPM.AffiliateID = PORD.AffiliateID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls     " & vbCrLf & _                              
                              "         WHERE MONTH(PORM.Period) = MONTH('" & pDate & "') AND YEAR(PORM.Period) = YEAR('" & pDate & "')  " & vbCrLf & _
                              "         AND PORM.PORevNo='" & pPORevNo.Trim & "' AND PORM.PONo='" & pPONo.Trim & "' AND PORM.SupplierID='" & pSupplierID.Trim & "'  " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND MPART.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If

            ls_SQL = ls_SQL + _
                              "            GROUP BY PORD.PONo,PORD.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ " & vbCrLf

            ls_SQL = ls_SQL + "             ,PORM.SeqNo,QtyBox,PORD.POQty,MPART.Maker,MonthlyProductionCapacity ,PORM.Period,MSC.PartNo,PORM.AffiliateID   " & vbCrLf & _
                              "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10        " & vbCrLf & _
                              "      		,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20          " & vbCrLf & _
                              "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31   " & vbCrLf & _
                              "  	)detail2    " & vbCrLf & _
                              "  ORDER BY sort, PartNo DESC    " & vbCrLf & _
                              "  END   " & vbCrLf & _
                              "  ELSE   " & vbCrLf & _
                              "  BEGIN   " & vbCrLf & _
                              "  SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description     " & vbCrLf & _
                              "    ,MOQ ,MinOrderQty,SeqNo,QtyBox,Maker ,ISNULL(MonthlyProductionCapacity,0)MonthlyProductionCapacity ,BYWHAT     " & vbCrLf & _
                              "    ,POQty ,ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31    " & vbCrLf & _
                              "     FROM (   " & vbCrLf & _
                              "        SELECT row_number() over (order by AD.PONo) as Sort   ,CONVERT(CHAR,row_number() over (order by AD.PONo)) as NoUrut     " & vbCrLf

            ls_SQL = ls_SQL + "    		,AD.PartNo as PartNo ,AD.PartNo AS PartNos,PartName ,CASE WHEN AD.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls ,MU.Description     " & vbCrLf & _
                              "    		,MOQ =CONVERT(CHAR,MOQ),MinOrderQty = MOQ,AD.SeqNo,QtyBox = CONVERT(CHAR,QtyBox) ,AD.Maker    " & vbCrLf & _
                              "         ,(SELECT ISNULL(QtyRemaining, MonthlyProductionCapacity) from MS_SupplierCapacity A   " & vbCrLf & _
                              "            LEFT JOIN RemainingCapacity B ON A.PartNo = B.PartNo AND A.SupplierID = B.SupplierID AND AD.PartNo = B.PartNo   " & vbCrLf & _
                              "             WHERE B.Period = '" & Format(pDate, "yyyyMM") & "' AND A.SupplierID='" & pSupplierID.Trim & "') MonthlyProductionCapacity    " & vbCrLf & _
                              "         ,'REV. BY AFFILIATE' BYWHAT     " & vbCrLf & _
                              "    		,POQtyOld POqty   " & vbCrLf & _
                              "    		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,PORM.Period))),0)     " & vbCrLf & _
                              "     	,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,PORM.Period))),0)     " & vbCrLf & _
                              "     	,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,PORM.Period))),0)                               " & vbCrLf & _
                              "       	,DeliveryD1Old DeliveryD1,DeliveryD2Old DeliveryD2,DeliveryD3Old DeliveryD3,DeliveryD4Old DeliveryD4,DeliveryD5Old DeliveryD5    " & vbCrLf

            ls_SQL = ls_SQL + "    		,DeliveryD6Old DeliveryD6,DeliveryD7Old DeliveryD7,DeliveryD8Old DeliveryD8,DeliveryD9Old DeliveryD9,DeliveryD10Old DeliveryD10    " & vbCrLf & _
                              "    		,DeliveryD11Old DeliveryD11,DeliveryD12Old DeliveryD12,DeliveryD13Old DeliveryD13,DeliveryD14Old DeliveryD14,DeliveryD15Old DeliveryD15    " & vbCrLf & _
                              "    		,DeliveryD16Old DeliveryD16,DeliveryD17Old DeliveryD17,DeliveryD18 DeliveryD18,DeliveryD19Old DeliveryD19,DeliveryD20Old DeliveryD20    " & vbCrLf & _
                              "    		,DeliveryD21Old DeliveryD21,DeliveryD22Old DeliveryD22,DeliveryD23Old DeliveryD23,DeliveryD24Old DeliveryD24,DeliveryD25Old DeliveryD25    " & vbCrLf & _
                              "    		,DeliveryD26Old DeliveryD26,DeliveryD27Old DeliveryD27,DeliveryD28Old DeliveryD28,DeliveryD29Old DeliveryD29,DeliveryD30Old DeliveryD30,DeliveryD31Old DeliveryD31    " & vbCrLf & _
                              "    		FROM dbo.AffiliateRev_Detail AD " & vbCrLf & _
                              "    		LEFT JOIN PORev_Master PORM ON PORM.PONo = AD.PONo AND PORM.PORevNo = AD.PORevNo AND PORM.AffiliateID = AD.AffiliateID AND PORM.SupplierID = AD.SupplierID " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo    " & vbCrLf & _
                              "         LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = AD.PartNo and MPM.SupplierID = AD.SupplierID and MPM.AffiliateID = AD.AffiliateID  " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID           " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID       " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "    		LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls           " & vbCrLf & _                              
                              "    		WHERE AD.PORevNo='" & pPORevNo.Trim & "' AND AD.PONo='" & pPONo.Trim & "' AND AD.SupplierID='" & pSupplierID.Trim & "'   " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND MPART.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If

            ls_SQL = ls_SQL + "    		GROUP BY AD.PONo,AD.PartNo,PartName,AD.KanbanCls,POQtyOld,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity,AD.SeqNo,AD.AffiliateID     " & vbCrLf & _
                              "   		,MSC.PartNo,PORM.Period,DeliveryD1Old,DeliveryD2Old,DeliveryD3Old,DeliveryD4Old,DeliveryD5Old    " & vbCrLf & _
                              "    		,DeliveryD6Old,DeliveryD7Old,DeliveryD8Old,DeliveryD9Old,DeliveryD10Old    " & vbCrLf & _
                              "    		,DeliveryD11Old,DeliveryD12Old,DeliveryD13Old,DeliveryD14Old,DeliveryD15Old    " & vbCrLf & _
                              "    		,DeliveryD16Old,DeliveryD17Old,DeliveryD18,DeliveryD19Old,DeliveryD20Old    " & vbCrLf & _
                              "    		,DeliveryD21Old,DeliveryD22Old,DeliveryD23Old,DeliveryD24Old,DeliveryD25Old    " & vbCrLf & _
                              "    		,DeliveryD26Old,DeliveryD27Old,DeliveryD28Old,DeliveryD29Old,DeliveryD30Old,DeliveryD31Old  )detail1    " & vbCrLf & _
                              "   	UNION ALL    "

            ls_SQL = ls_SQL + "   	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description     " & vbCrLf & _
                              "        ,MOQ = MOQ ,MinOrderQty,SeqNo,QtyBox = QtyBox ,Maker ,MonthlyProductionCapacity ,BYWHAT     " & vbCrLf & _
                              "        ,POqty ,ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              "        FROM (     " & vbCrLf & _
                              "    		SELECT row_number() over (order by AD.PONo) as Sort,'' as NoUrut,'' PartNo,AD.PartNo PartNos,''PartName,'' KanbanCls,''Description,MOQ = ''      "

            ls_SQL = ls_SQL + "    		,MinOrderQty = MOQ,AD.SeqNo,'' QtyBox,ISNULL(AD.Maker,'')Maker ,0 MonthlyProductionCapacity,'REV. BY PASI' BYWHAT ,POQty   " & vbCrLf & _
                              "    		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,PORM.Period))),0)     " & vbCrLf & _
                              "     	,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,PORM.Period))),0)     " & vbCrLf & _
                              "     	,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,PORM.Period))),0)                               " & vbCrLf & _
                              "       	,DeliveryD1 ,DeliveryD2 ,DeliveryD3 ,DeliveryD4 ,DeliveryD5     " & vbCrLf & _
                              "   		,DeliveryD6 ,DeliveryD7 ,DeliveryD8 ,DeliveryD9 ,DeliveryD10     " & vbCrLf & _
                              "   		,DeliveryD11 ,DeliveryD12 ,DeliveryD13 ,DeliveryD14     " & vbCrLf & _
                              "   		,DeliveryD15 ,DeliveryD16 ,DeliveryD17 ,DeliveryD18,DeliveryD19 ,DeliveryD20 ,DeliveryD21     " & vbCrLf & _
                              "   		,DeliveryD22 ,DeliveryD23 ,DeliveryD24 ,DeliveryD25 ,DeliveryD26 ,DeliveryD27 ,DeliveryD28 ,DeliveryD29     " & vbCrLf & _
                              "   		,DeliveryD30 ,DeliveryD31     " & vbCrLf & _
                              "   		FROM dbo.AffiliateRev_Detail AD  "

            ls_SQL = ls_SQL + "   		 LEFT JOIN PORev_Master PORM ON PORM.PONo = AD.PONo AND PORM.PORevNo = AD.PORevNo AND PORM.AffiliateID = AD.AffiliateID AND PORM.SupplierID = AD.SupplierID   " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo    " & vbCrLf & _
                              "         LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = AD.PartNo and MPM.SupplierID = AD.SupplierID and MPM.AffiliateID = AD.AffiliateID  " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID    " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID       " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID           " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls    " & vbCrLf & _
                              "          WHERE AD.PORevNo='" & pPORevNo.Trim & "' AND AD.PONo='" & pPONo.Trim & "' AND AD.SupplierID='" & pSupplierID.Trim & "'   " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND MPART.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If

            ls_SQL = ls_SQL + "   		 GROUP BY AD.PONo,AD.PartNo,PartName,AD.KanbanCls,POQty,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity,AD.SeqNo,AD.AffiliateID   " & vbCrLf & _
                              "       		,MSC.PartNo,PORM.Period,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5    " & vbCrLf & _
                              "    			,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10    "

            ls_SQL = ls_SQL + "    			,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15    " & vbCrLf & _
                              "    			,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20    " & vbCrLf & _
                              "    			,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25    " & vbCrLf & _
                              "    			,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31   " & vbCrLf & _
                              "   		)detail2   	  " & vbCrLf & _
                              "  	ORDER BY sort, PartNo DESC   " & vbCrLf & _
                              "  END   " & vbCrLf & _
                              "    "




            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Dim pDateDay As DateTime = CDate(Format(pDate, "MM") + "/01/" + Format(pDate, "yyyy"))
                Select Case Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, pDateDay)))
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
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub SaveDataMaster(ByVal pIsNewData As Boolean, _
                         Optional ByVal pDate As String = "", _
                         Optional ByVal pPORevNo As String = "", _
                         Optional ByVal pPONo As String = "", _
                         Optional ByVal pAffCode As String = "", _
                         Optional ByVal pSuppCode As String = "", _
                         Optional ByVal pComm As String = "", _
                         Optional ByVal pKanban As String = "", _
                         Optional ByVal pShipBy As String = "", _
                         Optional ByVal pSeqNo As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("PO")
                    Dim sqlComm As New SqlCommand()
                    ls_SQL = "  IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Master WHERE PORevNo='" & Trim(pPORevNo) & "' AND PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSuppCode) & "')  " & vbCrLf & _
                  "  BEGIN  " & vbCrLf & _
                  "  INSERT INTO dbo.AffiliateRev_Master " & vbCrLf & _
                  "          ( PORevNo ,PONo ,AffiliateID ,SupplierID ,SeqNo ,EntryDate ,EntryUser ,UpdateDate ,UpdateUSer) " & vbCrLf & _
                  "  VALUES  ( '" & Trim(pPORevNo) & "' , '" & Trim(pPONo) & "' , '" & Trim(pAffCode) & "' ,'" & Trim(pSuppCode) & "', '" & Trim(pSeqNo) & "' , GETDATE(), '" & Session("UserID") & "' , getdate() ,  '" & Session("UserID") & "')  " & vbCrLf & _
                  "          END  " & vbCrLf & _
                  "          ELSE  " & vbCrLf & _
                  "          BEGIN  " & vbCrLf & _
                  "          UPDATE dbo.AffiliateRev_Master  " & vbCrLf & _
                  "          SET UpdateDate = GETDATE() " & vbCrLf & _
                  "          ,UpdateUSer= '" & Session("UserID") & "' " & vbCrLf

                    ls_SQL = ls_SQL + "  WHERE PORevNo='" & Trim(pPORevNo) & "' AND PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSuppCode) & "' " & vbCrLf & _
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

    Private Sub SaveDataDetail(ByVal pIsNewData As Boolean, _
                         Optional ByVal pDate As String = "", _
                         Optional ByVal pPORevNo As String = "", _
                         Optional ByVal pPONo As String = "", _
                         Optional ByVal pAffCode As String = "", _
                         Optional ByVal pSuppCode As String = "", _
                         Optional ByVal pComm As String = "", _
                         Optional ByVal pKanban As String = "", _
                         Optional ByVal pShipBy As String = "")

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

        Dim ls_diffCls As String = ""

        Dim ls_SeqNo As String


        Dim admin As String = Session("UserID").ToString

        Try
            Dim iLoop As Long = 0, jLoop As Long = 0
            Dim ls_UserID As String = ""

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
                            'If FlagGrid = 1 Then
                            '    GoTo EndNext
                            'End If
                            ls_SeqNo = .GetRowValues(iLoop, "SeqNo")
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
                            If byWhat = "REV. BY AFFILIATE" Then 'OLD
                                ls_POQtyOld = .GetRowValues(iLoop, "POQty")
                                If ls_POQty = ls_POQtyOld Then
                                    ls_diffCls = "0"
                                Else
                                    ls_diffCls = "1"
                                End If
                                'Dim ls_AmountAff As Double = .GetRowValues(iLoop, "PriceAff") * .GetRowValues(iLoop, "POQty")
                                Dim ls_AmountAff As Double = 0
                                ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(txtPORev.Text) & "' AND PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' AND PartNo='" & .GetRowValues(iLoop, "PartNos").ToString & "')  " & vbCrLf & _
                                  " BEGIN  " & vbCrLf & _
                                  " 	INSERT INTO dbo.AffiliateRev_Detail " & vbCrLf & _
                                  "         ( PORevNo, " & vbCrLf & _
                                  "           PONo , " & vbCrLf & _
                                  "           AffiliateID , " & vbCrLf & _
                                  "           SupplierID , " & vbCrLf & _
                                  "           PartNo , " & vbCrLf & _
                                  "           SeqNo , " & vbCrLf & _
                                  "           DifferenceCls , " & vbCrLf & _
                                  "           KanbanCls , " & vbCrLf & _
                                  "           Maker , " & vbCrLf & _
                                  "           POQtyOld , " & vbCrLf

                                ls_SQL = ls_SQL + "           CurrCls , " & vbCrLf & _
                                                  "           Price , " & vbCrLf & _
                                                  "           Amount , " & vbCrLf & _
                                                  "           DeliveryD1Old , " & vbCrLf & _
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
                                                  " 	VALUES  ( '" & Trim(pPORevNo) & "' , -- PORevNo - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pPONo) & "' , -- PONo - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pAffCode) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pSuppCode) & "' , -- SupplierID - char(20) " & vbCrLf & _
                                                  "           '" & .GetRowValues(iLoop, "PartNos").ToString & "' , -- PartNo - char(25) " & vbCrLf & _
                                                  "           '" & ls_diffCls & "' , -- PartNo - char(25) " & vbCrLf & _
                                                  "           '" & ls_SeqNo & "' , -- PartNo - char(25) " & vbCrLf & _
                                                  "           '" & ls_Kanban & "' , -- KanbanCls - char(1) " & vbCrLf


                                ls_SQL = ls_SQL + "           '" & .GetRowValues(iLoop, "Maker") & "', -- Maker - char(20) " & vbCrLf & _
                                                  "           " & ls_POQtyOld & " , -- POQtyOld - numeric " & vbCrLf & _
                                                  "           '' , -- CurrCls - char(2) " & vbCrLf & _
                                                  "           '0' , -- Price - numeric " & vbCrLf & _
                                                  "           " & ls_AmountAff & " , -- Amount - numeric " & vbCrLf & _
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
                                                  "            UPDATE [dbo].[AffiliateRev_Detail] " & vbCrLf

                                ls_SQL = ls_SQL + " 		   SET [KanbanCls] = '" & ls_Kanban & "' " & vbCrLf & _
                                                  "               ,[Maker] = '" & .GetRowValues(iLoop, "Maker") & "' " & vbCrLf & _
                                                  " 			  ,[POQtyOld] = " & ls_POQtyOld & " " & vbCrLf & _
                                                  " 			  ,[CurrCls] = '' " & vbCrLf & _
                                                  " 			  ,[Price] = 0 " & vbCrLf & _
                                                  " 			  ,[Amount] = " & ls_AmountAff & " " & vbCrLf & _
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
                                                  " 			  ,[DeliveryD18Old] = " & ls_DeliveryD18 & " "

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
                                                  " 			WHERE PORevNo='" & Trim(txtPORev.Text) & "' " & vbCrLf & _
                                                  "               AND [PONo] = '" & Trim(txtPONo.Text) & "' " & vbCrLf & _
                                                  " 			  AND [AffiliateID] ='" & Trim(txtAffiliateID.Text) & "' " & vbCrLf & _
                                                  " 			  AND [SupplierID] = '" & Trim(txtSupplierCode.Text) & "'" & vbCrLf

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
                                If ls_POQty = ls_POQtyOld Then
                                    ls_diffCls = "1"
                                Else
                                    ls_diffCls = "0"
                                End If
                                ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(txtPORev.Text) & "' AND PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' AND PartNo='" & .GetRowValues(iLoop, "PartNos").ToString & "')  " & vbCrLf & _
                                  " BEGIN  " & vbCrLf & _
                                  " 	INSERT INTO dbo.AffiliateRev_Detail " & vbCrLf & _
                                  "         ( PORevNo, " & vbCrLf & _
                                  "           PONo , " & vbCrLf & _
                                  "           AffiliateID , " & vbCrLf & _
                                  "           SupplierID , " & vbCrLf & _
                                  "           PartNo , " & vbCrLf & _
                                  "           --KanbanCls , " & vbCrLf & _
                                  "           Maker , " & vbCrLf

                                ls_SQL = ls_SQL + "           POQty , " & vbCrLf & _
                                                  "           CurrCls , " & vbCrLf & _
                                                  "           Price , " & vbCrLf & _
                                                  "           Amount , " & vbCrLf & _
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
                                                  " 	VALUES  ( '" & Trim(pPORevNo) & "' , -- PORevNo - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pPONo) & "' , -- PONo - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pAffCode) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                                  "           '" & Trim(pSuppCode) & "' , -- SupplierID - char(20) " & vbCrLf & _
                                                  "           '" & .GetRowValues(iLoop, "PartNos").ToString & "' , -- PartNo - char(25) " & vbCrLf & _
                                                  "           --'" & Trim(ls_Kanban) & "' , -- KanbanCls - char(1) " & vbCrLf

                                ls_SQL = ls_SQL + "           '" & .GetRowValues(iLoop, "Maker") & "', -- Maker - char(20) " & vbCrLf & _
                                                  "           " & ls_POQty & " , -- POQtyOld - numeric " & vbCrLf & _
                                                  "           '' , -- CurrCls - char(2) " & vbCrLf & _
                                                  "           0 , -- Price - numeric " & vbCrLf & _
                                                  "           0 , -- Amount - numeric " & vbCrLf & _
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
                                                  "            UPDATE [dbo].[AffiliateRev_Detail] " & vbCrLf

                                ls_SQL = ls_SQL + " 		   SET --[KanbanCls] = '" & ls_Kanban & "' " & vbCrLf & _
                                                  " 			  [Maker] = '" & .GetRowValues(iLoop, "Maker") & "' " & vbCrLf & _
                                                  " 			  ,[POQty] = " & ls_POQty & " " & vbCrLf & _
                                                  " 			  ,[CurrCls] = '' " & vbCrLf & _
                                                  " 			  ,[Price] = 0 " & vbCrLf & _
                                                  " 			  ,[Amount] = 0 " & vbCrLf & _
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
                                                  " 			WHERE PORevNo='" & Trim(txtPORev.Text) & "' " & vbCrLf & _
                                                  "               AND [PONo] = '" & Trim(txtPONo.Text) & "' " & vbCrLf & _
                                                  " 			  AND [AffiliateID] ='" & Trim(txtAffiliateID.Text) & "' " & vbCrLf & _
                                                  " 			  AND [SupplierID] = '" & Trim(txtSupplierCode.Text) & "'" & vbCrLf

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

    Private Sub UpdatePO(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffCode As String = "", _
                         Optional ByVal pPORevNo As String = "", _
                         Optional ByVal pPONo As String = "", _
                         Optional ByVal pSuppCode As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " UPDATE dbo.PORev_Master " & vbCrLf & _
                          " SET PASISendAffiliateUser='" & admin & "' " & vbCrLf & _
                          " ,PASISendAffiliateDate=getdate() " & vbCrLf & _
                          " WHERE PORevNo='" & pPORevNo & "' " & vbCrLf & _
                          " AND PONo='" & pPONo & "'  " & vbCrLf & _
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
            '        lblInfo.Text = "Affiliate ID with ID " & txtAffiliateID.Text & " already exists in the database."
            '        Return False
            '    End If
            '    Return True
            '    sqlConn.Close()
            'End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

    End Function

    Private Function BindDataExcel() As DataSet
        Dim ls_SQL As String = ""
        Dim tanggal As Date = FormatDateTime(Trim(txtPeriod.Text), DateFormat.ShortDate)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(txtPORev.Text) & "' AND PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "')   " & vbCrLf & _
                  " BEGIN  " & vbCrLf & _
                  " 	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
                  " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
                  "       ,MOQ = LEFT(MOQ,LEN(MOQ)-3) , QtyBox = LEFT(QtyBox,LEN(QtyBox)-3) ,Maker   " & vbCrLf & _
                  "       ,POQty     " & vbCrLf & _
                  "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
                  "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf

            ls_SQL = ls_SQL + "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              "       FROM (    " & vbCrLf & _
                              " 			SELECT CONVERT(CHAR,row_number() over (order by PMU.PONo)) as NoUrut,PDU.PartNo,PDU.PartNo PartNos,PartName ,PMU.PONo     " & vbCrLf & _
                              "        		,CASE WHEN MPART.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,MU.DESCRIPTION  " & vbCrLf & _
                              "        		,MOQ =CONVERT(CHAR,MOQ),QtyBox = CONVERT(CHAR,QtyBox),ISNULL(MPART.Maker,'')Maker       " & vbCrLf & _
                              " 			,PDU.POQty  " & vbCrLf & _
                              " 			,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    			,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    			,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)                              " & vbCrLf & _
                              "      		,'BEFORE' BYWHAT " & vbCrLf & _
                              "      		,PDU.DeliveryD1 ,PDU.DeliveryD2 ,PDU.DeliveryD3 ,PDU.DeliveryD4 ,PDU.DeliveryD5 ,PDU.DeliveryD6 ,PDU.DeliveryD7 ,PDU.DeliveryD8 ,PDU.DeliveryD9 ,PDU.DeliveryD10  " & vbCrLf

            ls_SQL = ls_SQL + "      		,PDU.DeliveryD11 ,PDU.DeliveryD12 ,PDU.DeliveryD13 ,PDU.DeliveryD14 ,PDU.DeliveryD15 ,PDU.DeliveryD16 ,PDU.DeliveryD17 ,PDU.DeliveryD18 ,PDU.DeliveryD19 ,PDU.DeliveryD20  " & vbCrLf & _
                              "      	,PDU.DeliveryD21 ,PDU.DeliveryD22 ,PDU.DeliveryD23 ,PDU.DeliveryD24 ,PDU.DeliveryD25 ,PDU.DeliveryD26 ,PDU.DeliveryD27 ,PDU.DeliveryD28 ,PDU.DeliveryD29 ,PDU.DeliveryD30 ,PDU.DeliveryD31  " & vbCrLf & _
                              "      	,row_number() over (order by PDU.PONo) as Sort      " & vbCrLf & _
                              "      	FROM dbo.PO_MasterUpload PMU " & vbCrLf & _
                              "  		INNER JOIN dbo.PO_DetailUpload PDU ON PMU.PONo = PDU.PONo  AND PMU.AffiliateID = PDU.AffiliateID AND PMU.SupplierID = PDU.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN PO_Master POM ON PDU.AffiliateID = POM.AffiliateID AND PDU.PONo = POM.PONo AND PDU.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = PDU.PartNo and MP.AffiliateID = PDU.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Parts MPART ON PDU.PartNo = MPART.PartNo         " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Supplier MS ON PDU.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Affiliate MA ON PDU.AffiliateID = MA.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.MS_SupplierCapacity MSC ON PDU.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PDU.SupplierID=MSC.SupplierID          " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls                LEFT JOIN dbo.MS_CurrCls MCUR1 ON PDU.CurrCls = MCUR1.CurrCls          " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf & _
                              "         WHERE  PMU.PONo='" & Trim(txtPONo.Text) & "' AND PMU.SupplierID='" & Trim(txtSupplierCode.Text) & "'  " & vbCrLf & _
                              "            GROUP BY PMU.PONo,PDU.PONo,PDU.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ,QtyBox,PDU.poqty,MPART.Maker,MonthlyProductionCapacity       " & vbCrLf & _
                              "      		,PDU.CurrCls,MCUR1.Description,PDU.Price,PDU.Amount,MP.CurrCls,MCUR2.Description,MP.Price   " & vbCrLf & _
                              "       		,MSC.PartNo    " & vbCrLf & _
                              "      		,PDU.DeliveryD1,PDU.DeliveryD2,PDU.DeliveryD3,PDU.DeliveryD4,PDU.DeliveryD5,PDU.DeliveryD6,PDU.DeliveryD7,PDU.DeliveryD8,PDU.DeliveryD9,PDU.DeliveryD10          " & vbCrLf & _
                              "      		,PDU.DeliveryD11,PDU.DeliveryD12,PDU.DeliveryD13,PDU.DeliveryD14,PDU.DeliveryD15,PDU.DeliveryD16,PDU.DeliveryD17,PDU.DeliveryD18,PDU.DeliveryD19,PDU.DeliveryD20        		    " & vbCrLf & _
                              "      		,PDU.DeliveryD21,PDU.DeliveryD22,PDU.DeliveryD23,PDU.DeliveryD24,PDU.DeliveryD25,PDU.DeliveryD26,PDU.DeliveryD27,PDU.DeliveryD28,PDU.DeliveryD29,PDU.DeliveryD30,PDU.DeliveryD31     " & vbCrLf & _
                              " 	)detail1 " & vbCrLf

            ls_SQL = ls_SQL + " 	UNION ALL  " & vbCrLf & _
                              " 	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
                              " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              "       ,MOQ = MOQ , QtyBox = QtyBox ,Maker   " & vbCrLf & _
                              "       ,POQty     " & vbCrLf & _
                              "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
                              "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf

            ls_SQL = ls_SQL + "       FROM (   " & vbCrLf & _
                              "  		SELECT '' NoUrut,'' PartNo,PORD.PartNo PartNos,''PartName ,'' PONo,'' KanbanCls,'' DESCRIPTION  " & vbCrLf & _
                              "  		,''MOQ,''QtyBox,ISNULL(MPART.Maker,'')Maker       " & vbCrLf & _
                              "         ,PORD.POQty ,'AFTER' BYWHAT " & vbCrLf & _
                              "         ,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    	    ,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)                              " & vbCrLf & _
                              "      	,PORD.DeliveryD1 ,PORD.DeliveryD2 ,PORD.DeliveryD3 ,PORD.DeliveryD4 ,PORD.DeliveryD5 ,PORD.DeliveryD6 ,PORD.DeliveryD7 ,PORD.DeliveryD8 ,PORD.DeliveryD9 ,PORD.DeliveryD10  " & vbCrLf & _
                              "      	,PORD.DeliveryD11 ,PORD.DeliveryD12 ,PORD.DeliveryD13 ,PORD.DeliveryD14 ,PORD.DeliveryD15 ,PORD.DeliveryD16 ,PORD.DeliveryD17 ,PORD.DeliveryD18 ,PORD.DeliveryD19 ,PORD.DeliveryD20  " & vbCrLf & _
                              "      	,PORD.DeliveryD21 ,PORD.DeliveryD22 ,PORD.DeliveryD23 ,PORD.DeliveryD24 ,PORD.DeliveryD25 ,PORD.DeliveryD26 ,PORD.DeliveryD27 ,PORD.DeliveryD28 ,PORD.DeliveryD29 ,PORD.DeliveryD30 ,PORD.DeliveryD31  " & vbCrLf & _
                              "      	,row_number() over (order by PORD.PONo) as Sort      " & vbCrLf

            ls_SQL = ls_SQL + "      	FROM dbo.PORev_Master PORM      " & vbCrLf & _
                              "  		INNER JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = PORD.PartNo and MP.AffiliateID = PORD.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Parts MPART ON PORD.PartNo = MPART.PartNo         " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID          " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls    " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_CurrCls MCUR1 ON PORD.CurrCls = MCUR1.CurrCls  " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf

            ls_SQL = ls_SQL + "         WHERE MONTH(PORM.Period) = MONTH('" & tanggal & "') AND YEAR(PORM.Period) = YEAR('" & tanggal & "')  " & vbCrLf & _
                              "         AND PORM.PORevNo='" & Trim(txtPORev.Text) & "' AND PORM.PONo='" & Trim(txtPONo.Text) & "' AND PORM.SupplierID='" & Trim(txtSupplierCode.Text) & "'  " & vbCrLf & _
                              "         GROUP BY PORD.PONo,PORD.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ,PORM.SeqNo,QtyBox,PORD.poqty,MPART.Maker,MonthlyProductionCapacity       " & vbCrLf & _
                              "      		,PORD.CurrCls,MCUR1.Description,PORD.Price,PORD.Amount,MP.CurrCls,MCUR2.Description,MP.Price   " & vbCrLf & _
                              "       		,PORM.Period,MSC.PartNo    " & vbCrLf & _
                              "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10          " & vbCrLf & _
                              "      		,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20        		    " & vbCrLf & _
                              "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31     " & vbCrLf & _
                              " 	) Detail2 " & vbCrLf & _
                              " 	UNION ALL    " & vbCrLf & _
                              "     SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf

            ls_SQL = ls_SQL + " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              "       ,MOQ = MOQ , QtyBox = QtyBox ,Maker   " & vbCrLf & _
                              "       ,POQty     " & vbCrLf & _
                              "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
                              "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              "       FROM (    " & vbCrLf & _
                              "      		SELECT '' as NoUrut,'' PartNo,PORD.PartNo PartNos,''PartName,''PONo,'' KanbanCls,''Description,MOQ = '',MinOrderQty = MOQ ,PORM.SeqNo    " & vbCrLf

            ls_SQL = ls_SQL + "  			,'' QtyBox,ISNULL(MPART.Maker,'')Maker,'' MonthlyProductionCapacity  " & vbCrLf & _
                              "  			,'SUPPLIER APPROVAL' BYWHAT  " & vbCrLf & _
                              "      		,PORD.POQty  " & vbCrLf & _
                              "      		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    			,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    			,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)     " & vbCrLf & _
                              "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10 " & vbCrLf & _
                              " 			,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20 " & vbCrLf & _
                              "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31 " & vbCrLf & _
                              "      		,row_number() over (order by PORD.PONo) as Sort      " & vbCrLf & _
                              "      	FROM dbo.PORev_Master PORM      " & vbCrLf

            ls_SQL = ls_SQL + "  		INNER JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = PORD.PartNo  and MP.AffiliateID = PORD.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Parts MPART ON PORD.PartNo = MPART.PartNo         " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls     " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_CurrCls MCUR1 ON PORD.CurrCls = MCUR1.CurrCls      " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls     " & vbCrLf

            ls_SQL = ls_SQL + "         WHERE MONTH(PORM.Period) = MONTH('" & tanggal & "') AND YEAR(PORM.Period) = YEAR('" & tanggal & "')  " & vbCrLf & _
                              "         AND PORM.PORevNo='" & Trim(txtPORev.Text) & "' AND PORM.PONo='" & Trim(txtPONo.Text) & "' AND PORM.SupplierID='" & Trim(txtSupplierCode.Text) & "'  " & vbCrLf & _
                              "            GROUP BY PORD.PONo,PORD.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ,PORM.SeqNo,QtyBox,PORD.POQty,MPART.Maker,MonthlyProductionCapacity  " & vbCrLf & _
                              "      		,PORD.CurrCls,MCUR1.Description,PORD.Price,PORD.Amount,MP.CurrCls,MCUR2.Description,MP.Price   " & vbCrLf & _
                              "              ,PORM.Period,MSC.PartNo   " & vbCrLf & _
                              "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10        " & vbCrLf & _
                              "      		,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20          " & vbCrLf & _
                              "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31   " & vbCrLf & _
                              " 		)detail3 " & vbCrLf & _
                              "      	ORDER BY sort, PartNo DESC  " & vbCrLf & _
                              " END   " & vbCrLf

            ls_SQL = ls_SQL + " ELSE   " & vbCrLf & _
                              " BEGIN   " & vbCrLf & _
                              " 	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
                              " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              "       ,MOQ = LEFT(MOQ,LEN(MOQ)-3) , QtyBox = LEFT(QtyBox,LEN(QtyBox)-3) ,Maker   " & vbCrLf & _
                              "       ,POQty     " & vbCrLf & _
                              "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
                              "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf

            ls_SQL = ls_SQL + "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              "       FROM (    " & vbCrLf & _
                              " 			SELECT CONVERT(CHAR,row_number() over (order by PMU.PONo)) as NoUrut,PDU.PartNo,PDU.PartNo PartNos,PartName ,PMU.PONo     " & vbCrLf & _
                              "        		,CASE WHEN MPART.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,MU.DESCRIPTION  " & vbCrLf & _
                              "        		,MOQ =CONVERT(CHAR,MOQ),QtyBox = CONVERT(CHAR,QtyBox),ISNULL(MPART.Maker,'')Maker       " & vbCrLf & _
                              " 			,PDU.POQty  " & vbCrLf & _
                              " 			,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    			,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    			,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)                              " & vbCrLf & _
                              "      		,'BEFORE' BYWHAT " & vbCrLf & _
                              "      		,PDU.DeliveryD1 ,PDU.DeliveryD2 ,PDU.DeliveryD3 ,PDU.DeliveryD4 ,PDU.DeliveryD5 ,PDU.DeliveryD6 ,PDU.DeliveryD7 ,PDU.DeliveryD8 ,PDU.DeliveryD9 ,PDU.DeliveryD10  " & vbCrLf

            ls_SQL = ls_SQL + "      		,PDU.DeliveryD11 ,PDU.DeliveryD12 ,PDU.DeliveryD13 ,PDU.DeliveryD14 ,PDU.DeliveryD15 ,PDU.DeliveryD16 ,PDU.DeliveryD17 ,PDU.DeliveryD18 ,PDU.DeliveryD19 ,PDU.DeliveryD20  " & vbCrLf & _
                              "      	,PDU.DeliveryD21 ,PDU.DeliveryD22 ,PDU.DeliveryD23 ,PDU.DeliveryD24 ,PDU.DeliveryD25 ,PDU.DeliveryD26 ,PDU.DeliveryD27 ,PDU.DeliveryD28 ,PDU.DeliveryD29 ,PDU.DeliveryD30 ,PDU.DeliveryD31  " & vbCrLf & _
                              "      	,row_number() over (order by PDU.PONo) as Sort      " & vbCrLf & _
                              "      	FROM dbo.PO_MasterUpload PMU " & vbCrLf & _
                              "  		INNER JOIN dbo.PO_DetailUpload PDU ON PMU.PONo = PDU.PONo  AND PMU.AffiliateID = PDU.AffiliateID AND PMU.SupplierID = PDU.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN PO_Master POM ON PDU.AffiliateID = POM.AffiliateID AND PDU.PONo = POM.PONo AND PDU.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = PDU.PartNo and MP.AffiliateID = PDU.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Parts MPART ON PDU.PartNo = MPART.PartNo         " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Supplier MS ON PDU.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON PDU.AffiliateID = MA.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.MS_SupplierCapacity MSC ON PDU.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PDU.SupplierID=MSC.SupplierID          " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls                LEFT JOIN dbo.MS_CurrCls MCUR1 ON PDU.CurrCls = MCUR1.CurrCls          " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf & _
                              "         WHERE  PMU.PONo='" & Trim(txtPONo.Text) & "' AND PMU.SupplierID='" & Trim(txtSupplierCode.Text) & "'  " & vbCrLf & _
                              "            GROUP BY PMU.PONo,PDU.PONo,PDU.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ,QtyBox,PDU.poqty,MPART.Maker,MonthlyProductionCapacity       " & vbCrLf & _
                              "      		,PDU.CurrCls,MCUR1.Description,PDU.Price,PDU.Amount,MP.CurrCls,MCUR2.Description,MP.Price   " & vbCrLf & _
                              "       		,MSC.PartNo    " & vbCrLf & _
                              "      		,PDU.DeliveryD1,PDU.DeliveryD2,PDU.DeliveryD3,PDU.DeliveryD4,PDU.DeliveryD5,PDU.DeliveryD6,PDU.DeliveryD7,PDU.DeliveryD8,PDU.DeliveryD9,PDU.DeliveryD10          " & vbCrLf & _
                              "      		,PDU.DeliveryD11,PDU.DeliveryD12,PDU.DeliveryD13,PDU.DeliveryD14,PDU.DeliveryD15,PDU.DeliveryD16,PDU.DeliveryD17,PDU.DeliveryD18,PDU.DeliveryD19,PDU.DeliveryD20        		    " & vbCrLf & _
                              "      		,PDU.DeliveryD21,PDU.DeliveryD22,PDU.DeliveryD23,PDU.DeliveryD24,PDU.DeliveryD25,PDU.DeliveryD26,PDU.DeliveryD27,PDU.DeliveryD28,PDU.DeliveryD29,PDU.DeliveryD30,PDU.DeliveryD31     " & vbCrLf & _
                              " 	)detail1 " & vbCrLf

            ls_SQL = ls_SQL + " 	UNION ALL   " & vbCrLf & _
                              " 	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
                              " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              "       ,MOQ = MOQ , QtyBox = QtyBox,Maker   " & vbCrLf & _
                              "       ,POQty     " & vbCrLf & _
                              "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
                              "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf

            ls_SQL = ls_SQL + "       FROM (   " & vbCrLf & _
                              "       SELECT row_number() over (order by AD.PONo) as Sort ,'' as NoUrut ,'' PartNo ,AD.PartNo AS PartNos,'' PartName ,'' PONo, '' KanbanCls ,''Description ,'' MOQ,'' QtyBox ,AD.Maker   " & vbCrLf & _
                              "         ,POQtyOld POqty " & vbCrLf & _
                              "   		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    		,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)                              " & vbCrLf & _
                              "    		,'AFTER' BYWHAT  " & vbCrLf & _
                              "      	,DeliveryD1Old DeliveryD1,DeliveryD2Old DeliveryD2,DeliveryD3Old DeliveryD3,DeliveryD4Old DeliveryD4,DeliveryD5Old DeliveryD5   " & vbCrLf & _
                              "   		,DeliveryD6Old DeliveryD6,DeliveryD7Old DeliveryD7,DeliveryD8Old DeliveryD8,DeliveryD9Old DeliveryD9,DeliveryD10Old DeliveryD10   " & vbCrLf & _
                              "   		,DeliveryD11Old DeliveryD11,DeliveryD12Old DeliveryD12,DeliveryD13Old DeliveryD13,DeliveryD14Old DeliveryD14,DeliveryD15Old DeliveryD15   " & vbCrLf & _
                              "   		,DeliveryD16Old DeliveryD16,DeliveryD17Old DeliveryD17,DeliveryD18 DeliveryD18,DeliveryD19Old DeliveryD19,DeliveryD20Old DeliveryD20   " & vbCrLf

            ls_SQL = ls_SQL + "   		,DeliveryD21Old DeliveryD21,DeliveryD22Old DeliveryD22,DeliveryD23Old DeliveryD23,DeliveryD24Old DeliveryD24,DeliveryD25Old DeliveryD25   " & vbCrLf & _
                              "   		,DeliveryD26Old DeliveryD26,DeliveryD27Old DeliveryD27,DeliveryD28Old DeliveryD28,DeliveryD29Old DeliveryD29,DeliveryD30Old DeliveryD30,DeliveryD31Old DeliveryD31   " & vbCrLf & _
                              "   		FROM dbo.AffiliateRev_Detail AD   " & vbCrLf & _
                              "   		LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo   " & vbCrLf & _
                              "   		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = AD.PartNo and MP.AffiliateID = AD.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
                              "   		LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "   		LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
                              "   		LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID          " & vbCrLf & _
                              "   		LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls          " & vbCrLf & _
                              "   		LEFT JOIN dbo.MS_CurrCls MCUR1 ON AD.CurrCls = MCUR1.CurrCls          " & vbCrLf & _
                              "   		LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf

            ls_SQL = ls_SQL + "         WHERE AD.PORevNo='" & Trim(txtPORev.Text) & "' AND AD.PONo='" & Trim(txtPONo.Text) & "' AND AD.SupplierID='" & Trim(txtSupplierCode.Text) & "'  " & vbCrLf & _
                              "   		GROUP BY PONo,AD.PartNo,PartName,AD.KanbanCls,POQtyOld,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity   ,SeqNo    " & vbCrLf & _
                              "  		,AD.CurrCls,MCUR1.Description,AD.Price,Amount,MP.CurrCls,MCUR2.Description,MP.Price,MSC.PartNo    " & vbCrLf & _
                              "      	,DeliveryD1Old,DeliveryD2Old,DeliveryD3Old,DeliveryD4Old,DeliveryD5Old   " & vbCrLf & _
                              "   		,DeliveryD6Old,DeliveryD7Old,DeliveryD8Old,DeliveryD9Old,DeliveryD10Old   " & vbCrLf & _
                              "   		,DeliveryD11Old,DeliveryD12Old,DeliveryD13Old,DeliveryD14Old,DeliveryD15Old   " & vbCrLf & _
                              "   		,DeliveryD16Old,DeliveryD17Old,DeliveryD18,DeliveryD19Old,DeliveryD20Old   " & vbCrLf & _
                              "   		,DeliveryD21Old,DeliveryD22Old,DeliveryD23Old,DeliveryD24Old,DeliveryD25Old   " & vbCrLf & _
                              "   		,DeliveryD26Old,DeliveryD27Old,DeliveryD28Old,DeliveryD29Old,DeliveryD30Old,DeliveryD31Old   " & vbCrLf & _
                              "   	 )detail2 " & vbCrLf & _
                              " 	 UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "   	 SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
                              " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              "       ,MOQ = MOQ , QtyBox = QtyBox,Maker   " & vbCrLf & _
                              "       ,POQty     " & vbCrLf & _
                              "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
                              "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              "       FROM (   " & vbCrLf

            ls_SQL = ls_SQL + "       SELECT row_number() over (order by AD.PONo) as Sort ,'' as NoUrut ,'' PartNo ,AD.PartNo AS PartNos,'' PartName ,'' PONo, '' KanbanCls ,''Description ,'' MOQ,'' QtyBox ,AD.Maker   " & vbCrLf & _
                              "         ,POQtyOld POqty " & vbCrLf & _
                              "         ,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    		,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)    " & vbCrLf & _
                              "    		,'SUPPLIER APPROVAL' BYWHAT  " & vbCrLf & _
                              "  		,DeliveryD1 ,DeliveryD2 ,DeliveryD3 ,DeliveryD4 ,DeliveryD5    " & vbCrLf & _
                              "  		,DeliveryD6 ,DeliveryD7 ,DeliveryD8 ,DeliveryD9 ,DeliveryD10    " & vbCrLf & _
                              "  		,DeliveryD11 ,DeliveryD12 ,DeliveryD13 ,DeliveryD14    " & vbCrLf & _
                              "  		,DeliveryD15 ,DeliveryD16 ,DeliveryD17 ,DeliveryD18,DeliveryD19 ,DeliveryD20 ,DeliveryD21    " & vbCrLf & _
                              "  		,DeliveryD22 ,DeliveryD23 ,DeliveryD24 ,DeliveryD25 ,DeliveryD26 ,DeliveryD27 ,DeliveryD28 ,DeliveryD29    " & vbCrLf

            ls_SQL = ls_SQL + "  		,DeliveryD30 ,DeliveryD31    " & vbCrLf & _
                              "  		FROM dbo.AffiliateRev_Detail AD   " & vbCrLf & _
                              "  		 LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo   " & vbCrLf & _
                              "  		 LEFT JOIN dbo.MS_Price MP ON MP.PartNo = AD.PartNo and MP.AffiliateID = AD.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
                              "  		 LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID   " & vbCrLf & _
                              "  		 LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
                              "  		 LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID          " & vbCrLf & _
                              "  		 LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls   " & vbCrLf & _
                              "  		 LEFT JOIN dbo.MS_CurrCls MCUR1 ON AD.CurrCls = MCUR1.CurrCls  " & vbCrLf & _
                              "  		 LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf & _
                              "         WHERE AD.PORevNo='" & Trim(txtPORev.Text) & "' AND AD.PONo='" & Trim(txtPONo.Text) & "' AND AD.SupplierID='" & Trim(txtSupplierCode.Text) & "'  " & vbCrLf

            ls_SQL = ls_SQL + "  		 GROUP BY PONo,AD.PartNo,PartName,AD.KanbanCls,POQty,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity ,POQtyOld,MSC.PartNo    " & vbCrLf & _
                              "      		,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5   " & vbCrLf & _
                              "   			,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10   " & vbCrLf & _
                              "   			,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15   " & vbCrLf & _
                              "   			,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20   " & vbCrLf & _
                              "   			,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25   " & vbCrLf & _
                              "   			,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31  " & vbCrLf & _
                              " 	)detail3 " & vbCrLf & _
                              " 	ORDER BY sort, PartNo DESC  " & vbCrLf & _
                              " END   " & vbCrLf & _
                              "    "



            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            Return ds
        End Using
    End Function

    Private Function Supplier(ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID='" & ls_value & "'"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Function Affiliate(ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID='" & ls_value & "'"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Function EmailToEmailCC() As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            'ls_SQL = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID='" & ls_value & "'"

            ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                     " select 'AFF' flag,affiliatepocc, affiliatepoto='',toEmail='' from ms_emailaffiliate where AffiliateID='" & Trim(txtAffiliateID.Text) & "'" & vbCrLf & _
                     " union all " & vbCrLf & _
                     " --PASI TO -CC " & vbCrLf & _
                     " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI where AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
                     " union all " & vbCrLf & _
                     " --Supplier TO- CC " & vbCrLf & _
                     " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail='' from ms_emailSupplier where SupplierID='" & Trim(txtSupplierCode.Text) & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Sub UpdateExcel(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffCode As String = "", _
                         Optional ByVal pRevNo As String = "", _
                         Optional ByVal pPONo As String = "", _
                         Optional ByVal pSuppCode As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " UPDATE dbo.AffiliateRev_Master " & vbCrLf & _
                          " SET ExcelCls='1'" & vbCrLf & _
                          " WHERE PORevNo = '" & pRevNo & "' " & vbCrLf & _
                          " AND PONo='" & pPONo & "'  " & vbCrLf & _
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

    Private Function getApp(ByVal pPORevNo As String, ByVal pPONo As String) As Boolean
        Dim ls_SQL As String = ""
        Dim doneApp As Boolean = False
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = " SELECT * FROM PORev_Master " & vbCrLf & _
                  " WHERE PORevNo='" & pPORevNo & "' AND PONO='" & pPONo & "' AND  " & vbCrLf & _
                  " (ISNULL(SupplierApproveDate,'') <> '' OR " & vbCrLf & _
                  " ISNULL(SupplierApprovePendingDate,'') <> '' OR  " & vbCrLf & _
                  " ISNULL(SupplierUnApproveDate,'') <> '' OR  " & vbCrLf & _
                  " ISNULL(PASIApproveDate,'') <> '' OR  " & vbCrLf & _
                  " ISNULL(FinalApproveDate,'') <> '') "
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                doneApp = True
            End If
        End Using
        Return doneApp
    End Function
#Region "EXCEL"
    Private Sub Excel()
        On Error GoTo ErrHandler
        Dim strFileSize As String = ""

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        'copy file from server to local
        Dim fileTocopy As String
        Dim NewFileCopy As String

        fileTocopy = Server.MapPath("~/Template/Template PO Revision.xlsm") 'File dari server
        NewFileCopy = "D:/Template/Template PO Revision.xlsm" 'File dari local

        'Copy Tempalte Excel dari Server ke Local untuk diisi
        If System.IO.File.Exists(fileTocopy) = True Then
            System.IO.File.Delete(NewFileCopy)
            System.IO.File.Copy(fileTocopy, NewFileCopy)
        Else
            System.IO.File.Copy(fileTocopy, NewFileCopy)
        End If
        'copy file from server to local

        'For Each fi In aryFi
        Dim xlApp = New Excel.Application
        'Dim ls_file As String = "D:\PASI\Source Code Terakhir\PASISystem\PASISystem\Template\Template PO Revision.xlsm"
        'Dim ls_file As String = Server.MapPath("~\Template\Template PO Revision.xlsm")
        Dim ls_file As String = NewFileCopy
        '
        ExcelBook = xlApp.Workbooks.Open(ls_file)
        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

        Dim ds As New DataSet
        ds = BindDataExcel()
        If ds.Tables(0).Rows.Count > 0 Then

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCC()
            '1 CC Affiliate
            '2 CC PASI
            '3 CC & TO Supplier
            For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                If receiptCCEmail = "" Then
                    receiptCCEmail = dsEmail.Tables(0).Rows(i)("affiliatepocc")
                Else
                    receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("affiliatepocc")
                End If
                receiptEmail = dsEmail.Tables(0).Rows(i)("affiliatepoto")
            Next

            ExcelSheet.Range("H1").Value = "POR"
            ExcelSheet.Range("H2").Value = receiptEmail
            ExcelSheet.Range("H3").Value = Trim(txtAffiliateID.Text)
            'ExcelSheet.Range("H3").Value = Session("AffiliateID")
            'ExcelSheet.Range("H4").Value = Trim(cbolocation.Text)
            ExcelSheet.Range("H5").Value = Trim(txtSupplierCode.Text)

            ExcelSheet.Range("R8").Value = "PO REVISION NO : " & Trim(txtPORev.Text)
            ExcelSheet.Range("I9").Value = Trim(txtPONo.Text)
            ExcelSheet.Range("T9").Value = txtPeriod.Text

            ExcelSheet.Range("Y2").Value = receiptCCEmail

            ExcelSheet.Range("I11").Value = Trim(txtSupplierName.Text)
            Dim dsSupp As New DataSet
            dsSupp = Supplier(Trim(txtSupplierCode.Text))
            ExcelSheet.Range("I12").Value = dsSupp.Tables(0).Rows(0)("Address")

            ExcelSheet.Range("I16").Value = txtAffiliateName.Text
            Dim dsAffp As New DataSet
            dsAffp = Affiliate(Trim(txtAffiliateID.Text))
            ExcelSheet.Range("I17").Value = dsAffp.Tables(0).Rows(0)("Address")

            ExcelSheet.Range("AE12").Value = Trim(txtCommercial.Text)
            ExcelSheet.Range("AE14").Value = Trim(txtShipBy.Text)

            'If rblDelivery.Value = "0" Then
            '    ExcelSheet.Range("AE16").Value = txtAffiliateName.Text
            'ElseIf rblDelivery.Value = "1" Then
            ExcelSheet.Range("AE16").Value = Session("AffiliateID")
            Dim dsAffp2 As New DataSet
            dsAffp2 = Affiliate(Trim(txtAffiliateID.Text))
            'End If
            ExcelSheet.Range("AE17").Value = dsAffp2.Tables(0).Rows(0)("Address")

            'ExcelSheet.Range("AE36").Value = Trim(txttime1.Text)
            'ExcelSheet.Range("AI36").Value = Trim(txttime2.Text)
            'ExcelSheet.Range("AM36").Value = Trim(txttime3.Text)
            'ExcelSheet.Range("AQ36").Value = Trim(txttime4.Text)


            For i = 0 To ds.Tables(0).Rows.Count - 1
                'If ds.Tables(0).Rows(i)("cols") = "1" Then
                'Header
                ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).Merge()
                ExcelSheet.Range("D" & i + 36 & ": H" & i + 36).Merge()
                ExcelSheet.Range("I" & i + 36 & ": P" & i + 36).Merge()
                ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).Merge()
                ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).Merge()
                ExcelSheet.Range("V" & i + 36 & ": W" & i + 36).Merge()
                ExcelSheet.Range("X" & i + 36 & ": Y" & i + 36).Merge()
                ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Merge()
                ExcelSheet.Range("AC" & i + 36 & ": AE" & i + 36).Merge()
                'ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Merge()
                ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Merge()
                ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Merge()
                ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Merge()
                ExcelSheet.Range("AR" & i + 36 & ": AW" & i + 36).Merge()
                ExcelSheet.Range("AX" & i + 36 & ": AY" & i + 36).Merge() '1
                ExcelSheet.Range("AZ" & i + 36 & ": BA" & i + 36).Merge() '2
                ExcelSheet.Range("BB" & i + 36 & ": BC" & i + 36).Merge() '3
                ExcelSheet.Range("BD" & i + 36 & ": BE" & i + 36).Merge() '4
                ExcelSheet.Range("BF" & i + 36 & ": BG" & i + 36).Merge() '5
                ExcelSheet.Range("BH" & i + 36 & ": BI" & i + 36).Merge() '6
                ExcelSheet.Range("BJ" & i + 36 & ": BK" & i + 36).Merge() '7
                ExcelSheet.Range("BL" & i + 36 & ": BM" & i + 36).Merge() '8
                ExcelSheet.Range("BN" & i + 36 & ": BO" & i + 36).Merge() '9
                ExcelSheet.Range("BP" & i + 36 & ": BQ" & i + 36).Merge() '10
                ExcelSheet.Range("BR" & i + 36 & ": BS" & i + 36).Merge() '11
                ExcelSheet.Range("BT" & i + 36 & ": BU" & i + 36).Merge() '12
                ExcelSheet.Range("BV" & i + 36 & ": BW" & i + 36).Merge() '13
                ExcelSheet.Range("BX" & i + 36 & ": BY" & i + 36).Merge() '14
                ExcelSheet.Range("BZ" & i + 36 & ": CA" & i + 36).Merge() '15
                ExcelSheet.Range("CB" & i + 36 & ": CC" & i + 36).Merge() '16
                ExcelSheet.Range("CD" & i + 36 & ": CE" & i + 36).Merge() '17
                ExcelSheet.Range("CF" & i + 36 & ": CG" & i + 36).Merge() '18
                ExcelSheet.Range("CH" & i + 36 & ": CI" & i + 36).Merge() '19
                ExcelSheet.Range("CJ" & i + 36 & ": CK" & i + 36).Merge() '20
                ExcelSheet.Range("CL" & i + 36 & ": CM" & i + 36).Merge() '21
                ExcelSheet.Range("CN" & i + 36 & ": CO" & i + 36).Merge() '22
                ExcelSheet.Range("CP" & i + 36 & ": CQ" & i + 36).Merge() '23
                ExcelSheet.Range("CR" & i + 36 & ": CS" & i + 36).Merge() '24
                ExcelSheet.Range("CT" & i + 36 & ": CU" & i + 36).Merge() '25
                ExcelSheet.Range("CV" & i + 36 & ": CW" & i + 36).Merge() '26
                ExcelSheet.Range("CX" & i + 36 & ": CY" & i + 36).Merge() '27
                ExcelSheet.Range("CZ" & i + 36 & ": DA" & i + 36).Merge() '28
                ExcelSheet.Range("DB" & i + 36 & ": DC" & i + 36).Merge() '29
                ExcelSheet.Range("DD" & i + 36 & ": DE" & i + 36).Merge() '30
                ExcelSheet.Range("DF" & i + 36 & ": DG" & i + 36).Merge() '31

                ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("NoUrut"))
                ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                ExcelSheet.Range("D" & i + 36 & ": H" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("PartNo"))
                ExcelSheet.Range("I" & i + 36 & ": P" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("PartName"))
                ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("POKanbanCls"))
                ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("Description"))
                ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("V" & i + 36 & ": W" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("MOQ"))
                ExcelSheet.Range("X" & i + 36 & ": Y" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("QtyBox"))
                ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("Maker"))
                ExcelSheet.Range("AC" & i + 36 & ": AE" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("POQty"))
                ExcelSheet.Range("AC" & i + 36 & ": DE" & i + 36).NumberFormat = "#,##0"
                'ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("colapprove"))
                ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("ForecastN1"))
                ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).NumberFormat = "#,##0"

                ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("ForecastN2"))
                ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).NumberFormat = "#,##0"

                ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("ForecastN3"))
                ExcelSheet.Range("AO" & i + 36 & ": AN" & i + 36).NumberFormat = "#,##0"

                ExcelSheet.Range("AR" & i + 36 & ": AW" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("BYWHAT"))

                If Trim(ds.Tables(0).Rows(i)("BYWHAT")) = "BEFORE" Then
                    ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Interior.Color = ColorYellow
                ElseIf Trim(ds.Tables(0).Rows(i)("BYWHAT")) = "SUPPLIER APPROVAL" Then
                    ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).Interior.Color = ColorYellow
                End If

                ExcelSheet.Range("AX" & i + 36 & ": AY" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD1")) '1
                ExcelSheet.Range("AZ" & i + 36 & ": BA" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD2")) '2
                ExcelSheet.Range("BB" & i + 36 & ": BC" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD3")) '3
                ExcelSheet.Range("BD" & i + 36 & ": BE" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD4")) '4
                ExcelSheet.Range("BF" & i + 36 & ": BG" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD5")) '5
                ExcelSheet.Range("BH" & i + 36 & ": BI" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD6")) '6
                ExcelSheet.Range("BJ" & i + 36 & ": BK" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD7")) '7
                ExcelSheet.Range("BL" & i + 36 & ": BM" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD8")) '8
                ExcelSheet.Range("BN" & i + 36 & ": BO" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD9")) '9
                ExcelSheet.Range("BP" & i + 36 & ": BQ" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD10")) '10
                ExcelSheet.Range("BR" & i + 36 & ": BS" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD11")) '11
                ExcelSheet.Range("BT" & i + 36 & ": BU" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD12")) '12
                ExcelSheet.Range("BV" & i + 36 & ": BW" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD13")) '13
                ExcelSheet.Range("BX" & i + 36 & ": BY" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD14")) '14
                ExcelSheet.Range("BZ" & i + 36 & ": CA" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD15")) '15
                ExcelSheet.Range("CB" & i + 36 & ": CC" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD16")) '16
                ExcelSheet.Range("CD" & i + 36 & ": CE" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD17")) '17
                ExcelSheet.Range("CF" & i + 36 & ": CG" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD18")) '18
                ExcelSheet.Range("CH" & i + 36 & ": CI" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD19")) '19
                ExcelSheet.Range("CJ" & i + 36 & ": CK" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD20")) '20
                ExcelSheet.Range("CL" & i + 36 & ": CM" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD21")) '21
                ExcelSheet.Range("CN" & i + 36 & ": CO" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD22")) '22
                ExcelSheet.Range("CP" & i + 36 & ": CQ" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD23")) '23
                ExcelSheet.Range("CR" & i + 36 & ": CS" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD24")) '24
                ExcelSheet.Range("CT" & i + 36 & ": CU" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD25")) '25
                ExcelSheet.Range("CV" & i + 36 & ": CW" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD26")) '26
                ExcelSheet.Range("CX" & i + 36 & ": CY" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD27")) '27
                ExcelSheet.Range("CZ" & i + 36 & ": DA" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD28")) '28
                ExcelSheet.Range("DB" & i + 36 & ": DC" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD29")) '29
                ExcelSheet.Range("DD" & i + 36 & ": DE" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD30")) '30
                ExcelSheet.Range("DF" & i + 36 & ": DG" & i + 36).Value = Trim(ds.Tables(0).Rows(i)("DeliveryD31")) '31
                ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).NumberFormat = "#,##0"
                ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                DrawAllBorders(ExcelSheet.Range("B" & i + 36 & ": AE" & i + 36))
                DrawAllBorders(ExcelSheet.Range("AI" & i + 36 & ": DG" & i + 36))
                ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ' End If
            Next
            ExcelSheet.Range("B42").Interior.Color = Color.White
            ExcelSheet.Range("B42").Font.Color = Color.Black
            ExcelSheet.Range("B" & i + 36).Value = "E"
            ExcelSheet.Range("B" & i + 36).Interior.Color = Color.Black
            ExcelSheet.Range("B" & i + 36).Font.Color = Color.White

            ''BORDER
            'ExcelSheet.Range("B36:" & "AW" & i + 36).Select()
            'ExcelSheet.Range("B36:" & "AW" & i + 36).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            'ExcelSheet.Range("B36:" & "AW" & i + 36).Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            'ExcelSheet.Range("B36:" & "AW" & i + 36).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            'ExcelSheet.Range("B36:" & "AW" & i + 36).Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            'ExcelSheet.Range("B36:" & "AW" & i + 39).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            ''BORDER

            'ExcelBook.SaveAs("D:\PASI\Source Code Terakhir\PASISystem\PASISystem\Template\PO Revision.xlsm")

            'Save ke Server
            ExcelBook.SaveAs(Server.MapPath("~\Template\Result\PO Revision.xlsm"))

            'Save ke local
            ExcelBook.SaveAs("D:\Template\Result\PO Revision.xlsm")

            '*****Copy Excel Local ke Server
            'If System.IO.File.Exists(fileTocopy) = True Then
            '    System.IO.File.Delete(fileTocopy)
            '    System.IO.File.Copy(NewFileCopy, fileTocopy)
            'Else
            '    System.IO.File.Copy(NewFileCopy, fileTocopy)
            'End If

            'System.IO.File.Delete(NewFileCopy)
            xlApp.Workbooks.Close()

            xlApp.Quit()
            Call sendEmail()
        End If
        Exit Sub
ErrHandler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
        xlApp.Workbooks.Close()
        xlApp.Quit()
    End Sub

    Private Sub FormatExcel(ByVal pExl As Microsoft.Office.Interop.Excel.Application)

        With pExl
            For iRow = 0 To grid.VisibleRowCount - 1
                For iCol = 2 To 40

                Next
            Next
            'Dim rgAll As Microsoft.Office.Interop.Excel.Range = .Range(.Cells(1, 1), .Cells(Grid.VisibleRowCount + 2, 40))
            'DrawAllBorders(rgAll)

            'Dim rgHeader As Microsoft.Office.Interop.Excel.Range = .Range(.Cells(1, 1), .Cells(2, 40))
            'rgHeader.Interior.Color = ColorOrange
            'rgHeader.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        End With
    End Sub

    Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
        With Rg
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
    End Sub

    Private Sub sendEmail()
        Dim TempFilePath As String
        Dim TempFileName As String
        Dim receiptEmail As String = ""
        Dim receiptCCEmail As String = ""
        Dim fromEmail As String = ""

        'TempFilePath = "D:\PASI\Source Code Terakhir\PASISystem\PASISystem\Template\"
        'TempFilePath = Server.MapPath("~\Template\")
        'receiptEmail = "kristriyana@tos.co.id"
        'receiptEmail = "kris.trieyana@gmail.com"

        '*******File di Server
        TempFilePath = Server.MapPath("~\Template\Result\")
        TempFileName = "PO Revision.xlsm"

        'File di local
        'TempFilePath = "D:\Template\Result\"
        'TempFileName = "PO Revision.xlsm"


        Dim dsEmail As New DataSet
        dsEmail = EmailToEmailCC()
        '1 CC Affiliate
        '2 CC PASI
        '3 CC & TO Supplier
        For i = 0 To dsEmail.Tables(0).Rows.Count - 1
            If receiptCCEmail = "" Then
                receiptCCEmail = dsEmail.Tables(0).Rows(i)("affiliatepocc")
            Else
                receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(i)("affiliatepocc")
            End If
            If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                fromEmail = dsEmail.Tables(0).Rows(i)("toEmail")
            End If
            If receiptEmail = "" Then
                receiptEmail = dsEmail.Tables(0).Rows(i)("affiliatepoto")
            Else
                receiptEmail = receiptEmail & "," & dsEmail.Tables(0).Rows(i)("affiliatepoto")
            End If
            receiptEmail = dsEmail.Tables(0).Rows(i)("affiliatepoto")
        Next
        receiptCCEmail = Replace(receiptCCEmail, ",", ";")
        receiptEmail = Replace(receiptEmail, ",", ";")

        If receiptEmail = "" Then
            MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
            Exit Sub
        End If

        If fromEmail = "" Then
            MsgBox("Mailer's e-mail address is not found", vbCritical, "Warning")
            Exit Sub
        End If

        'Make a copy of the file/Open it/Mail it/Delete it
        'If you want to change the file name then change only TempFileName

        'Dim mailMessage As New Mail.MailMessage(fromEmail, receiptEmail)
        Dim mailMessage As New Mail.MailMessage()
        mailMessage.From = New MailAddress(fromEmail)
        mailMessage.Subject = "PO Revision Template Testing"
        If receiptEmail <> "" Then
            For Each recipient In receiptEmail.Split(";"c)
                If recipient <> "" Then
                    Dim mailAddress As New MailAddress(recipient)
                    mailMessage.To.Add(mailAddress)
                End If
            Next
        End If
        If receiptCCEmail <> "" Then
            For Each recipientCC In receiptCCEmail.Split(";"c)
                If recipientCC <> "" Then
                    Dim mailAddress As New MailAddress(recipientCC)
                    mailMessage.CC.Add(mailAddress)
                End If
            Next
        End If

        mailMessage.Body = "PO Revision Testing"
        Dim filename As String = TempFilePath & TempFileName
        mailMessage.Attachments.Add(New Attachment(filename))
        mailMessage.IsBodyHtml = False
        Dim smtp As New SmtpClient
        smtp.Host = "smtp.atisicloud.com"
        'smtp.Host = "mail.fast.net.id"
        smtp.EnableSsl = False
        smtp.UseDefaultCredentials = True
        smtp.Port = 25
        smtp.Send(mailMessage)

        'Delete the file
        'Kill(TempFilePath & TempFileName)        
    End Sub
#End Region
#End Region

End Class