Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports DevExpress.Web.ASPxMenu
Imports OfficeOpenXml
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.Net

Public Class KanbanCreate
    Inherits System.Web.UI.Page
    Private processAddNewRow As Boolean

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim paramDT1 As Date
    Dim paramDT2 As Date
    Dim paramSupplier As String
    Dim paramLocation As String


    Dim pkanbandate As Date
    Dim psuppID As String
    Dim psuppname As String
    Dim paffentrydate As String
    Dim paffentryname As String
    Dim paffappdate As String
    Dim paffappname As String
    Dim psuppappdate As String
    Dim psuppappname As String
    Dim pdt1 As Date
    Dim pdt2 As Date
    Dim pcbosupplier As String
    Dim pcbosupplierName As String
    Dim pcbolocation As String
    Dim pLocation As String
    Dim pcbolocation1 As String
    Dim pLocation1 As String
    Dim pKanbanno As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "C02"
    Dim xKanbanNo As String = ""

    'session("KNEW") = untuk tau status no kanban dalam tgl yg sama udah pernah dibuat apa belom
    'session("ALREADY") = untuk tau status edit atau new
#End Region

#Region "CONTROL EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)
            'If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
            Session("M01Url") = Request.QueryString("Session")
            Session("MenuDesc") = "KANBAN CREATE"
            'End Ifs

            '================ init ===================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                    Session("MenuDesc") = "KANBAN CREATE"
                End If
                If Not IsNothing(Request.QueryString("prm")) Then
                    Dim param As String = Request.QueryString("prm").ToString

                    If param = "  'back'" Then
                        btnsubmenu.Text = "BACK"
                        dtkanban.Value = Now
                        dt1.Value = Now

                        'coba
                        dt1.Value = Format(Now, "MMM yyyy")
                        dtkanban.Value = Now
                        Call up_fillcombo()
                        Call fillHeader()
                        Call up_GridLoad(dtkanban.Value, cbosupplier.Text, cboseqno.Text)
                        dt1.Value = Now
                        'coba
                    Else
                        Session.Remove("FilterKanbanNo")

                        ' update coba coba
                        Session.Remove("xparam_kanbandate")
                        Session.Remove("xparam_suppID")
                        Session.Remove("xparam_suppname")
                        Session.Remove("xparam_affentrydate")
                        Session.Remove("xparam_affentryname")
                        Session.Remove("xparam_affappdate")
                        Session.Remove("xparam_affappname")
                        Session.Remove("xparam_suppappdate")
                        Session.Remove("xparam_suppappname")
                        Session.Remove("xparam_dt1")
                        Session.Remove("xparam_dt2")
                        Session.Remove("xparam_cbosupplier")
                        Session.Remove("xparam_cbosupplierName")
                        Session.Remove("xparam_cbolocation")
                        Session.Remove("xparam_Location")
                        Session.Remove("xparam_cbolocation1")
                        Session.Remove("xparam_Location1")
                        Session.Remove("xparam_Kanbanno")
                        'end
                        Session("MenuDesc") = "KANBAN CREATE"
                        pkanbandate = Split(param, "|")(0)
                        psuppID = Split(param, "|")(1)
                        psuppname = Split(param, "|")(2)
                        paffentrydate = Split(param, "|")(3)
                        paffentryname = Split(param, "|")(4)
                        paffappdate = Split(param, "|")(5)
                        paffappname = Split(param, "|")(6)
                        psuppappdate = Split(param, "|")(7)
                        psuppappname = Split(param, "|")(8)
                        pdt1 = Split(param, "|")(9)
                        pdt2 = Split(param, "|")(10)
                        pcbosupplier = Split(param, "|")(11)
                        pcbosupplierName = Split(param, "|")(12)
                        pcbolocation = Split(param, "|")(13)
                        pLocation = Split(param, "|")(14)
                        pcbolocation1 = Split(param, "|")(15)
                        pLocation1 = Split(param, "|")(16)
                        pKanbanno = Split(param, "|")(17)
                        xKanbanNo = pKanbanno
                        Session("xKanbanNo") = xKanbanNo
                        Session("xparam_kanbandate") = pkanbandate
                        Session("xparam_suppID") = psuppID
                        Session("xparam_suppname") = psuppname
                        Session("xparam_affentrydate") = paffentrydate
                        Session("xparam_affentryname") = paffentryname
                        Session("xparam_affappdate") = paffappdate
                        Session("xparam_affappname") = paffappname
                        Session("xparam_suppappdate") = psuppappdate
                        Session("xparam_suppappname") = psuppappname
                        Session("xparam_dt1") = pdt1
                        Session("xparam_dt2") = pdt2
                        Session("xparam_cbosupplier") = pcbosupplier
                        Session("xparam_cbosupplierName") = pcbosupplierName
                        Session("xparam_cbolocation") = pcbolocation
                        Session("xparam_Location") = pLocation
                        Session("xparam_cbolocation1") = pcbolocation1
                        Session("xparam_Location1") = pLocation1
                        Session("xparam_Kanbanno") = pKanbanno

                        If psuppID <> "" Then btnsubmenu.Text = "BACK"
                        cbosupplier.Text = psuppID
                        txtsuppliername.Text = psuppname
                        cbolocation.Text = pcbolocation
                        txtlocation.Text = pLocation
                        dtkanban.Value = pkanbandate
                        If paffappdate <> "" Then txtaffiliateappdate.Text = paffappdate
                        txtaffiliateappname.Text = paffappname
                        txtaffiliateentrydate.Text = paffentrydate
                        txtaffiliateentryname.Text = paffentryname
                        'txtsupplierapprovaldate.Text = psuppappdate
                        'txtsupplierapprovalname.Text = psuppappname
                        'dt1.Value = Format(pkanbandate, "MMM yyyy")
                        Session("cycle") = "1-4"
                        Session("FilterKanbanNo") = pKanbanno

                        Session("xparam_kanbandate") = pkanbandate
                        Session("xparam_suppID") = psuppID
                        Session("xparam_suppname") = psuppname
                        Session("xparam_affentrydate") = paffentrydate
                        Session("xparam_affentryname") = paffentryname
                        Session("xparam_affappdate") = paffappdate
                        Session("xparam_affappname") = paffappname
                        Session("xparam_suppappdate") = psuppappdate
                        Session("xparam_suppappname") = psuppappname
                        Session("xparam_dt1") = pdt1
                        Session("xparam_dt2") = pdt2
                        Session("xparam_cbosupplier") = pcbosupplier
                        Session("xparam_cbosupplierName") = pcbosupplierName
                        Session("xparam_cbolocation") = pcbolocation
                        Session("xparam_Location") = pLocation
                        Session("xparam_cbolocation1") = pcbolocation1
                        Session("xparam_Location1") = pLocation1
                        Session("xparam_Kanbanno") = pKanbanno


                        Call fillHeader()
                        Call up_GridLoad(pkanbandate, psuppID, 1)

                    End If

                    btnsubmenu.Text = "BACK"
                Else
                    If Not IsNothing(Session("K-Param")) Then
                        btnsubmenu.Text = "BACK"
                    End If
                    dt1.Value = Format(Now, "MMM yyyy")
                    dtkanban.Value = Now
                    If Session("KCR-ReportCode") <> "" Then
                        dtkanban.Text = Session("KCR-KanbanDate")
                        cbosupplier.Value = Session("KCR-SupplierCode")
                        cbolocation.Value = Replace(Session("KCR-DeliveryLocation"), "'", "")
                        Call fillHeader()
                        Call up_GridLoad(Session("KCR-KanbanDate"), Session("KCR-SupplierCode"), 1)
                        btnsubmenu.Text = Session("btn")

                        Session.Remove("YA030ReqNo")
                        Session.Remove("KCR-KanbanNo")
                        Session.Remove("KCR-ReportCode")
                        Session.Remove("KCR-KanbanDate")
                        Session.Remove("KCR-SupplierCode")
                        Session.Remove("KCR-DeliveryLocation")
                        Session.Remove("KCR-Form")
                        Session.Remove("btn")
                    Else
                        Call up_GridLoad(dtkanban.Value, Trim(cbosupplier.Text), 1)
                    End If
                    dt1.Value = Now
                End If
            End If
            '================ init ===================


            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_fillcombo()
                'Call up_GridLoad(dtkanban.Value, Trim(cbosupplier.Text))
                'Clear()
                lblerrmessage.Text = ""
                'dt1.Value = Format(dtkanban.Value, "MMM yyyy")
            End If

            Call colorGrid()
            'If ls_AllowDelete = False Then btndelete.Enabled = False
            'If ls_AllowUpdate = False Then btnsubmit.Enabled = False
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Session.Remove("xKanbanNo")
        If btnsubmenu.Text = "BACK" Then

            Dim param As String
            If Not IsNothing(Session("K-Param")) Then
                param = Session("K-Param")
            Else
                param = Request.QueryString("prm").ToString
            End If

            If param <> "  'back'" Then
                Dim pdt1 As Date = Split(param, "|")(9)
                Dim pdt2 As Date = Split(param, "|")(10)
                Dim pcbosupplier As String = Split(param, "|")(11)
                Dim psuppname As String = Split(param, "|")(12)
                Dim pcbolocation As String = Split(param, "|")(15)
                Dim plocationname As String = Split(param, "|")(16)
                Dim pKanbanno As String = Split(param, "|")(17)

                Response.Redirect("~/Kanban/KanbanList.aspx?prm=" + pdt1 + "|" + pdt2 + "|" + pcbosupplier + "|" + psuppname + "|" + pcbolocation + "|" + plocationname + "|" + pKanbanno + "")
            Else
                Response.Redirect("~/Kanban/KanbanList.aspx")
            End If
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Public Sub btnclear_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnclear.Click
        Clear()
        up_GridLoad(dtkanban.Value, "xxx", "")
    End Sub

    'Private Sub Grid_InitNewRow(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataInitNewRowEventArgs) Handles Grid.InitNewRow
    '    processAddNewRow = True
    '    Grid.FilterExpression = "false"

    '    Dim menu As ASPxMenu = CType(Grid.FindEditFormTemplateControl("editFormMenu"), ASPxMenu)
    '    'menu.Items.FindByName("idx").ClientEnabled = False
    '    menu.Items.FindByName("cols").ClientEnabled = False
    'End Sub

    Private Sub Grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles Grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 1
        Dim isStatusNew As Boolean
        Dim pIsUpdate As Boolean
        Dim sqlstring As String
        Dim i As Long = 0
        Dim ls_kanbanNo As String
        Dim ls_CycleQty As Double
        Dim ls_kanbanTime As String
        Dim ls_seq As Integer

        isStatusNew = False
        ls_kanbanNo = ""
        ls_CycleQty = 0

        If cboseq.Text = "1-4" Then
            ls_seq = "1"
        ElseIf cboseq.Text = "5-8" Then
            ls_seq = "2"
        ElseIf cboseq.Text = "9-12" Then
            ls_seq = "3"
        ElseIf cboseq.Text = "13-16" Then
            ls_seq = "4"
        ElseIf cboseq.Text = "17-20" Then
            ls_seq = "5"
        End If

        If cbosupplier.Text = "" And cbolocation.Text = "" Then
            Call clsMsg.DisplayMessage(lblerrmessage, "6010", clsMessage.MsgType.ErrorMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            Session("msgE02") = lblerrmessage.Text
            lblerrmessage.Text = Session("msgE02")
            Exit Sub
        End If

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            'cek data udah approve atau belum
            ls_SQL = "SELECT isnull(AffiliateApproveUser,'') as AffiliateApproveUser FROM dbo.kanban_master WHERE Kanbandate ='" & Format(dtkanban.Value, "yyyy-MM-dd") & "'" & vbCrLf & _
                    " AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                    " AND SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                    " AND isnull(affiliateApproveUser,'') <> '' " & vbCrLf & _
                    " and deliveryLocationcode = '" & Trim(cbolocation.Text) & "'" & vbCrLf & _
                    " and KanbanSeq_No = '" & Trim(ls_seq) & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(i)("AffiliateApproveUser") <> "" Then
                    Call clsMsg.DisplayMessage(lblerrmessage, "6027", clsMessage.MsgType.ErrorMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Session("msgE02") = lblerrmessage.Text
                    lblerrmessage.Text = Session("msgE02")
                    Exit Sub
                End If
            End If
            'cek data udah approve atau belum

            Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")
                Dim sqlComm As New SqlCommand(ls_SQL, cn, sqlTran)


                With Grid
                    For iLoop = 0 To e.UpdateValues.Count - 1
                        ls_Active = (e.UpdateValues(iLoop).NewValues("cols").ToString())
                        If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

                        'cek kanbanqty ga boleh lebih besar dari poremaining
                        If CDbl(e.UpdateValues(iLoop).NewValues("colkanbanqty").ToString()) > CDbl(e.UpdateValues(iLoop).NewValues("colremainingpo").ToString()) Then
                            sqlTran.Commit()
                            Call clsMsg.DisplayMessage(lblerrmessage, "6031", clsMessage.MsgType.ErrorMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("msgE02") = lblerrmessage.Text
                            Exit Sub
                        End If
                        'cek kanbanqty ga boleh lebih besar dari poremaining

                        For i = 1 To 4

                            If i = 1 Then
                                If txtkanban1.Text = "" Then
                                    ls_kanbanNo = Trim(e.UpdateValues(iLoop).NewValues("kanbanno1").ToString())
                                    ls_kanbanTime = Trim(e.UpdateValues(iLoop).NewValues("kanbantime1").ToString())
                                Else
                                    ls_kanbanNo = Trim(txtkanban1.Text)
                                    ls_kanbanTime = Trim(txttime1.Text)
                                End If
                                ls_CycleQty = CDbl(e.UpdateValues(iLoop).NewValues("colcycle1").ToString())
                            ElseIf i = 2 Then
                                If txtkanban2.Text = "" Then
                                    ls_kanbanNo = Trim(e.UpdateValues(iLoop).NewValues("kanbanno2").ToString())
                                    ls_kanbanTime = Trim(e.UpdateValues(iLoop).NewValues("kanbantime2").ToString())
                                Else
                                    ls_kanbanNo = Trim(txtkanban2.Text)
                                    ls_kanbanTime = Trim(txttime2.Text)
                                End If
                                ls_CycleQty = CDbl(e.UpdateValues(iLoop).NewValues("colcycle2").ToString())
                            ElseIf i = 3 Then
                                If txtkanban3.Text = "" Then
                                    ls_kanbanNo = Trim(e.UpdateValues(iLoop).NewValues("kanbanno3").ToString())
                                    ls_kanbanTime = Trim(e.UpdateValues(iLoop).NewValues("kanbantime3").ToString())
                                Else
                                    ls_kanbanNo = Trim(txtkanban3.Text)
                                    ls_kanbanTime = Trim(txttime3.Text)
                                End If
                                ls_CycleQty = CDbl(e.UpdateValues(iLoop).NewValues("colcycle3").ToString())
                            ElseIf i = 4 Then
                                If txtkanban4.Text = "" Then
                                    ls_kanbanNo = Trim(e.UpdateValues(iLoop).NewValues("kanbanno4").ToString())
                                    ls_kanbanTime = Trim(e.UpdateValues(iLoop).NewValues("kanbantime4").ToString())
                                Else
                                    ls_kanbanNo = Trim(txtkanban4.Text)
                                    ls_kanbanTime = Trim(txttime4.Text)
                                End If
                                ls_CycleQty = CDbl(e.UpdateValues(iLoop).NewValues("colcycle4").ToString())

                            End If

                            'insert master
                            sqlstring = "SELECT * FROM dbo.kanban_master WHERE KanbanNo ='" & Trim(ls_kanbanNo) & "'" & vbCrLf & _
                                        " AND AffiliateID = '" & Session("affiliateid") & "' " & vbCrLf & _
                                        " AND SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                                        " and deliverylocationcode = '" & Trim(cbolocation.Text) & "'" & vbCrLf & _
                                        " and KanbanSeq_No = '" & Trim(ls_seq) & "'"

                            sqlComm = New SqlCommand(sqlstring, cn, sqlTran)
                            Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

                            If sqlRdrM.Read And ls_Active = "1" Then
                                'UPDATE
                                ls_SQL = " UPDATE dbo.Kanban_Master SET " & vbCrLf & _
                                         " KanbanTime = 'ISNULL(" & ls_kanbanTime & ",'00:00:00')', UpdateDate = GETDATE() " & vbCrLf & _
                                         " , DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                                         " WHERE KanbanNo ='" & Trim(ls_kanbanNo) & "'" & vbCrLf & _
                                         " AND AffiliateID = '" & Session("affiliateid") & "' " & vbCrLf & _
                                         " AND SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                                         " AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                                         " and KanbanSeq_No = '" & Trim(ls_seq) & "'"
                            ElseIf Not sqlRdrM.Read And ls_Active = "1" Then
                                'INSERT
                                ls_SQL = " INSERT INTO dbo.Kanban_Master " & vbCrLf & _
                                         "         ( KanbanNo , " & vbCrLf & _
                                         "           AffiliateID , " & vbCrLf & _
                                         "           SupplierID , " & vbCrLf & _
                                         "           DeliveryLocationCode, " & vbCrLf & _
                                         "           KanbanCycle , " & vbCrLf & _
                                         "           KanbanDate , " & vbCrLf & _
                                         "           KanbanTime , " & vbCrLf & _
                                         "           KanbanStatus , " & vbCrLf & _
                                         "           AffiliateApproveUser , " & vbCrLf & _
                                         "           AffiliateApproveDate , " & vbCrLf & _
                                         "           SupplierApproveUser , " & vbCrLf

                                ls_SQL = ls_SQL + "           SupplierApproveDate , " & vbCrLf & _
                                                  "           EntryDate , " & vbCrLf & _
                                                  "           EntryUser , " & vbCrLf & _
                                                  "           UpdateDate , " & vbCrLf & _
                                                  "           UpdateUser, " & vbCrLf & _
                                                  "           KanbanSeq_No " & vbCrLf & _
                                                  "         ) " & vbCrLf & _
                                                  " VALUES  ( '" & ls_kanbanNo & "' , -- KanbanNo - char(20) " & vbCrLf & _
                                                  "           '" & Session("affiliateid").ToString & "' , -- AffiliateID - char(10) " & vbCrLf & _
                                                  "           '" & Trim(cbosupplier.Text) & "' , -- SupplierID - char(10) " & vbCrLf & _
                                                  "           '" & Trim(cbolocation.Text) & "', " & vbCrLf & _
                                                  "           '" & i & "' , -- KanbanCycle - char(1) " & vbCrLf & _
                                                  "           '" & dtkanban.Value & "' , -- KanbanDate - date " & vbCrLf

                                ls_SQL = ls_SQL + "           'ISNULL(" & ls_kanbanTime & ",'00:00:00')' , -- KanbanTime - time " & vbCrLf & _
                                                  "           '' , -- KanbanStatus - char(1) " & vbCrLf & _
                                                  "           '' , -- AffiliateApproveUser - char(15) " & vbCrLf & _
                                                  "           '' , -- AffiliateApproveDate - date " & vbCrLf & _
                                                  "           '' , -- SupplierApproveUser - char(15) " & vbCrLf & _
                                                  "           '' , -- SupplierApproveDate - date " & vbCrLf & _
                                                  "           getdate() , -- EntryDate - datetime " & vbCrLf & _
                                                  "           '" & Session("UserID").ToString & "' , -- EntryUser - char(15) " & vbCrLf & _
                                                  "           getdate() , -- UpdateDate - datetime " & vbCrLf & _
                                                  "           '" & Session("UserID").ToString & "',  -- UpdateUser - char(15) " & vbCrLf & _
                                                  "           " & ls_seq & "' " & vbCrLf & _
                                                  "         ) " & vbCrLf
                                'ElseIf sqlRdrM.Read = False And ls_Active = "0" Then
                                '    'Delete Data
                                '    ls_SQL = " DELETE From Kanban_Master  " & vbCrLf & _
                                '             " WHERE KanbanNo ='" & Trim(ls_kanbanNo) & "'" & vbCrLf & _
                                '             " AND SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf
                            End If
                            sqlRdrM.Close()
                            If ls_Active = "1" Then
                                sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                Session("Kstatus") = "TRUE"
                            End If
                            sqlRdrM.Close()
                            'insert master

                            sqlstring = "SELECT * FROM dbo.Kanban_Detail WHERE KanbanNo ='" & Trim(ls_kanbanNo) & "'" & vbCrLf & _
                                        " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' " & vbCrLf & _
                                        " AND poNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' " & vbCrLf & _
                                        " AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                                        " AND SupplierID = '" & Trim(cbosupplier.Text) & "'" & vbCrLf & _
                                        " and DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                                        " and AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf


                            sqlComm = New SqlCommand(sqlstring, cn, sqlTran)
                            Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                            If sqlRdr.Read Then
                                pIsUpdate = True
                            Else
                                pIsUpdate = False
                            End If
                            sqlRdr.Close()

                            If pIsUpdate = False And ls_Active = "1" Then 'And CDbl(ls_CycleQty) <> 0 Then
                                ls_SQL = ""
                                'INSERT KANBAN
                                ls_SQL = " INSERT INTO dbo.Kanban_Detail " & vbCrLf & _
                                         "         ( KanbanNo , " & vbCrLf & _
                                         "           AffiliateID , " & vbCrLf & _
                                         "           SupplierID , " & vbCrLf & _
                                         "           PartNo , " & vbCrLf & _
                                         "           PONo , " & vbCrLf & _
                                         "           UnitCls , " & vbCrLf & _
                                         "           KanbanQty, DeliveryLocationCode, " & vbCrLf & _
										 "           POMOQ, POQtyBox " & vbCrLf & _
                                         "         ) " & vbCrLf & _
                                         " VALUES  ( '" & Trim(ls_kanbanNo) & "' , -- KanbanNo - char(20) " & vbCrLf & _
                                         "           '" & Session("affiliateid").ToString & " ' , -- AffiliateID - char(10) " & vbCrLf

                                ls_SQL = ls_SQL + "           '" & Trim(cbosupplier.Text) & "' , -- SupplierID - char(10) " & vbCrLf & _
                                                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' , -- PartNo - char(25) " & vbCrLf & _
                                                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' , -- PONo - char(20) " & vbCrLf & _
                                                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("coluomcode").ToString()) & "' , -- UnitCls - char(2) " & vbCrLf & _
                                                  "           " & CDbl(ls_CycleQty) & ",'" & Trim(cbolocation.Text) & "',  -- KanbanQty - numeric " & vbCrLf & _
                                                  "           '" & uf_GetMOQ(Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()), Trim(cbosupplier.Text), Session("AffiliateID")) & "',  -- MOQ - numeric " & vbCrLf & _
                                                  "           '" & uf_GetQtybox(Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()), Trim(cbosupplier.Text), Session("AffiliateID")) & "'  -- QtyBox - numeric " & vbCrLf & _
                                                  "         ) " & vbCrLf

                            ElseIf pIsUpdate = True And ls_Active = "1" Then 'And CDbl(ls_CycleQty) <> 0 Then
                                'Update Data
                                ls_SQL = " Update kanban_Detail set " & vbCrLf & _
                                         " KanbanQty = '" & CDbl(ls_CycleQty) & "' " & vbCrLf & _
                                         " WHERE KanbanNo ='" & Trim(ls_kanbanNo) & "'" & vbCrLf & _
                                         " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' " & vbCrLf & _
                                         " AND poNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' " & vbCrLf & _
                                         " AND SupplierId = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                                         " and AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                                         " and DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' "

                            ElseIf pIsUpdate = True And ls_Active = "0" Then
                                'Delete Data
                                ls_SQL = " DELETE From Kanban_Detail  " & vbCrLf & _
                                         " WHERE KanbanNo ='" & Trim(ls_kanbanNo) & "'" & vbCrLf & _
                                         " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' " & vbCrLf & _
                                         " AND poNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' " & vbCrLf & _
                                         " AND SupplierId = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                                         " and AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                                         " and DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' "
                            End If

                            sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("msgE02") = lblerrmessage.Text
                            Session("Kstatus") = "TRUE"
                        Next i
                    Next iLoop
                End With

                sqlComm.Dispose()
                sqlTran.Commit()

            End Using

            cn.Close()
        End Using
        Call colorGrid()

    End Sub

    Private Sub Grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles Grid.CellEditorInitialize

        If (e.Column.FieldName = "cols" Or e.Column.FieldName = "colkanbanqty" Or e.Column.FieldName = "colcycle1" Or e.Column.FieldName = "colcycle2" Or e.Column.FieldName = "colcycle3" Or e.Column.FieldName = "colcycle4") Then
            e.Editor.ReadOnly = False
        Else
            e.Editor.ReadOnly = True
        End If
    End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Dim pDate As Date = Split(e.Parameters, "|")(1)
        Dim Psupplier As String = Split(e.Parameters, "|")(2)
        Dim pSeq As String = Split(e.Parameters, "|")(3)

        If Format(dt1.Value, "MMM yyyy") <> Format(dtkanban.Value, "MMM yyyy") Then
            Call clsMsg.DisplayMessage(lblerrmessage, "6030", clsMessage.MsgType.ErrorMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            Exit Sub
        End If

        Try
            Select Case pAction

                Case "gridload"
                    Call fillHeader()
                    Call up_GridLoad(pDate, Psupplier, pSeq)
                    If cbosupplier.Text <> "" And cbolocation.Text <> "" Then

                        If pAction = "gridload" Then
                            If Grid.VisibleRowCount = 0 Then
                                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Else
                                Grid.JSProperties("cpMessage") = ""
                                lblerrmessage.Text = ""
                            End If
                        End If
                        Call colorGrid()
                    End If
                Case "approve"
                    Grid.JSProperties("cpMessage") = Session("msgapprove")
                    lblerrmessage.Text = Session("msgapprove")
                    Session.Remove("msgapprove")

                Case "save"

                    xKanbanNo = txtkanban1.Text
                    Call fillHeader()
                    Call up_GridLoad(pDate, Psupplier, pSeq)

                    If Session("Kstatus") = "TRUE" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        lblerrmessage.Text = lblerrmessage.Text
                    ElseIf Session("Kstatus") = "FALSE" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "6017", clsMessage.MsgType.ErrorMessage)
                        Grid.JSProperties("cpMessage") = ""
                        lblerrmessage.Text = lblerrmessage.Text
                    End If
                    Session.Remove("Kstatus")
                    If Not IsNothing(Session("msgE02")) Then
                        lblerrmessage.Text = Session("msgE02")
                        Grid.JSProperties("cpMessage") = Session("msgE02")
                    End If
                Case "change"
                    Call fillHeader()
                    up_GridLoadWhenEventChange()
                    lblerrmessage.Text = ""
                    Grid.JSProperties("cpMessage") = ""
                    Grid.JSProperties("cpPeriod") = Format(dtkanban.Value, "MMM yyyy")
                    dt1.Text = Grid.JSProperties("cpPeriod")
            End Select
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            Grid.FocusedRowIndex = -1

        Finally
            'If (Not IsNothing(Session("YA010Msg"))) Then Grid.JSProperties("cpMessage") = Session("YA010Msg") : Session.Remove("YA010Msg")
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
        End Try
    End Sub

    Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
        Dim x As Integer = CInt(e.VisibleIndex.ToString())



        Dim pDeliveryQty As Double
        Dim pSupplierCapacity As Double

        If x > Grid.VisibleRowCount Then Exit Sub
        'e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        If e.DataColumn.FieldName = "colkanbanqty" Then
            pDeliveryQty = CDbl(e.GetValue("colkanbanqty"))
            pSupplierCapacity = CDbl(e.GetValue("colremainingsupplier"))
        End If

        With Grid
            If .VisibleRowCount > 0 Then
                If pDeliveryQty > pSupplierCapacity Then
                    If e.DataColumn.FieldName = "colkanbanqty" Then
                        e.Cell.BackColor = Color.HotPink
                    End If
                End If

            End If
        End With

    End Sub

    Protected Sub btnprintcard_Click(sender As Object, e As EventArgs) Handles btnprintcard.Click
        Session.Remove("YA030ReqNo")
        Session.Remove("KCR-KanbanNo")
        Session.Remove("KCR-ReportCode")
        Session.Remove("KCR-KanbanDate")
        Session.Remove("KCR-SupplierCode")
        Session.Remove("KCR-DeliveryLocation")
        Session.Remove("KCR-Form")
        Session.Remove("K-Param")
        Session.Remove("btn")
        Session.Remove("KCR-kanbanno")

        Session.Remove("tmp_kanbandate")
        Session.Remove("tmp_suppID")
        Session.Remove("tmp_suppname")
        Session.Remove("tmp_affentrydate")
        Session.Remove("tmp_affentryname")
        Session.Remove("tmp_affappdate")
        Session.Remove("tmp_affappname")
        Session.Remove("tmp_suppappdate")
        Session.Remove("tmp_suppappname")
        Session.Remove("tmp_dt1")
        Session.Remove("tmp_dt2")
        Session.Remove("tmp_cbosupplier")
        Session.Remove("tmp_cbosupplierName")
        Session.Remove("tmp_cbolocation")
        Session.Remove("tmp_Location")
        Session.Remove("tmp_cbolocation1")
        Session.Remove("tmp_Location1")
        Session.Remove("tmp_Kanbanno")


        If txtaffiliateappdate.Text <> "" Then
            Session("KCR-KanbanNo") = "'" + Trim(txtkanban1.Text) + "'"
            Session("KCR-KanbanDate") = Format(dtkanban.Value, "dd MMM yyyy")
            Session("KCR-SupplierCode") = Trim(cbosupplier.Text)
            Session("KCR-DeliveryLocation") = "" & Trim(cbolocation.Text) & ""
            Session("KCR-ReportCode") = "KanbanCard"
            Session("KCR-Form") = "KanbanCreate"
            Session("btn") = btnsubmenu.Text
            Session("KCR-kanbanno") = "'" + Session("FilterKanbanNo") + "'"

            Session("tmp_kanbandate") = Session("xparam_kanbandate")
            Session("tmp_suppID") = Session("xparam_suppID")
            Session("tmp_suppname") = Session("xparam_suppname")
            Session("tmp_affentrydate") = Session("xparam_affentrydate")
            Session("tmp_affentryname") = Session("xparam_affentryname")
            Session("tmp_affappdate") = Session("xparam_affappdate")
            Session("tmp_affappname") = Session("xparam_affappname")
            Session("tmp_suppappdate") = Session("xparam_suppappdate")
            Session("tmp_suppappname") = Session("xparam_suppappname")
            Session("tmp_dt1") = Session("xparam_dt1")
            Session("tmp_dt2") = Session("xparam_dt2")
            Session("tmp_cbosupplier") = Session("xparam_cbosupplier")
            Session("tmp_cbosupplierName") = Session("xparam_cbosupplierName")
            Session("tmp_cbolocation") = Session("xparam_cbolocation")
            Session("tmp_Location") = Session("xparam_Location")
            Session("tmp_cbolocation1") = Session("xparam_cbolocation1")
            Session("tmp_Location1") = Session("xparam_Location1")
            Session("tmp_Kanbanno") = Session("xparam_Kanbanno")


            If Not IsNothing(Request.QueryString("prm")) Then
                Session("K-Param") = Session("xparam_kanbandate") + "|" + Session("xparam_suppID") + "|" + Session("xparam_suppname") + "|" + Session("xparam_affentrydate") + "|" + Session("xparam_affentryname") + "|" + Session("xparam_affappdate") + "|" + Session("xparam_affappname") + "|" + Session("xparam_suppappdate") + "|" + Session("xparam_suppappname") + "|" + Session("xparam_dt1") + "|" + Session("xparam_dt2") + "|" + Session("xparam_cbosupplier") + "|" + Session("xparam_cbosupplierName") + "|" + Session("xparam_cbolocation") + "|" + Session("xparam_Location") + "|" + Session("xparam_cbolocation1") + "|" + Session("xparam_Location1") + "|" + Session("xparam_Kanbanno")
            End If
            Response.Redirect("~/Kanban/ViewReport.aspx")
        Else
            Call clsMsg.DisplayMessage(lblerrmessage, "7011", clsMessage.MsgType.ErrorMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            fillHeader()
        End If

    End Sub

    Protected Sub btnprintcycle_Click(sender As Object, e As System.EventArgs) Handles btnprintcycle.Click
        Session.Remove("YA030ReqNo")
        Session.Remove("KCR-KanbanNo")
        Session.Remove("KCR-ReportCode")
        Session.Remove("KCR-KanbanDate")
        Session.Remove("KCR-SupplierCode")
        Session.Remove("KCR-DeliveryLocation")
        Session.Remove("KCR-Form")
        Session.Remove("K-Param")
        Session.Remove("btn")
        Session.Remove("KCR-kanbanno")

        Session.Remove("tmp_kanbandate")
        Session.Remove("tmp_suppID")
        Session.Remove("tmp_suppname")
        Session.Remove("tmp_affentrydate")
        Session.Remove("tmp_affentryname")
        Session.Remove("tmp_affappdate")
        Session.Remove("tmp_affappname")
        Session.Remove("tmp_suppappdate")
        Session.Remove("tmp_suppappname")
        Session.Remove("tmp_dt1")
        Session.Remove("tmp_dt2")
        Session.Remove("tmp_cbosupplier")
        Session.Remove("tmp_cbosupplierName")
        Session.Remove("tmp_cbolocation")
        Session.Remove("tmp_Location")
        Session.Remove("tmp_cbolocation1")
        Session.Remove("tmp_Location1")
        Session.Remove("tmp_Kanbanno")

        If txtaffiliateappdate.Text <> "" Then
            Session("KCR-KanbanNo") = Trim(txtkanban1.Text)
            Session("KCR-KanbanDate") = Format(dtkanban.Value, "dd MMM yyyy")
            Session("KCR-SupplierCode") = Trim(cbosupplier.Text)
            Session("KCR-DeliveryLocation") = Trim(cbolocation.Text)
            Session("KCR-ReportCode") = "KanbanCycle"
            Session("KCR-Form") = "KanbanCreate"
            Session("btn") = btnsubmenu.Text
            Session("KCR-kanbanno") = Session("FilterKanbanNo")

            Session("tmp_kanbandate") = Session("xparam_kanbandate")
            Session("tmp_suppID") = Session("xparam_suppID")
            Session("tmp_suppname") = Session("xparam_suppname")
            Session("tmp_affentrydate") = Session("xparam_affentrydate")
            Session("tmp_affentryname") = Session("xparam_affentryname")
            Session("tmp_affappdate") = Session("xparam_affappdate")
            Session("tmp_affappname") = Session("xparam_affappname")
            Session("tmp_suppappdate") = Session("xparam_suppappdate")
            Session("tmp_suppappname") = Session("xparam_suppappname")
            Session("tmp_dt1") = Session("xparam_dt1")
            Session("tmp_dt2") = Session("xparam_dt2")
            Session("tmp_cbosupplier") = Session("xparam_cbosupplier")
            Session("tmp_cbosupplierName") = Session("xparam_cbosupplierName")
            Session("tmp_cbolocation") = Session("xparam_cbolocation")
            Session("tmp_Location") = Session("xparam_Location")
            Session("tmp_cbolocation1") = Session("xparam_cbolocation1")
            Session("tmp_Location1") = Session("xparam_Location1")
            Session("tmp_Kanbanno") = Session("xparam_Kanbanno")

            If Not IsNothing(Request.QueryString("prm")) Then
                Session("K-Param") = Session("xparam_kanbandate") + "|" + Session("xparam_suppID") + "|" + Session("xparam_suppname") + "|" + Session("xparam_affentrydate") + "|" + Session("xparam_affentryname") + "|" + Session("xparam_affappdate") + "|" + Session("xparam_affappname") + "|" + Session("xparam_suppappdate") + "|" + Session("xparam_suppappname") + "|" + Session("xparam_dt1") + "|" + Session("xparam_dt2") + "|" + Session("xparam_cbosupplier") + "|" + Session("xparam_cbosupplierName") + "|" + Session("xparam_cbolocation") + "|" + Session("xparam_Location") + "|" + Session("xparam_cbolocation1") + "|" + Session("xparam_Location1") + "|" + Session("xparam_Kanbanno")
            End If

            Response.Redirect("~/Kanban/ViewReport.aspx")
        Else
            Call clsMsg.DisplayMessage(lblerrmessage, "7011", clsMessage.MsgType.ErrorMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            fillHeader()
        End If
    End Sub
#End Region

#Region "PROCEDURE"
    Private Function uf_GetMOQ(ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String) As Integer
        Dim MOQ As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(MOQ,0) MOQ FROM dbo.MS_PartMapping WHERE PartNo='" + pPartNo + "' AND SupplierID='" + pSupplierID + "' AND AffiliateID='" + pAffiliateID + "'"
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                MOQ = dt.Rows(0)("MOQ")
            End If
            Return MOQ
        End Using
    End Function

    Private Function uf_GetQtybox(ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String) As Integer
        Dim Qty As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(QtyBox,0) Qty FROM dbo.MS_PartMapping WHERE PartNo='" + pPartNo + "' AND SupplierID='" + pSupplierID + "' AND AffiliateID='" + pAffiliateID + "'"
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                Qty = dt.Rows(0)("Qty")
            End If
        End Using
        Return Qty
    End Function

    Public Function uf_GetDataTable(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataTable
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            Dim dt As New DataTable
            da.Fill(ds)
            Return ds.Tables(0)
        Else
            Using Cn As New SqlConnection(clsGlobal.ConnectionString)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                Dim dt As New DataTable
                da.Fill(ds)
                Return ds.Tables(0)
            End Using
        End If
    End Function

    Private Sub colorGrid()
        Grid.VisibleColumns(0).CellStyle.BackColor = Color.White
        Grid.VisibleColumns(12).CellStyle.BackColor = Color.White
        Grid.VisibleColumns(13).CellStyle.BackColor = Color.White
        Grid.VisibleColumns(14).CellStyle.BackColor = Color.White
        Grid.VisibleColumns(15).CellStyle.BackColor = Color.White
        Grid.VisibleColumns(16).CellStyle.BackColor = Color.White

        Grid.VisibleColumns(1).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(2).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(3).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(4).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(5).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(6).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(7).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(8).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(9).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(10).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(11).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(17).CellStyle.BackColor = Color.LightYellow

    End Sub

    Private Sub Clear()
        'dt1.Value = Format(Now, "MMM yyyy")
        dtkanban.Value = ""
        cbosupplier.Text = ""
        txtsuppliername.Text = ""
        txtaffiliateappdate.Text = ""
        txtaffiliateappname.Text = ""
        txtaffiliateentrydate.Text = ""
        txtaffiliateentryname.Text = ""
        txtkanban1.Text = ""
        txtkanban2.Text = ""
        txtkanban3.Text = ""
        txtkanban4.Text = ""
        txttime1.Text = ""
        txttime2.Text = ""
        txttime3.Text = ""
        txttime4.Text = ""
        txtlocation.Text = ""
        txtsuppliername.Text = ""
        cbolocation.Text = ""
        cbosupplier.Text = ""
        Call up_fillcombo()
        dt1.Enabled = True
        cbosupplier.Enabled = True
        cbolocation.Enabled = True
        dtkanban.Enabled = True
        txtkanban1.Enabled = True
        txtkanban2.Enabled = True
        txtkanban3.Enabled = True
        txtkanban4.Enabled = True
        txttime1.Enabled = True
        txttime2.Enabled = True
        txttime3.Enabled = True
        txttime4.Enabled = True
        'Call fillHeader()
        'Call up_GridLoad(dtkanban.Value, Trim(cbosupplier.Text))
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_sql As String
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_sql = " SELECT top 0 " & vbCrLf & _
                  "     idx= '', " & vbCrLf & _
                  " 	cols = '',  " & vbCrLf & _
                  "  	colno = '',  " & vbCrLf & _
                  "  	colpartno = '',  " & vbCrLf & _
                  "  	coldescription='',  " & vbCrLf & _
                  "  	colpono='',  " & vbCrLf & _
                  "  	coluom='',  " & vbCrLf & _
                  "  	colmoq='',  " & vbCrLf & _
                  "  	colqty='',  " & vbCrLf & _
                  "  	colpoqty= '',  " & vbCrLf & _
                  "  	colremainingpo='' ,  "

            ls_sql = ls_sql + "  	colremainingsupplier = '',  " & vbCrLf & _
                              "  	colkanbanqty='',  " & vbCrLf & _
                              "  	colcycle1 = '',  " & vbCrLf & _
                              "  	colcycle2 = '',  " & vbCrLf & _
                              "  	colcycle3 = '',  " & vbCrLf & _
                              "  	colcycle4 = '',  " & vbCrLf & _
                              "  	colbox = '',  " & vbCrLf & _
                              "  	cols1 = '', coluomcode = ''  " & vbCrLf & _
                              "     ,kanbanno1 = '',kanbanno2 = '',kanbanno3 = '',kanbanno4 = '' " & vbCrLf & _
                              "     ,kanbantime1 = '',kanbantime2 = '',kanbantime3 = '',kanbantime4 = '' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Sub fillHeader()
        Dim ls_sql As String
        Dim i As Integer
        Dim ls_seq As String = ""
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

        i = 0
        ls_sql = ""
        ls_sql = "select * from kanban_Master where SupplierID = '" & Trim(cbosupplier.Text) & "' and kanbanDate = '" & Format(dtkanban.Value, "yyyy-MM-dd") & "' and AffiliateID = '" & Session("AffiliateID") & "'"
        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim dsCek As New DataSet
            sqlDA.Fill(dsCek)

            If dsCek.Tables(0).Rows.Count > 0 Then
                Session("KNEW") = "FALSE"
            Else
                Session("KNEW") = "TRUE"
            End If
        End Using

        If cboseq.Text = "1-4" Then
            ls_seq = " AND KM.KanbanCycle in ('1','2','3','4') " & vbCrLf & _
                     " AND MK.KanbanCycle in ('1','2','3','4') " & vbCrLf
        ElseIf cboseq.Text = "5-8" Then
            ls_seq = " AND KM.KanbanCycle in ('5','6','7','8') " & vbCrLf & _
                     " AND MK.KanbanCycle in ('5','6','7','8') "
        ElseIf cboseq.Text = "9-12" Then
            ls_seq = " AND KM.KanbanCycle in ('9','10','11','12') " & vbCrLf & _
                     " AND MK.KanbanCycle in ('9','10','11','12') "
        ElseIf cboseq.Text = "13-16" Then
            ls_seq = " AND KM.KanbanCycle in ('13','14','15','16') " & vbCrLf & _
                     " AND MK.KanbanCycle in ('13','14','15','16') "
        ElseIf cboseq.Text = "17-20" Then
            ls_seq = " AND KM.KanbanCycle in ('17','18','19','20') " & vbCrLf & _
                     " AND MK.KanbanCycle in ('17','18','19','20') "
        End If

        If (Session("FilterKanbanNo") <> "" And Session("FilterKanbanNo") <> clsGlobal.gs_All) Then
            ls_seq = ls_seq + " AND KM.Kanbanno = '" & Trim(Session("FilterKanbanNo")) & "'" & vbCrLf
        End If

        ls_sql = " SELECT KanbanStatus = KM.KanbanStatus,  " & vbCrLf & _
                  " KM.supplierID, " & vbCrLf & _
                  " MSS.SupplierName, " & vbCrLf & _
                  " KM.kanbanDate, " & vbCrLf & _
                  " KM.KanbanCycle, " & vbCrLf & _
                  " ISNULL(KM.EntryUser, '') AS EntryUser , " & vbCrLf & _
                  "         case when isnull(KM.entrydate,'') = '' then '' else convert(char(19),convert(datetime,KM.entrydate),120) end AS EntryDate , " & vbCrLf & _
                  "         ISNULL(SupplierApproveUser, '') AS SupplierUser , " & vbCrLf & _
                  "         case when isnull(SupplierApproveDate,'') = '' then '' else convert(char(19),convert(datetime,SupplierApproveDate),120) end AS supplierDate , " & vbCrLf & _
                  "         ISNULL(AffiliateApproveUser, '') AS AffiliateUser , " & vbCrLf & _
                  "         case when isnull(affiliateApproveDate,'') = '' then '' else convert(char(19),convert(datetime,affiliateApproveDate),120) end AS AffiliateDate, " & vbCrLf & _
                  "         KanbanNo , " & vbCrLf & _
                  "         isnull(KM.DeliveryLocationCode,'') as LocationCode, " & vbCrLf & _
                  "         isnull(DeliveryLocationName, '') as LocationName " & vbCrLf & _
                  "         ,AffiliateName " & vbCrLf & _
                  "         ,ALAMAT = RTRIM(MSA.Address) + ' ' + RTRIM(MSA.City) + ' '+ RTRIM(MSA.PostalCode), " & vbCrLf


        ls_sql = ls_sql + "         ISNULL(CONVERT(CHAR(5), CONVERT(DATETIME, MK.KanbanTime), 114),'00:00') AS KanbanTime " & vbCrLf & _
                          " FROM    kanban_Master KM " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Supplier MSS ON KM.SupplierID = MSS.SupplierID " & vbCrLf & _
                          " LEFT JOIN MS_DeliveryPlace MDP on KM.DeliveryLocationCode = MDP.DeliveryLocationCode " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Affiliate MSA ON MSA.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " LEFT JOIN MS_KanbanTime MK ON MK.AffiliateID = KM.AffiliateID AND MK.kanbanCycle = KM.KanbanCycle " & vbCrLf & _
                          " where KM.SupplierID = '" & Trim(cbosupplier.Text) & "' and kanbanDate = '" & Format(dtkanban.Value, "yyyy-MM-dd") & "' " & vbCrLf
        'If cbolocation.Text <> "" Then
        ls_sql = ls_sql + " and KM.DeliveryLocationCode = '" & Trim(cbolocation.Text) & "'"
        'End If
        ls_sql = ls_sql + ls_seq & vbCrLf

        If cbotype.Text <> "NORMAL" And cbotype.Text <> "" Then
            ls_sql = ls_sql + " AND Right(RTRIM(KM.kanbanno),1) = 'E' " & vbCrLf
        End If

        ls_sql = ls_sql + " order by KanbanNo,KM.KanbanCycle "

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()
            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Session("ALREADY") = "YES"
                Session("kanbanstatus") = ds.Tables(0).Rows(0)("KanbanStatus")
                txtsuppliername.Text = ds.Tables(0).Rows(0)("SupplierName")
                dt1.Text = Format(ds.Tables(0).Rows(0)("kanbandate"), "MMM yyyy")
                Session("KAffiliateName") = Trim(ds.Tables(0).Rows(0)("Affiliatename"))
                Session("KAlamat") = Trim(ds.Tables(0).Rows(0)("ALAMAT"))

                Grid.JSProperties("cpKanban1") = ""
                Grid.JSProperties("cpKanban2") = ""
                Grid.JSProperties("cpKanban3") = ""
                Grid.JSProperties("cpKanban4") = ""
                txtkanban1.Text = ""
                txtkanban2.Text = ""
                txtkanban3.Text = ""
                txtkanban4.Text = ""

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    txtaffiliateentrydate.Text = ds.Tables(0).Rows(i)("entryDate")
                    'txtsupplierapprovaldate.Text = ds.Tables(0).Rows(i)("supplierDate")
                    txtaffiliateappdate.Text = ds.Tables(0).Rows(i)("AffiliateDate")

                    Grid.JSProperties("cpAEDate") = ds.Tables(0).Rows(i)("entrydate")
                    Grid.JSProperties("cpAEName") = ds.Tables(0).Rows(i)("entryuser")
                    Grid.JSProperties("cpASName") = ds.Tables(0).Rows(i)("SupplierUser")
                    Grid.JSProperties("cpAAName") = ds.Tables(0).Rows(i)("AffiliateUser")

                    'If ds.Tables(0).Rows(i)("AffiliateDate") <> "1/1/1900" Then Grid.JSProperties("cpAADate") = Format(ds.Tables(0).Rows(i)("AffiliateDate"), "yyyy-MM-dd HH:mm:ss")
                    'If ds.Tables(0).Rows(i)("supplierDate") <> "1/1/1900" Then Grid.JSProperties("cpASDate") = Format(ds.Tables(0).Rows(i)("supplierDate"), "yyyy-MM-dd HH:mm:ss")

                    Grid.JSProperties("cpASDate") = ds.Tables(0).Rows(i)("supplierDate")
                    Grid.JSProperties("cpAADate") = ds.Tables(0).Rows(i)("AffiliateDate")

                    'txtaffiliateappdate.Text = Format(ds.Tables(0).Rows(i)("AffiliateDate"), "yyyy-MM-dd HH:mm:ss")
                    txtaffiliateentryname.Text = ds.Tables(0).Rows(i)("entryuser")
                    'txtsupplierapprovalname.Text = ds.Tables(0).Rows(i)("SupplierUser")
                    txtaffiliateappname.Text = ds.Tables(0).Rows(i)("AffiliateUser")

                    'If ds.Tables(0).Rows(i)("AffiliateDate") <> "1/1/1900 12:00:00 AM" Then txtaffiliateentrydate.Text = Format(ds.Tables(0).Rows(i)("AffiliateDate"), "yyyy-MM-dd HH:mm:ss")
                    'If ds.Tables(0).Rows(i)("supplierDate") <> "1/1/1900 12:00:00 AM" Then txtsupplierapprovaldate.Text = Format(ds.Tables(0).Rows(i)("supplierDate"), "yyyy-MM-dd HH:mm:ss")


                    'If i = 0 Then
                    '    Grid.JSProperties("cpKanban1") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Grid.JSProperties("cpTime1") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 1 Then
                    '    Grid.JSProperties("cpKanban2") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Grid.JSProperties("cpTime2") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 2 Then
                    '    Grid.JSProperties("cpKanban3") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Grid.JSProperties("cpTime3") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 3 Then
                    '    Grid.JSProperties("cpKanban4") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Grid.JSProperties("cpTime4") = ds.Tables(0).Rows(i)("kanbantime")
                    'End If

                    If ds.Tables(0).Rows(i)("KanbanCycle") = "1" Or ds.Tables(0).Rows(i)("KanbanCycle") = "5" Or ds.Tables(0).Rows(i)("KanbanCycle") = "9" Or ds.Tables(0).Rows(i)("KanbanCycle") = "13" Or ds.Tables(0).Rows(i)("KanbanCycle") = "17" Then
                        Grid.JSProperties("cpKanban1") = ds.Tables(0).Rows(i)("kanbanno")
                        Grid.JSProperties("cpTime1") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "2" Or ds.Tables(0).Rows(i)("KanbanCycle") = "6" Or ds.Tables(0).Rows(i)("KanbanCycle") = "10" Or ds.Tables(0).Rows(i)("KanbanCycle") = "14" Or ds.Tables(0).Rows(i)("KanbanCycle") = "18" Then
                        Grid.JSProperties("cpKanban2") = ds.Tables(0).Rows(i)("kanbanno")
                        Grid.JSProperties("cpTime2") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "3" Or ds.Tables(0).Rows(i)("KanbanCycle") = "7" Or ds.Tables(0).Rows(i)("KanbanCycle") = "11" Or ds.Tables(0).Rows(i)("KanbanCycle") = "15" Or ds.Tables(0).Rows(i)("KanbanCycle") = "19" Then
                        Grid.JSProperties("cpKanban3") = ds.Tables(0).Rows(i)("kanbanno")
                        Grid.JSProperties("cpTime3") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "4" Or ds.Tables(0).Rows(i)("KanbanCycle") = "8" Or ds.Tables(0).Rows(i)("KanbanCycle") = "12" Or ds.Tables(0).Rows(i)("KanbanCycle") = "16" Or ds.Tables(0).Rows(i)("KanbanCycle") = "20" Then
                        Grid.JSProperties("cpKanban4") = ds.Tables(0).Rows(i)("kanbanno")
                        Grid.JSProperties("cpTime4") = ds.Tables(0).Rows(i)("kanbantime")
                    End If

                    '===========================================

                    'If i = 0 Then
                    '    txtkanban1.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime1.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 1 Then
                    '    txtkanban2.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime2.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 2 Then
                    '    txtkanban3.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime3.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 3 Then
                    '    txtkanban4.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime4.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'End If
                    If ds.Tables(0).Rows(i)("KanbanCycle") = "1" Or ds.Tables(0).Rows(i)("KanbanCycle") = "5" Or ds.Tables(0).Rows(i)("KanbanCycle") = "9" Or ds.Tables(0).Rows(i)("KanbanCycle") = "13" Or ds.Tables(0).Rows(i)("KanbanCycle") = "17" Then
                        txtkanban1.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime1.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "2" Or ds.Tables(0).Rows(i)("KanbanCycle") = "6" Or ds.Tables(0).Rows(i)("KanbanCycle") = "10" Or ds.Tables(0).Rows(i)("KanbanCycle") = "14" Or ds.Tables(0).Rows(i)("KanbanCycle") = "18" Then
                        txtkanban2.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime2.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "3" Or ds.Tables(0).Rows(i)("KanbanCycle") = "7" Or ds.Tables(0).Rows(i)("KanbanCycle") = "11" Or ds.Tables(0).Rows(i)("KanbanCycle") = "15" Or ds.Tables(0).Rows(i)("KanbanCycle") = "19" Then
                        txtkanban3.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime3.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "4" Or ds.Tables(0).Rows(i)("KanbanCycle") = "8" Or ds.Tables(0).Rows(i)("KanbanCycle") = "12" Or ds.Tables(0).Rows(i)("KanbanCycle") = "16" Or ds.Tables(0).Rows(i)("KanbanCycle") = "20" Then
                        txtkanban4.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime4.Text = ds.Tables(0).Rows(i)("kanbantime")
                    End If
                Next

                'dt1.Enabled = False
                'cbosupplier.Enabled = False
                'cbolocation.Enabled = False
                'dtkanban.Enabled = False
            Else
                Session("ALREADY") = "NO"
                Session("kanbanstatus") = 0
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                dt1.Enabled = True
                cbosupplier.Enabled = True
                cbolocation.Enabled = True
                dtkanban.Enabled = True
                txtkanban1.Enabled = True
                txtkanban2.Enabled = True
                txtkanban3.Enabled = True
                txtkanban4.Enabled = True

                Grid.JSProperties("cpAEDate") = ""
                Grid.JSProperties("cpAEName") = ""
                Grid.JSProperties("cpASName") = ""
                Grid.JSProperties("cpAAName") = ""

                Grid.JSProperties("cpAADate") = ""
                Grid.JSProperties("cpASDate") = ""

                Grid.JSProperties("cpASDate") = ""
                Grid.JSProperties("cpAADate") = ""

                txtaffiliateentrydate.Text = ""
                txtaffiliateentryname.Text = ""
                'txtsupplierapprovalname.Text = ""
                txtaffiliateappname.Text = ""

                txtaffiliateappdate.Text = ""
                'txtsupplierapprovaldate.Text = ""

                ls_sql = "SELECT urut = Count(kanbanno) --urut = CASE WHEN RIGHT(Rtrim(KanbanNo),2) LIKE '-%' THEN RIGHT(RTRIM(KanbanNo),1) ELSE RIGHT(RTRIM(KanbanNo),2) END " & vbCrLf & _
                         "  FROM( " & vbCrLf & _
                         " SELECT TOP 1 * FROM dbo.Kanban_Master " & vbCrLf & _
                         " where SupplierID = '" & Trim(cbosupplier.Text) & "' and kanbanDate = '" & Format(dtkanban.Value, "yyyy-MM-dd") & "' AND AffiliateID = '" & Session("AffiliateID") & "' and deliverylocationcode = '" & Trim(cbolocation.Text) & "' and kanbanstatus = 0 " & vbCrLf & _
                         " ORDER BY KanbanNo desc )x" & vbCrLf
                Dim sqlS As New SqlDataAdapter(ls_sql, cn)
                Dim ds1 As New DataSet
                sqlS.Fill(ds1)

                Dim ls_CK1 As String = ""
                Dim ls_CK2 As String = ""
                Dim ls_CK3 As String = ""
                Dim ls_CK4 As String = ""

                If cboseq.Text = "1-4" Then
                    ls_CK1 = "1"
                    ls_CK2 = "2"
                    ls_CK3 = "3"
                    ls_CK4 = "4"
                ElseIf cboseq.Text = "5-8" Then
                    ls_CK1 = "5"
                    ls_CK2 = "6"
                    ls_CK3 = "7"
                    ls_CK4 = "8"
                ElseIf cboseq.Text = "9-12" Then
                    ls_CK1 = "9"
                    ls_CK2 = "10"
                    ls_CK3 = "11"
                    ls_CK4 = "12"
                ElseIf cboseq.Text = "13-16" Then
                    ls_CK1 = "13"
                    ls_CK2 = "14"
                    ls_CK3 = "15"
                    ls_CK4 = "16"
                ElseIf cboseq.Text = "17-20" Then
                    ls_CK1 = "17"
                    ls_CK2 = "18"
                    ls_CK3 = "19"
                    ls_CK4 = "20"
                End If

                ls_sql = " select distinct AffiliateID,  " & vbCrLf & _
                  " cycle1 = isnull((select isnull(convert(char(5),kanbantime),'00:00') from ms_kanbantime where affiliateid = '" & Session("AffiliateID") & "' and kanbancycle = '" & ls_CK1 & "'),'00:00'), " & vbCrLf & _
                  " cycle2 = isnull((select isnull(convert(char(5),kanbantime),'00:00') from ms_kanbantime where affiliateid = '" & Session("AffiliateID") & "' and kanbancycle = '" & ls_CK2 & "'),'00:00'), " & vbCrLf & _
                  " cycle3 = isnull((select isnull(convert(char(5),kanbantime),'00:00') from ms_kanbantime where affiliateid = '" & Session("AffiliateID") & "' and kanbancycle = '" & ls_CK3 & "'),'00:00'), " & vbCrLf & _
                  " cycle4 = isnull((select isnull(convert(char(5),kanbantime),'00:00') from ms_kanbantime where affiliateid = '" & Session("AffiliateID") & "' and kanbancycle = '" & ls_CK4 & "'),'00:00') " & vbCrLf & _
                  " from ms_kanbantime where affiliateID = '" & Session("AffiliateID") & "' "

                Dim sqlS2 As New SqlDataAdapter(ls_sql, cn)
                Dim ds2 As New DataSet

                sqlS2.Fill(ds2)

                If ds1.Tables(0).Rows.Count > 0 Then
                    If Format(dtkanban.Value, "yyyyMMdd") <> "" Then
                        Grid.JSProperties("cpKanban1") = Format(dtkanban.Value, "yyyyMMdd") & "-" & Trim(ds1.Tables(0).Rows(0)("urut")) + 1
                        Grid.JSProperties("cpKanban2") = Format(dtkanban.Value, "yyyyMMdd") & "-" & Trim(ds1.Tables(0).Rows(0)("urut")) + 2
                        Grid.JSProperties("cpKanban3") = Format(dtkanban.Value, "yyyyMMdd") & "-" & Trim(ds1.Tables(0).Rows(0)("urut")) + 3
                        Grid.JSProperties("cpKanban4") = Format(dtkanban.Value, "yyyyMMdd") & "-" & Trim(ds1.Tables(0).Rows(0)("urut")) + 4
                    End If
                Else

                    If Format(dtkanban.Value, "yyyyMMdd") <> "" Then
                        If cboseq.Text = "1-4" Then
                            Grid.JSProperties("cpKanban1") = Format(dtkanban.Value, "yyyyMMdd") & "-1"
                            Grid.JSProperties("cpKanban2") = Format(dtkanban.Value, "yyyyMMdd") & "-2"
                            Grid.JSProperties("cpKanban3") = Format(dtkanban.Value, "yyyyMMdd") & "-3"
                            Grid.JSProperties("cpKanban4") = Format(dtkanban.Value, "yyyyMMdd") & "-4"
                        End If

                        If cboseq.Text = "5-8" Then
                            Grid.JSProperties("cpKanban1") = Format(dtkanban.Value, "yyyyMMdd") & "-5"
                            Grid.JSProperties("cpKanban2") = Format(dtkanban.Value, "yyyyMMdd") & "-6"
                            Grid.JSProperties("cpKanban3") = Format(dtkanban.Value, "yyyyMMdd") & "-7"
                            Grid.JSProperties("cpKanban4") = Format(dtkanban.Value, "yyyyMMdd") & "-8"
                        End If

                        If cboseq.Text = "9-12" Then
                            Grid.JSProperties("cpKanban1") = Format(dtkanban.Value, "yyyyMMdd") & "-9"
                            Grid.JSProperties("cpKanban2") = Format(dtkanban.Value, "yyyyMMdd") & "-10"
                            Grid.JSProperties("cpKanban3") = Format(dtkanban.Value, "yyyyMMdd") & "-11"
                            Grid.JSProperties("cpKanban4") = Format(dtkanban.Value, "yyyyMMdd") & "-12"
                        End If

                        If cboseq.Text = "13-16" Then
                            Grid.JSProperties("cpKanban1") = Format(dtkanban.Value, "yyyyMMdd") & "-13"
                            Grid.JSProperties("cpKanban2") = Format(dtkanban.Value, "yyyyMMdd") & "-14"
                            Grid.JSProperties("cpKanban3") = Format(dtkanban.Value, "yyyyMMdd") & "-15"
                            Grid.JSProperties("cpKanban4") = Format(dtkanban.Value, "yyyyMMdd") & "-16"
                        End If

                        If cboseq.Text = "17-20" Then
                            Grid.JSProperties("cpKanban1") = Format(dtkanban.Value, "yyyyMMdd") & "-17"
                            Grid.JSProperties("cpKanban2") = Format(dtkanban.Value, "yyyyMMdd") & "-18"
                            Grid.JSProperties("cpKanban3") = Format(dtkanban.Value, "yyyyMMdd") & "-19"
                            Grid.JSProperties("cpKanban4") = Format(dtkanban.Value, "yyyyMMdd") & "-20"
                        End If

                    End If
                End If

                If ds2.Tables(0).Rows.Count > 0 Then
                    Grid.JSProperties("cpTime1") = ds2.Tables(0).Rows(i)("cycle1")
                    Grid.JSProperties("cpTime2") = ds2.Tables(0).Rows(i)("cycle2")
                    Grid.JSProperties("cpTime3") = ds2.Tables(0).Rows(i)("cycle3")
                    Grid.JSProperties("cpTime4") = ds2.Tables(0).Rows(i)("cycle4")

                    txttime1.Text = ds2.Tables(0).Rows(i)("cycle1")
                    txttime2.Text = ds2.Tables(0).Rows(i)("cycle2")
                    txttime3.Text = ds2.Tables(0).Rows(i)("cycle3")
                    txttime4.Text = ds2.Tables(0).Rows(i)("cycle4")
                Else

                    Grid.JSProperties("cpTime1") = "00:00"
                    Grid.JSProperties("cpTime2") = "00:00"
                    Grid.JSProperties("cpTime3") = "00:00"
                    Grid.JSProperties("cpTime4") = "00:00"
                End If
            End If

            If txtaffiliateappdate.Text = "" Then
                Approve.JSProperties("cpButton") = "APPROVE"
            ElseIf txtaffiliateappdate.Text <> "" Then
                Approve.JSProperties("cpButton") = "UNAPPROVE"
            End If

            cn.Close()
        End Using
    End Sub

    Private Sub fillHeaderAfterApprove()
        Dim ls_sql As String
        Dim i As Integer
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

        i = 0
        ls_sql = ""
        ls_sql = " SELECT   " & vbCrLf & _
                  " KM.supplierID, " & vbCrLf & _
                  " MSS.SupplierName, " & vbCrLf & _
                  " KM.kanbanDate, " & vbCrLf & _
                  " KM.KanbanCycle, " & vbCrLf & _
                  " ISNULL(KM.EntryUser, '') AS EntryUser , " & vbCrLf & _
                  "         case when isnull(KM.entrydate,'') = '' then '' else convert(char(19),convert(datetime,KM.entrydate),120) end AS EntryDate , " & vbCrLf & _
                  "         ISNULL(SupplierApproveUser, '') AS SupplierUser , " & vbCrLf & _
                  "         case when isnull(SupplierApproveDate,'') = '' then '' else convert(char(19),convert(datetime,SupplierApproveDate),120) end AS supplierDate , " & vbCrLf & _
                  "         ISNULL(AffiliateApproveUser, '') AS AffiliateUser , " & vbCrLf & _
                  "         case when isnull(affiliateApproveDate,'') = '' then '' else convert(char(19),convert(datetime,affiliateApproveDate),120) end AS AffiliateDate, " & vbCrLf & _
                  "         KanbanNo , " & vbCrLf & _
                  "         isnull(KM.DeliveryLocationCode,'') as LocationCode, " & vbCrLf & _
                  "         isnull(DeliveryLocationName, '') as LocationName " & vbCrLf & _
                  "         ,AffiliateName " & vbCrLf & _
                  "         ,ALAMAT = RTRIM(MSA.Address) + ' ' + RTRIM(MSA.City) + ' '+ RTRIM(MSA.PostalCode), " & vbCrLf


        ls_sql = ls_sql + "         CONVERT(CHAR(5), ISNULL(CONVERT(DATETIME, KanbanTime),'00:00:00'), 114) AS KanbanTime " & vbCrLf & _
                          " FROM    kanban_Master KM " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Supplier MSS ON KM.SupplierID = MSS.SupplierID " & vbCrLf & _
                          " LEFT JOIN MS_DeliveryPlace MDP on KM.DeliveryLocationCode = MDP.DeliveryLocationCode " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Affiliate MSA ON MSA.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " where KM.SupplierID = '" & Trim(cbosupplier.Text) & "' and kanbanDate = '" & Format(dtkanban.Value, "yyyy-MM-dd") & "' " & vbCrLf
        ls_sql = ls_sql + " and KM.DeliveryLocationCode = '" & Trim(cbolocation.Text) & "'"
        If (Session("FilterKanbanNo") <> "" And Session("FilterKanbanNo") <> clsGlobal.gs_All) Then
            ls_sql = ls_sql + " AND KM.Kanbanno = '" & Trim(Session("FilterKanbanNo")) & "'" & vbCrLf
        End If
        ls_sql = ls_sql + " order by KanbanNo,KanbanCycle "

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtsuppliername.Text = ds.Tables(0).Rows(0)("SupplierName")
                dt1.Text = Format(ds.Tables(0).Rows(0)("kanbandate"), "MMM yyyy")
                Session("KAffiliateName") = Trim(ds.Tables(0).Rows(0)("Affiliatename"))
                Session("KAlamat") = Trim(ds.Tables(0).Rows(0)("ALAMAT"))
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    txtaffiliateentrydate.Text = ds.Tables(0).Rows(i)("entryDate")
                    'txtsupplierapprovaldate.Text = ds.Tables(0).Rows(i)("supplierDate")
                    txtaffiliateappdate.Text = ds.Tables(0).Rows(i)("AffiliateDate")

                    Approve.JSProperties("cpAEDate") = ds.Tables(0).Rows(i)("entrydate")
                    Approve.JSProperties("cpAEName") = ds.Tables(0).Rows(i)("entryuser")
                    Approve.JSProperties("cpASName") = ds.Tables(0).Rows(i)("SupplierUser")
                    Approve.JSProperties("cpAAName") = ds.Tables(0).Rows(i)("AffiliateUser")
                    Approve.JSProperties("cpASDate") = ds.Tables(0).Rows(i)("supplierDate")
                    Approve.JSProperties("cpAADate") = ds.Tables(0).Rows(i)("AffiliateDate")

                    txtaffiliateentryname.Text = ds.Tables(0).Rows(i)("entryuser")
                    'txtsupplierapprovalname.Text = ds.Tables(0).Rows(i)("SupplierUser")
                    txtaffiliateappname.Text = ds.Tables(0).Rows(i)("AffiliateUser")

                    'If i = 0 Then
                    '    Approve.JSProperties("cpKanban1") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Approve.JSProperties("cpTime1") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 1 Then
                    '    Approve.JSProperties("cpKanban2") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Approve.JSProperties("cpTime2") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 2 Then
                    '    Approve.JSProperties("cpKanban3") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Approve.JSProperties("cpTime3") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 3 Then
                    '    Approve.JSProperties("cpKanban4") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Approve.JSProperties("cpTime4") = ds.Tables(0).Rows(i)("kanbantime")
                    'End If
                    If ds.Tables(0).Rows(i)("KanbanCycle") = "1" Or ds.Tables(0).Rows(i)("KanbanCycle") = "5" Or ds.Tables(0).Rows(i)("KanbanCycle") = "9" Or ds.Tables(0).Rows(i)("KanbanCycle") = "13" Or ds.Tables(0).Rows(i)("KanbanCycle") = "17" Then
                        Approve.JSProperties("cpKanban1") = ds.Tables(0).Rows(i)("kanbanno")
                        Approve.JSProperties("cpTime1") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "2" Or ds.Tables(0).Rows(i)("KanbanCycle") = "6" Or ds.Tables(0).Rows(i)("KanbanCycle") = "10" Or ds.Tables(0).Rows(i)("KanbanCycle") = "14" Or ds.Tables(0).Rows(i)("KanbanCycle") = "18" Then
                        Approve.JSProperties("cpKanban2") = ds.Tables(0).Rows(i)("kanbanno")
                        Approve.JSProperties("cpTime2") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "3" Or ds.Tables(0).Rows(i)("KanbanCycle") = "7" Or ds.Tables(0).Rows(i)("KanbanCycle") = "11" Or ds.Tables(0).Rows(i)("KanbanCycle") = "15" Or ds.Tables(0).Rows(i)("KanbanCycle") = "19" Then
                        Approve.JSProperties("cpKanban3") = ds.Tables(0).Rows(i)("kanbanno")
                        Approve.JSProperties("cpTime3") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "4" Or ds.Tables(0).Rows(i)("KanbanCycle") = "8" Or ds.Tables(0).Rows(i)("KanbanCycle") = "12" Or ds.Tables(0).Rows(i)("KanbanCycle") = "16" Or ds.Tables(0).Rows(i)("KanbanCycle") = "20" Then
                        Approve.JSProperties("cpKanban4") = ds.Tables(0).Rows(i)("kanbanno")
                        Approve.JSProperties("cpTime4") = ds.Tables(0).Rows(i)("kanbantime")
                    End If

                    '===========================================

                    'If i = 0 Then
                    '    txtkanban1.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime1.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 1 Then
                    '    txtkanban2.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime2.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 2 Then
                    '    txtkanban3.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime3.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 3 Then
                    '    txtkanban4.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime4.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'End If
                    If ds.Tables(0).Rows(i)("KanbanCycle") = "1" Or ds.Tables(0).Rows(i)("KanbanCycle") = "5" Or ds.Tables(0).Rows(i)("KanbanCycle") = "9" Or ds.Tables(0).Rows(i)("KanbanCycle") = "13" Or ds.Tables(0).Rows(i)("KanbanCycle") = "17" Then
                        txtkanban1.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime1.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "2" Or ds.Tables(0).Rows(i)("KanbanCycle") = "6" Or ds.Tables(0).Rows(i)("KanbanCycle") = "10" Or ds.Tables(0).Rows(i)("KanbanCycle") = "14" Or ds.Tables(0).Rows(i)("KanbanCycle") = "18" Then
                        txtkanban2.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime2.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "3" Or ds.Tables(0).Rows(i)("KanbanCycle") = "7" Or ds.Tables(0).Rows(i)("KanbanCycle") = "11" Or ds.Tables(0).Rows(i)("KanbanCycle") = "15" Or ds.Tables(0).Rows(i)("KanbanCycle") = "19" Then
                        txtkanban3.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime3.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "4" Or ds.Tables(0).Rows(i)("KanbanCycle") = "8" Or ds.Tables(0).Rows(i)("KanbanCycle") = "12" Or ds.Tables(0).Rows(i)("KanbanCycle") = "16" Or ds.Tables(0).Rows(i)("KanbanCycle") = "20" Then
                        txtkanban4.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime4.Text = ds.Tables(0).Rows(i)("kanbantime")
                    End If
                Next
            End If

            If txtaffiliateappname.Text = "" Then
                Approve.JSProperties("cpButton") = "APPROVE"
            Else
                Approve.JSProperties("cpButton") = "UNAPPROVE"
            End If
            cn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad(ByVal pkanbandate As Date, ByVal psuppID As String, ByVal pSeqno As String)

        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim pFieldDelivery As String
        Dim ls_seq As String = ""
        Dim pCycle1 As String = ""
        Dim pCycle2 As String = ""
        Dim pCycle3 As String = ""
        Dim pCycle4 As String = ""
        Dim ls_FilterKanbanNo As String = ""

        If Trim(xKanbanNo) <> "" Then
            ls_FilterKanbanNo = " and KanbanNo = '" & xKanbanNo & "'" & vbCrLf
        End If

        If Session("cycle") <> "" Then
            pSeqno = Session("cycle")
            Session("cycle") = ""
        End If

        If pSeqno = "1-4" Then
            ls_seq = " AND KM.KanbanCycle in ('1','2','3','4') "
            pCycle1 = "1"
            pCycle2 = "2"
            pCycle3 = "3"
            pCycle4 = "4"
        ElseIf pSeqno = "5-8" Then
            ls_seq = " AND KM.KanbanCycle in ('5','6','7','8') "
            pCycle1 = "5"
            pCycle2 = "6"
            pCycle3 = "7"
            pCycle4 = "8"
        ElseIf pSeqno = "9-12" Then
            ls_seq = " AND KM.KanbanCycle in ('9','10','11','12') "
            pCycle1 = "9"
            pCycle2 = "10"
            pCycle3 = "11"
            pCycle4 = "12"
        ElseIf pSeqno = "13-16" Then
            ls_seq = " AND KM.KanbanCycle in ('13','14','15','16') "
            pCycle1 = "13"
            pCycle2 = "14"
            pCycle3 = "15"
            pCycle4 = "16"
        ElseIf pSeqno = "17-20" Then
            ls_seq = " AND KM.KanbanCycle in ('17','18','19','20') "
            pCycle1 = "17"
            pCycle2 = "18"
            pCycle3 = "19"
            pCycle4 = "20"
        End If

        If (Session("FilterKanbanNo") <> "" And Session("FilterKanbanNo") <> clsGlobal.gs_All) Then
            ls_seq = ls_seq + " AND KM.Kanbanno = '" & Trim(Session("FilterKanbanNo")) & "'" & vbCrLf
        End If

        If Format(pkanbandate, "dd") = "01" Then
            pFieldDelivery = "DeliveryD1"
        ElseIf Format(pkanbandate, "dd") = "02" Then
            pFieldDelivery = "DeliveryD2"
        ElseIf Format(pkanbandate, "dd") = "03" Then
            pFieldDelivery = "DeliveryD3"
        ElseIf Format(pkanbandate, "dd") = "04" Then
            pFieldDelivery = "DeliveryD4"
        ElseIf Format(pkanbandate, "dd") = "05" Then
            pFieldDelivery = "DeliveryD5"
        ElseIf Format(pkanbandate, "dd") = "06" Then
            pFieldDelivery = "DeliveryD6"
        ElseIf Format(pkanbandate, "dd") = "07" Then
            pFieldDelivery = "DeliveryD7"
        ElseIf Format(pkanbandate, "dd") = "08" Then
            pFieldDelivery = "DeliveryD8"
        ElseIf Format(pkanbandate, "dd") = "09" Then
            pFieldDelivery = "DeliveryD9"
        Else
            pFieldDelivery = "DeliveryD" & Format(pkanbandate, "dd")
        End If

        If Format(pkanbandate, "dd") = "" Then
            pFieldDelivery = "DeliveryD1"
        End If
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""

            ls_SQL = "SELECT " & vbCrLf & _
                "idx = '', " & vbCrLf & _
                "cols = cols, " & vbCrLf & _
                "colno = ROW_NUMBER() OVER(ORDER BY cols DESC), " & vbCrLf & _
                "colpartno = colpartno, " & vbCrLf & _
                "coldescription = coldescription, " & vbCrLf & _
                "colpono = colpono, " & vbCrLf & _
                "coluom = coluom, " & vbCrLf & _
                "colmoq = colmoq, " & vbCrLf & _
                "colqty = colqty, " & vbCrLf & _
                "colpoqty = colpoqty, " & vbCrLf & _
                "colremainingpo = colremainingpo, " & vbCrLf & _
                "colremainingsupplier = ISNULL(colremainingsupplier,0), " & vbCrLf & _
                "coldeliveryqty = ISNULL(coldeliveryqty,0), " & vbCrLf & _
                "colkanbanqty = colkanbanqty, " & vbCrLf & _
                "colcycle1 = colcycle1, " & vbCrLf & _
                "colcycle2 = CASE WHEN cols = 1 THEN colcycle2 ELSE (CASE WHEN (colcycle1 + colcycle2) > colkanbanqty THEN (colcycle1 + colcycle2) - colkanbanqty ELSE colcycle2 END) END, " & vbCrLf & _
                "colcycle3 = CASE WHEN cols = 1 THEN colcycle3 ELSE (CASE WHEN (colcycle1 + colcycle2 + colcycle3) > colkanbanqty THEN (colcycle1 + colcycle2 + colcycle3) - colkanbanqty ELSE colcycle3 END) END, " & vbCrLf & _
                "colcycle4 = colcycle4, colbox = colbox, cols1 = cols1, coluomcode = coluomcode, kanbanno1, kanbanno2, kanbanno3, kanbanno4, kanbantime1, kanbantime2, kanbantime3, kanbantime4 " & vbCrLf

            ls_SQL = ls_SQL & "FROM( " & vbCrLf & _
                "Select DISTINCT " & vbCrLf & _
                "cols = 1, " & vbCrLf & _
                "colno = '0', " & vbCrLf & _
                "colpartno = KD.partNo, " & vbCrLf & _
                "coldescription = MP.partname, " & vbCrLf & _
                "colpono = KD.pono, " & vbCrLf & _
                "coluom = ISNULL(MUC.Description,''), " & vbCrLf & _
                "colmoq = PMP.MOQ, " & vbCrLf & _
                "colqty = PMP.QtyBox, " & vbCrLf & _
                "colpoqty = COALESCE(PRD.POQty,ISNULL(PD.POQty, 0)), " & vbCrLf & _
                "colremainingpo = COALESCE(PRD.POQty,ISNULL(PD.POQty, 0)) - ( " & vbCrLf & _
                "SELECT ISNULL(SUM(qty),0) " & vbCrLf & _
                "FROM( " & vbCrLf & _
                "SELECT DISTINCT MAS.Kanbanno, KanbanQty AS qty " & vbCrLf & _
                "FROM dbo.Kanban_Detail DET " & vbCrLf & _
                "INNER JOIN Kanban_Master MAS ON MAS.AffiliateID = DET.AffiliateID AND MAS.KanbanNo = DET.KanbanNo AND MAS.SupplierID = DET.SupplierID AND MAS.DeliveryLocationCode = DET.DeliveryLocationCode " & vbCrLf & _
                "WHERE PONo = PM.PONo " & vbCrLf & _
                "AND PartNo = PD.PartNo " & vbCrLf & _
                "AND DET.AffiliateID = PM.AffiliateID " & vbCrLf & _
                "AND DET.SupplierID = PM.SupplierID " & vbCrLf & _
                "AND MAS.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                ")x " & vbCrLf & _
                "), " & vbCrLf

            ls_SQL = ls_SQL & "colremainingsupplier = MSS.DailyDeliveryCapacity - ( " & vbCrLf & _
                "Select ISNULL(SUM(qty), 0) " & vbCrLf & _
                "FROM( " & vbCrLf & _
                "SELECT DISTINCT MAS.Kanbanno, KanbanQty AS qty " & vbCrLf & _
                "FROM dbo.Kanban_Detail DET " & vbCrLf & _
                "INNER JOIN Kanban_Master MAS ON MAS.AffiliateID = DET.AffiliateID AND MAS.KanbanNo = DET.KanbanNo AND MAS.SupplierID = DET.SupplierID AND MAS.DeliveryLocationCode = DET.DeliveryLocationCode " & vbCrLf & _
                "WHERE(PONo = PM.PONo) " & vbCrLf & _
                "AND PartNo = PD.PartNo " & vbCrLf & _
                "AND DET.AffiliateID = PM.AffiliateID " & vbCrLf & _
                "AND DET.SupplierID = PM.SupplierID " & vbCrLf & _
                "AND MAS.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                ")x " & vbCrLf & _
                "), " & vbCrLf & _
                "coldeliveryqty = ISNULL(PD.DeliveryD4,0), " & vbCrLf & _
                "colkanbanqty = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle1) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0) + " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle2) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0) + " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle3) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0) + " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle4) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0), " & vbCrLf

            ls_SQL = ls_SQL & "colcycle1 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle1) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0), " & vbCrLf & _
                "colcycle2 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle2) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0), " & vbCrLf & _
                "colcycle3 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle3) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0), " & vbCrLf & _
                "colcycle4 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle4) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0), " & vbCrLf

            ls_SQL = ls_SQL & "colbox = ( " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle1) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0) + " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle2) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0) + " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle3) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0) + " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KanbanQty " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle4) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), 0) " & vbCrLf & _
                ") / PMP.QtyBox, " & vbCrLf & _
                "cols1 = '1', " & vbCrLf & _
                "coluomcode = MP.UnitCls, " & vbCrLf

            ls_SQL = ls_SQL & "kanbanno1 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KMI.kanbanno " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle1) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf

            ls_SQL = ls_SQL + "), ''), " & vbCrLf & _
                "kanbanno2 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KMI.kanbanno " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle2) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), ''), " & vbCrLf & _
                "kanbanno3 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "Select KMI.kanbanno " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle3) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), ''), " & vbCrLf & _
                "kanbanno4 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "SELECT KMI.kanbanno " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle4) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), ''), " & vbCrLf

            ls_SQL = ls_SQL & "kanbantime1 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "SELECT CONVERT(CHAR(5), ISNULL(kanbantime,'00:00:00')) " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle1) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), '00:00'), " & vbCrLf & _
                "kanbantime2 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "SELECT CONVERT(CHAR(5), ISNULL(kanbantime,'00:00:00')) " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle2) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), '00:00'), " & vbCrLf & _
                "kanbantime3 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "SELECT CONVERT(CHAR(5), ISNULL(kanbantime,'00:00:00')) " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle3) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), '00:00'), " & vbCrLf & _
                "kanbantime4 = " & vbCrLf & _
                "ISNULL(( " & vbCrLf & _
                "SELECT CONVERT(CHAR(5), ISNULL(kanbantime,'00:00:00')) " & vbCrLf & _
                "FROM dbo.Kanban_Master KMI " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID AND KMI.SupplierID = KDI.SupplierID AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode " & vbCrLf & _
                "WHERE KMI.KanbanCycle = '" & Trim(pCycle4) & "' " & vbCrLf & _
                "AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                "AND KDI.PartNo = KD.partNo " & vbCrLf & _
                "AND KDI.PONo = KD.PONo " & vbCrLf & _
                "AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                "AND KMI.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "), '00:00') " & vbCrLf

            ls_SQL = ls_SQL & "FROM dbo.Kanban_Master KM " & vbCrLf & _
                "LEFT JOIN dbo.Kanban_Detail KD ON KM.KanbanNo = KD.KanbanNo AND KM.AffiliateID = KD.AffiliateID AND KM.SupplierID = KD.SupplierID AND KM.DeliveryLocationCode = KD.DeliveryLocationCode " & vbCrLf & _
                "LEFT JOIN dbo.po_detailUpload PD ON KD.PartNo = PD.PartNo AND KD.PONo = PD.PONo AND KD.SupplierID = PD.SupplierID AND KD.AffiliateID = PD.AffiliateID " & vbCrLf & _
                "LEFT JOIN PO_Master PM ON PM.PoNo = PD.PONo AND PM.AffiliateID = PD.AffiliateID AND PM.SupplierID = PD.SupplierID " & vbCrLf & _
                "LEFT JOIN dbo.PORev_Master PRM ON PM.AffiliateID = PRM.AffiliateID AND PRM.PONo = PM.PONo AND PRM.SupplierID = PM.SupplierID " & vbCrLf & _
                "LEFT JOIN dbo.PORev_Detail PRD ON PRD.PONo = PRM.PONo AND PRD.AffiliateID = PRM.AffiliateID AND PRD.SupplierID = PRM.SupplierID AND PRD.PartNo = PD.PartNo " & vbCrLf & _
                "AND PRD.SeqNo = ( " & vbCrLf & _
                "Select MAX(seqNO) " & vbCrLf & _
                "FROM PORev_Detail A " & vbCrLf & _
                "WHERE(A.PONo = PD.PONo) " & vbCrLf & _
                "AND A.AffiliateID = PD.AffiliateID " & vbCrLf & _
                "AND A.SupplierID = PD.SupplierID " & vbCrLf & _
                "AND A.PartNo = PD.PartNo " & vbCrLf & _
                ") " & vbCrLf & _
                "LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo " & vbCrLf & _
                "LEFT JOIN dbo.MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls " & vbCrLf & _
                "LEFT JOIN dbo.MS_SupplierCapacity MSS ON MSS.SupplierID = PD.SupplierID AND MSS.PartNo = PD.PartNo	" & vbCrLf & _
                "LEFT JOIN dbo.MS_PartMapping PMP ON PMP.PartNo = KD.PartNo AND PMP.AffiliateID = KD.AffiliateID AND PMP.SupplierID = KD.SupplierID " & vbCrLf & _
                "WHERE ISNULL(KD.partNo,'') <> '' " & vbCrLf & _
                "AND KM.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                "AND KM.AffiliateID = '" & Session("AffiliateID") & " ' " & vbCrLf & _
                "AND KM.SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                "AND KM.DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf

            ls_SQL = ls_SQL + ls_seq & vbCrLf

            If cbotype.Text <> "NORMAL" And cbotype.Text <> "" Then
                ls_SQL = ls_SQL + "AND Right(RTRIM(KM.kanbanno),1) = 'E' " & vbCrLf
            End If

            ls_SQL = ls_SQL & "/*UNION ALL " & vbCrLf

            ls_SQL = ls_SQL & "SELECT " & vbCrLf & _
                "cols, " & vbCrLf & _
                "colno, " & vbCrLf & _
                "colpartno, " & vbCrLf & _
                "coldescription, " & vbCrLf & _
                "colpono, " & vbCrLf & _
                "coluom, " & vbCrLf & _
                "colmoq, " & vbCrLf & _
                "colqty, " & vbCrLf & _
                "colpoqty, " & vbCrLf & _
                "colremainingpo = colremainingpo - totkanban, " & vbCrLf & _
                "colremainingsupplier = colremainingsupplier - totkanban, " & vbCrLf & _
                "coldeliveryqty = colkanbanqty, " & vbCrLf & _
                "colkanbanqty, " & vbCrLf & _
                "colcycle1 = " & vbCrLf & _
                "CASE WHEN ( " & vbCrLf & _
                "CASE WHEN (colkanbanqty - CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0)) > 0 " & vbCrLf & _
                "THEN CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0) ELSE 0 END " & vbCrLf & _
                ") = 0 THEN colkanbanqty " & vbCrLf & _
                "ELSE( " & vbCrLf & _
                "CASE WHEN (colkanbanqty - CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0)) > 0 " & vbCrLf & _
                "THEN CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0) ELSE 0 END " & vbCrLf & _
                ") END, " & vbCrLf & _
                "colcycle2 = " & vbCrLf & _
                "CASE WHEN (colkanbanqty - CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0)) > 0 " & vbCrLf & _
                "THEN CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0) " & vbCrLf & _
                "ELSE 0 END, " & vbCrLf & _
                "colcycle3 = " & vbCrLf & _
                "CASE WHEN (colkanbanqty - CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0)) > 0 " & vbCrLf & _
                "THEN CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0) " & vbCrLf & _
                "ELSE 0 END, " & vbCrLf & _
                "colcycle4 = " & vbCrLf & _
                "CASE WHEN (colkanbanqty - (CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0)) * 3) > 0 " & vbCrLf & _
                "THEN colkanbanqty - ((CEILING(FLOOR(colkanbanqty/4) / ISNULL(colqty,0)) * ISNULL(colqty,0)) )*3 " & vbCrLf & _
                "ELSE 0 END, " & vbCrLf & _
                "colbox = (colkanbanqty / ISNULL(colqty,0)), " & vbCrLf & _
                "cols1," & vbCrLf & _
                "coluomcode," & vbCrLf & _
                "kanbanno1, " & vbCrLf & _
                "kanbanno2, " & vbCrLf & _
                "kanbanno3, " & vbCrLf & _
                "kanbanno4, " & vbCrLf & _
                "kanbantime1, " & vbCrLf & _
                "kanbantime2, " & vbCrLf & _
                "kanbantime3, " & vbCrLf & _
                "kanbantime4 " & vbCrLf

            ls_SQL = ls_SQL & "FROM ( " & vbCrLf & _
                "SELECT " & vbCrLf & _
                "cols = 0, " & vbCrLf & _
                "colno = '1', " & vbCrLf & _
                "colpartno = POD.PartNo, " & vbCrLf & _
                "coldescription = ISNULL(MP.PartName,''), " & vbCrLf & _
                "colpono = POM.PONo, " & vbCrLf & _
                "coluom = ISNULL(MUC.Description,''), " & vbCrLf & _
                "colmoq = PMP.MOQ, " & vbCrLf & _
                "colqty = PMP.QtyBox, " & vbCrLf & _
                "colpoqty = COALESCE(PRD.POQty, POD.POQty), " & vbCrLf & _
                "colremainingpo = COALESCE(PRD.POQty,ISNULL(POD.POQty, 0)), " & vbCrLf & _
                "colremainingsupplier = MSS.DailyDeliveryCapacity, " & vbCrLf & _
                "coldeliveryqty = COALESCE(PRD.DeliveryD4, POD.DeliveryD4), " & vbCrLf & _
                "totkanban = ( " & vbCrLf & _
                "Select ISNULL(SUM(qty), 0) " & vbCrLf & _
                "FROM ( " & vbCrLf & _
                "SELECT DISTINCT MAS.Kanbanno, KanbanQty AS qty " & vbCrLf & _
                "FROM dbo.Kanban_Detail DET " & vbCrLf & _
                "INNER JOIN Kanban_Master MAS ON MAS.AffiliateID = DET.AffiliateID AND MAS.KanbanNo = DET.KanbanNo AND MAS.SupplierID = DET.SupplierID AND MAS.DeliveryLocationCode = DET.DeliveryLocationCode " & vbCrLf & _
                "WHERE PONo = POM.PONo " & vbCrLf & _
                "AND PartNo = POD.PartNo " & vbCrLf & _
                "AND DET.AffiliateID = POM.AffiliateID " & vbCrLf & _
                "AND DET.SupplierID = POM.SupplierID " & vbCrLf & _
                "AND MAS.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                ")x " & vbCrLf & _
                "), " & vbCrLf & _
                "colkanbanqty = " & vbCrLf & _
                "CASE WHEN COALESCE(PRD.DeliveryD4, POD.DeliveryD4) = 0 THEN 0 " & vbCrLf & _
                "ELSE COALESCE(PRD.DeliveryD4, POD.DeliveryD4) - ( " & vbCrLf & _
                "Select ISNULL(SUM(qty), 0) " & vbCrLf & _
                "FROM( " & vbCrLf & _
                "SELECT DISTINCT MAS.Kanbanno, KanbanQty AS qty " & vbCrLf & _
                "FROM dbo.Kanban_Detail DET " & vbCrLf & _
                "INNER JOIN Kanban_Master MAS ON MAS.AffiliateID = DET.AffiliateID AND MAS.KanbanNo = DET.KanbanNo AND MAS.SupplierID = DET.SupplierID AND MAS.DeliveryLocationCode = DET.DeliveryLocationCode " & vbCrLf & _
                "WHERE(PONo = POM.PONo) " & vbCrLf & _
                "AND PartNo = POD.PartNo " & vbCrLf & _
                "AND DET.AffiliateID = POM.AffiliateID " & vbCrLf & _
                "AND DET.SupplierID = POM.SupplierID " & vbCrLf & _
                "AND MAS.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                ")x " & vbCrLf & _
                ")END, " & vbCrLf

            ls_SQL = ls_SQL & "colcycle1 = " & vbCrLf & _
                "CASE WHEN ( " & vbCrLf & _
                "CASE WHEN (COALESCE(PRD.DeliveryD4, POD.DeliveryD4) - CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0)) > 0 " & vbCrLf & _
                "THEN CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0) " & vbCrLf & _
                "ELSE 0 END " & vbCrLf & _
                ") = 0 THEN COALESCE(PRD.DeliveryD4, POD.DeliveryD4) " & vbCrLf & _
                "ELSE ( " & vbCrLf & _
                "CASE WHEN (COALESCE(PRD.DeliveryD4, POD.DeliveryD4) - CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0)) > 0 " & vbCrLf & _
                "THEN CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0) " & vbCrLf & _
                "ELSE 0 END " & vbCrLf & _
                ") END, " & vbCrLf & _
                "colcycle2 = " & vbCrLf & _
                "CASE WHEN (COALESCE(PRD.DeliveryD4, POD.DeliveryD4) - CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0)) > 0 " & vbCrLf & _
                "THEN CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0) " & vbCrLf & _
                "ELSE 0 END, " & vbCrLf & _
                "colcycle3 = " & vbCrLf & _
                "CASE WHEN (COALESCE(PRD.DeliveryD4, POD.DeliveryD4) - CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0)) > 0 " & vbCrLf & _
                "THEN CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0) " & vbCrLf & _
                "ELSE 0 END, " & vbCrLf & _
                "colcycle4 = " & vbCrLf & _
                "CASE WHEN (COALESCE(PRD.DeliveryD4, POD.DeliveryD4) - (CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0)) * 3) > 0 " & vbCrLf & _
                "THEN COALESCE(PRD.DeliveryD4, POD.DeliveryD4) - ((CEILING(FLOOR(COALESCE(PRD.DeliveryD4, POD.DeliveryD4)/4) / ISNULL(PMP.QtyBox,0)) * ISNULL(PMP.QtyBox,0)) )*3 " & vbCrLf & _
                "ELSE 0 END, " & vbCrLf & _
                "colbox = (COALESCE(PRD.DeliveryD4, POD.DeliveryD4) / ISNULL(PMP.QtyBox,0)), " & vbCrLf & _
                "cols1 = '', " & vbCrLf & _
                "coluomcode = MP.UnitCls, " & vbCrLf

            If (Session("KNEW") = "TRUE" And Session("ALREADY") = "NO") Or (Session("KNEW") = "FALSE" And Session("ALREADY") = "NO") Then
                ls_SQL = ls_SQL & "kanbanno1 = ISNULL('20160404' + '-' +  CONVERT(CHAR,(SELECT urut = CASE WHEN RIGHT(Rtrim(KanbanNo),2) LIKE '-%' THEN ISNULL(RIGHT(RTRIM(KanbanNo),1),0) ELSE ISNULL(RIGHT(RTRIM(KanbanNo),2),0) END " & vbCrLf & _
                    "FROM( " & vbCrLf & _
                    "SELECT TOP 1 * FROM dbo.Kanban_Master " & vbCrLf & _
                    "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                    "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                    "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                    "ORDER BY KanbanNo DESC " & vbCrLf & _
                    ")x) + 1), '20160404-1'), " & vbCrLf & _
                    "kanbanno2 = ISNULL('20160404' + '-' +  CONVERT(CHAR,(SELECT urut = CASE WHEN RIGHT(Rtrim(KanbanNo),2) LIKE '-%' THEN ISNULL(RIGHT(RTRIM(KanbanNo),1),0) ELSE ISNULL(RIGHT(RTRIM(KanbanNo),2),0) END " & vbCrLf & _
                    "FROM( " & vbCrLf & _
                    "SELECT TOP 1 * FROM dbo.Kanban_Master " & vbCrLf & _
                    "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                    "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                    "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                    "ORDER BY KanbanNo DESC " & vbCrLf & _
                    ")x) + 2), '20160404-2'), " & vbCrLf & _
                    "kanbanno3 = ISNULL('20160404' + '-' +  CONVERT(CHAR,(SELECT urut = CASE WHEN RIGHT(Rtrim(KanbanNo),2) LIKE '-%' THEN ISNULL(RIGHT(RTRIM(KanbanNo),1),0) ELSE ISNULL(RIGHT(RTRIM(KanbanNo),2),0) END " & vbCrLf & _
                    "FROM( " & vbCrLf & _
                    "SELECT TOP 1 * FROM dbo.Kanban_Master " & vbCrLf & _
                    "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                    "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                    "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                    "ORDER BY KanbanNo DESC " & vbCrLf & _
                    ")x) + 3), '20160404-3'), " & vbCrLf & _
                    "kanbanno4 = ISNULL('20160404' + '-' +  CONVERT(CHAR,(SELECT urut = CASE WHEN RIGHT(Rtrim(KanbanNo),2) LIKE '-%' THEN ISNULL(RIGHT(RTRIM(KanbanNo),1),0) ELSE ISNULL(RIGHT(RTRIM(KanbanNo),2),0) END " & vbCrLf & _
                    "FROM( " & vbCrLf & _
                    "SELECT TOP 1 * FROM dbo.Kanban_Master " & vbCrLf & _
                    "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                    "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                    "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                    "ORDER BY KanbanNo DESC " & vbCrLf & _
                    ")x) + 4), '20160404-4'), " & vbCrLf & _
                    "kanbantime1 = (SELECT ISNULL(kanbantime,'00:00:00') FROM MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & Trim(pCycle1) & "'), " & vbCrLf & _
                    "kanbantime2 = (SELECT ISNULL(kanbantime,'00:00:00') FROM MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & Trim(pCycle2) & "'), " & vbCrLf & _
                    "kanbantime3 = (SELECT ISNULL(kanbantime,'00:00:00') FROM MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & Trim(pCycle3) & "'), " & vbCrLf & _
                    "kanbantime4 = (SELECT ISNULL(kanbantime,'00:00:00') FROM MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & Trim(pCycle4) & "') " & vbCrLf
            ElseIf Session("KNEW") = "FALSE" And Session("ALREADY") = "YES" Then
                ls_SQL = ls_SQL & "kanbanno1 = ( " & vbCrLf & _
                   "SELECT Kanbanno FROM Kanban_Master WITH(NOLOCK) " & vbCrLf & _
                   "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                   "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                   "AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                   "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                   "AND KanbanCycle = '" & Trim(pCycle1) & "' " & vbCrLf & _
                   "and kanbanstatus = '" & Trim(Session("kanbanstatus")) & "' " & vbCrLf

                If ls_FilterKanbanNo <> "" Then ls_SQL = ls_SQL + ls_FilterKanbanNo

                ls_SQL = ls_SQL + "), " & vbCrLf & _
                   "kanbanno2 = ( " & vbCrLf & _
                   "SELECT Kanbanno FROM Kanban_Master WITH(NOLOCK) " & vbCrLf & _
                   "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                   "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                   "AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                   "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                   "AND KanbanCycle = '" & Trim(pCycle2) & "' " & vbCrLf & _
                   "and kanbanstatus = '" & Trim(Session("kanbanstatus")) & "' " & vbCrLf
                If ls_FilterKanbanNo <> "" Then ls_SQL = ls_SQL + ls_FilterKanbanNo

                ls_SQL = ls_SQL + "), " & vbCrLf & _
                   "kanbanno3 = ( " & vbCrLf & _
                   "SELECT Kanbanno FROM Kanban_Master WITH(NOLOCK) " & vbCrLf & _
                   "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                   "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                   "AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                   "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                   "AND KanbanCycle = '" & Trim(pCycle3) & "' " & vbCrLf & _
                   "and kanbanstatus = '" & Trim(Session("kanbanstatus")) & "' " & vbCrLf
                If ls_FilterKanbanNo <> "" Then ls_SQL = ls_SQL + ls_FilterKanbanNo

                ls_SQL = ls_SQL + "), " & vbCrLf & _
                   "kanbanno4 = ( " & vbCrLf & _
                   "SELECT Kanbanno FROM Kanban_Master WITH(NOLOCK) " & vbCrLf & _
                   "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                   "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                   "AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                   "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                   "AND KanbanCycle = '" & Trim(pCycle4) & "' " & vbCrLf & _
                   "and kanbanstatus = '" & Trim(Session("kanbanstatus")) & "' " & vbCrLf
                If ls_FilterKanbanNo <> "" Then ls_SQL = ls_SQL + ls_FilterKanbanNo

                ls_SQL = ls_SQL + "), " & vbCrLf & _
                   "kanbantime1 = ( " & vbCrLf & _
                   "SELECT ISNULL(kanbantime,'00:00:00') FROM Kanban_Master WITH(NOLOCK) " & vbCrLf & _
                   "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                   "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                   "AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                   "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                   "AND KanbanCycle = '" & Trim(pCycle1) & "' " & vbCrLf & _
                   "and kanbanstatus = '" & Trim(Session("kanbanstatus")) & "' " & vbCrLf
                If ls_FilterKanbanNo <> "" Then ls_SQL = ls_SQL + ls_FilterKanbanNo

                ls_SQL = ls_SQL + "), " & vbCrLf & _
                   "kanbantime2 = ( " & vbCrLf & _
                   "SELECT ISNULL(kanbantime,'00:00:00') FROM Kanban_Master WITH(NOLOCK) " & vbCrLf & _
                   "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                   "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                   "AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                   "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                   "AND KanbanCycle = '" & Trim(pCycle2) & "' " & vbCrLf & _
                   "and kanbanstatus = '" & Trim(Session("kanbanstatus")) & "' " & vbCrLf
                If ls_FilterKanbanNo <> "" Then ls_SQL = ls_SQL + ls_FilterKanbanNo

                ls_SQL = ls_SQL + "), " & vbCrLf & _
                   "kanbantime3 = ( " & vbCrLf & _
                   "SELECT ISNULL(kanbantime,'00:00:00') FROM Kanban_Master WITH(NOLOCK) " & vbCrLf & _
                   "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                   "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                   "AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                   "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                   "AND KanbanCycle = '" & Trim(pCycle3) & "' " & vbCrLf & _
                   "and kanbanstatus = '" & Trim(Session("kanbanstatus")) & "' " & vbCrLf
                If ls_FilterKanbanNo <> "" Then ls_SQL = ls_SQL + ls_FilterKanbanNo

                ls_SQL = ls_SQL + "), " & vbCrLf & _
                   "kanbantime4 = ( " & vbCrLf & _
                   "SELECT ISNULL(kanbantime,'00:00:00') FROM Kanban_Master WITH(NOLOCK) " & vbCrLf & _
                   "WHERE SupplierID = '" & Trim(psuppID) & "' " & vbCrLf & _
                   "AND kanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                   "AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                   "AND AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                   "AND KanbanCycle = '" & Trim(pCycle4) & "' " & vbCrLf & _
                   "and kanbanstatus = '" & Trim(Session("kanbanstatus")) & "' " & vbCrLf
                If ls_FilterKanbanNo <> "" Then ls_SQL = ls_SQL + ls_FilterKanbanNo
                ls_SQL = ls_SQL + ") "
            Else
                ls_SQL = ls_SQL & "kanbanno1 = '', " & vbCrLf & _
                    "kanbanno2 = '', " & vbCrLf & _
                    "kanbanno3 = '', " & vbCrLf & _
                    "kanbanno4 = '', " & vbCrLf & _
                    "kanbantime1 = (SELECT ISNULL(Kanbantime,'10:00') FROM MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & Trim(pCycle1) & "'), " & vbCrLf & _
                    "kanbantime2 = (SELECT ISNULL(Kanbantime,'12:00') FROM MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & Trim(pCycle2) & "'), " & vbCrLf & _
                    "kanbantime3 = (SELECT ISNULL(Kanbantime,'15:00') FROM MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & Trim(pCycle3) & "'), " & vbCrLf & _
                    "kanbantime4 = (SELECT ISNULL(Kanbantime,'17:00') FROM MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & Trim(pCycle4) & "') " & vbCrLf
            End If

            ls_SQL = ls_SQL & "FROM dbo.PO_Master POM " & vbCrLf & _
                "LEFT JOIN dbo.po_detailUpload POD ON POM.PONo = POD.PoNo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID AND ISNULL(FinalApproveDate,'') <> '' " & vbCrLf & _
                "LEFT JOIN ( " & vbCrLf & _
                "SELECT MAX(seqno) seqno, pono, poRevNo, affiliateid, supplierid " & vbCrLf & _
                "FROM dbo.PORev_Master " & vbCrLf & _
                "WHERE FinalApproveUser <> '' " & vbCrLf & _
                "GROUP BY PONo, AffiliateID, SupplierID, PORevNo " & vbCrLf & _
                ")PRM ON POD.AffiliateID = PRM.AffiliateID AND PRM.PONo = POM.PONo AND PRM.SupplierID = POM.SupplierID " & vbCrLf & _
                "LEFT JOIN dbo.PORev_DetailUpload PRD ON PRD.PONo = PRM.PONo AND PRD.AffiliateID = PRM.AffiliateID AND PRD.SupplierID = PRM.SupplierID AND PRM.PORevNo = PRD.PORevNo AND PRD.PartNo = POD.PartNo " & vbCrLf & _
                "LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                "LEFT JOIN dbo.MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls " & vbCrLf & _
                "LEFT JOIN dbo.MS_SupplierCapacity MSS ON MSS.SupplierID = POD.SupplierID AND MSS.PartNo = POD.PartNo " & vbCrLf & _
                "LEFT JOIN dbo.MS_PartMapping PMP ON PMP.PartNo = POD.PartNo AND PMP.AffiliateID = POD.AffiliateID AND PMP.SupplierID = POD.SupplierID " & vbCrLf & _
                "WHERE (COALESCE(PRD.DeliveryD4, POD.DeliveryD4) - ( " & vbCrLf & _
                "Select ISNULL(SUM(qty), 0) " & vbCrLf & _
                "FROM ( " & vbCrLf & _
                "SELECT DISTINCT MAS.Kanbanno, KanbanQty AS qty " & vbCrLf & _
                "FROM dbo.Kanban_Detail DET " & vbCrLf & _
                "INNER JOIN Kanban_Master MAS ON MAS.AffiliateID = DET.AffiliateID AND MAS.KanbanNo = DET.KanbanNo AND MAS.SupplierID = DET.SupplierID AND MAS.DeliveryLocationCode = DET.DeliveryLocationCode " & vbCrLf & _
                "WHERE PONo = POM.PONo " & vbCrLf & _
                "AND PartNo = POD.PartNo " & vbCrLf & _
                "AND DET.AffiliateID = POM.AffiliateID " & vbCrLf & _
                "AND DET.SupplierID = POM.SupplierID " & vbCrLf & _
                "AND MAS.KanbanDate = '" & Format(pkanbandate, "yyyy-MM-dd") & "' " & vbCrLf & _
                ")x " & vbCrLf & _
                ")) > 0 " & vbCrLf & _
                "AND POD.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                "AND MONTH(POM.Period) = '" & Format(pkanbandate, "MM") & "' " & vbCrLf & _
                "AND YEAR(POM.Period) = '" & Format(pkanbandate, "yyyy") & "' " & vbCrLf

            If psuppID <> clsGlobal.gs_All And psuppID <> "" Then
                ls_SQL = ls_SQL & "AND POD.SupplierID = '" & Trim(psuppID) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL & ")DETAIL " & vbCrLf & _
                "*/)xx ORDER BY colno, colpono, colpartno "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
            End With
            Session.Remove("KNEW")
            sqlConn.Close()
        End Using
    End Sub

    Function getSeqno(ByVal kanbandate As Date) As String
        Dim ls_sql As String

        ls_sql = ""
        ls_sql = "select seqno = (count(kanbanno)/4) + 1 from  kanban_Master where kanbandate = '" & kanbandate & "' " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                getSeqno = ds.Tables(0).Rows(0)("seqno")
            End If

            sqlConn.Close()
        End Using

    End Function

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        ls_sql = "SELECT [Supplier Code] = RTRIM(supplierID) ,[Supplier Name] = RTRIM(SupplierName), ALAMAT = RTRIM(Address) + '  ' + RTRIM(City) + '  '+ RTRIM(PostalCode) FROM MS_Supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbosupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Supplier Code")
                .Columns(0).Width = 100
                .Columns.Add("Supplier Name")
                .Columns(1).Width = 240
                .Columns.Add("ALAMAT")
                .Columns(2).Width = 0

                .TextField = "Supplier Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Delivery Location
        ls_sql = "SELECT [Delivery Location Code] = RTRIM(DeliveryLocationCode) ,[Delivery Location Name] = RTRIM(DeliveryLocationName) FROM MS_DeliveryPlace where AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbolocation
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Delivery Location Code")
                .Columns(0).Width = 70
                .Columns.Add("Delivery Location Name")
                .Columns(1).Width = 240

                .TextField = "Delivery Location Code"
                .DataBind()
                '.SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        'SEQNO
        ls_sql = "select x='1-4' union all select x='5-8' union all select x='9-12' union all select x='13-16' union all select x='17-20'"
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboseq
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("x")
                .Columns(0).Width = 70

                .TextField = "SEQUENCE"
                .DataBind()
                '.SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        'TYPE
        ls_sql = "select x='NORMAL' union all select x='EMERGENCY'"
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbotype
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("x")
                .Columns(0).Width = 70

                .TextField = "TYPE"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_ApproveData()
        Dim ls_sql As String
        Dim status As String

        status = "nothing"
        ls_sql = ""
        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            ls_sql = "SELECT *, x = ISNULL((SELECT TOP 1 KB.KanbanNo FROM Kanban_Barcode KB WHERE KB.KanbanNo = Kanban_Master.KanbanNo AND KB.AffiliateID = Kanban_Master.AffiliateID AND KB.SupplierID = Kanban_Master.SupplierID), ''), y = ISNULL(AffiliateApproveUser,''), " & vbCrLf & _
                "z = ISNULL((SELECT TOP 1 SuratJalanNo FROM DOSupplier_Detail DO WHERE DO.KanbanNo = Kanban_Master.KanbanNo AND DO.AffiliateID = Kanban_Master.AffiliateID AND DO.SupplierID = Kanban_Master.SupplierID), '') " & vbCrLf & _
                "FROM dbo.Kanban_Master " & vbCrLf & _
                "WHERE AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                "AND SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                "AND Kanbanno = '" & Trim(Session("FilterKanbanNo")) & "' "

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            Using sqlTran As SqlTransaction = cn.BeginTransaction()
                Dim sqlComm As New SqlCommand(ls_sql, cn, sqlTran)

                If ds.Tables(0).Rows.Count > 0 Then

                    'sudah ada data DN Supplier
                    'If Trim(ds.Tables(0).Rows(0)("z")) <> "" Then
                    '    lblerrmessage.Text = "Can't resend DN, already exists DN Data from Supplier !"
                    '    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    '    status = "nothing"
                    '    Exit Sub
                    'End If

                    If Trim(ds.Tables(0).Rows(0)("x")) = "" Then
                        ls_sql = "UPDATE Kanban_Master " & vbCrLf & _
                            "SET AffiliateApproveUser = '" & Session("UserID").ToString & "', " & vbCrLf & _
                            "AffiliateApproveDate = GETDATE(), " & vbCrLf & _
                            "ExcelCls = '1' " & vbCrLf & _
                            "WHERE AffiliateID = '" & Session("AffiliateID") & " ' " & vbCrLf & _
                            "AND SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                            "AND KanbanNo = '" & Trim(Session("FilterKanbanNo")) & "'" & vbCrLf

                        ls_sql = ls_sql + " DECLARE @KanbanNo AS VARCHAR(25) , " & vbCrLf & _
                                          "     @SupplierID AS VARCHAR(10) , " & vbCrLf & _
                                          "     @SupplierName AS VARCHAR(100) , " & vbCrLf & _
                                          "     @DockID AS VARCHAR(20) , " & vbCrLf & _
                                          "     @PartNo AS VARCHAR(50) , " & vbCrLf & _
                                          "     @PartName AS VARCHAR(100) , " & vbCrLf & _
                                          "     @Qty AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @Cust AS VARCHAR(50) , " & vbCrLf & _
                                          "     @DeliveryDate AS VARCHAR(10) , " & vbCrLf & _
                                          "     @TIME AS VARCHAR(10) , " & vbCrLf & _
                                          "     @Location AS VARCHAR(50) , " & vbCrLf & _
                                          "     @AffCode As Varchar(20), " & vbCrLf

                        ls_sql = ls_sql + "     @DeliveryLocationCode AS VARCHAR(50) , " & vbCrLf & _
                                          "     @PONo AS VARCHAR(50) , " & vbCrLf & _
                                          "     @Barcode AS VARCHAR(1000) , " & vbCrLf & _
                                          "     @Barcode2 AS VARCHAR(1000) , " & vbCrLf & _
                                          "     @QtyBox AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @Loop AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @StartNo AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @EndNo AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @Total AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @PartNoSave AS CHAR(50) , " & vbCrLf & _
                                          "     @ETAAffiliate AS VARCHAR(10) , " & vbCrLf

                        ls_sql = ls_sql + "     @ETAPASI AS VARCHAR(10) , " & vbCrLf & _
                                          "     @BoxNo AS VARCHAR(10) , " & vbCrLf & _
                                          "     @AffiliateID AS VARCHAR(10) , " & vbCrLf & _
                                          "     @Cycle AS VARCHAR(3), " & vbCrLf & _
                                          "     @sequence AS Numeric(10,0), " & vbCrLf & _
                                          "     @LabelCode AS VARCHAR(10) " & vbCrLf & _
                                          "     SET @sequence = 0 " & vbCrLf & _
                                          "     SET @AffCode = (Select Top 1 Case When ISNULL(RTRIM(AffiliateCode),'') = '' Then '32G8' Else RTRIM(AffiliateCode) End from MS_Affiliate where AffiliateID = '" & Session("AffiliateID") & "') " & vbCrLf & _
                                          "   SELECT TOP 1 " & vbCrLf & _
                                          "             KanbanNo = CONVERT(CHAR(25), '') , " & vbCrLf & _
                                          "             AffiliateID = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             SupplierID = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             SupplierName = CONVERT(CHAR(100), '') , " & vbCrLf & _
                                          "             PartNo = CONVERT(CHAR(50), '') , " & vbCrLf & _
                                          "             PartName = CONVERT(CHAR(100), '') , " & vbCrLf

                        ls_sql = ls_sql + "             Qty = 0 , " & vbCrLf & _
                                          "             Cust = CONVERT(CHAR(50), '') , " & vbCrLf & _
                                          "             DeliveryDate = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             TIME = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             Location = CONVERT(CHAR(50), '') , " & vbCrLf & _
                                          "             PONo = CONVERT(CHAR(50), '') , " & vbCrLf & _
                                          "             Barcode2 = CONVERT(CHAR(1000), '') , " & vbCrLf & _
                                          "             qtybox = 0 , " & vbCrLf & _
                                          "             startno = 0 , " & vbCrLf & _
                                          "             EndNo = 0 , " & vbCrLf & _
                                          "             total = 0 , " & vbCrLf

                        ls_sql = ls_sql + "             DockID = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             DeliveryLocationCode = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             EtaPasi = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             EtaAffiliate = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             BoxNo = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             Barcode = CONVERT(CHAR(1000), '') , " & vbCrLf & _
                                          "             Cycle = CONVERT(CHAR(3), '') , " & vbCrLf & _
                                          "             LabelCode = CONVERT(VARCHAR(10), '') " & vbCrLf & _
                                          "   INTO      #data " & vbCrLf & _
                                          "   FROM      dbo.Kanban_Master KM   " & vbCrLf & _
                                          "      " & vbCrLf & _
                                          "   DELETE    FROM #data " & vbCrLf & _
                                          "   WHERE     KanbanNo = ''   " & vbCrLf

                        ls_sql = ls_sql + "   DECLARE cur_Print CURSOR FOR   " & vbCrLf & _
                                          "   SELECT  KM.KanbanNo AS kanbanNo ,Km.AffiliateID,   " & vbCrLf & _
                                          "   KM.SupplierID AS SupplierID ,   " & vbCrLf & _
                                          "   MSS.SupplierName AS SupplierName ,  KD.PartNo AS PartNo ,   " & vbCrLf & _
                                          "   MSP.PartName AS PartName ,   " & vbCrLf & _
                                          "   KD.KanbanQty Qty ,   " & vbCrLf & _
                                          "   KM.AffiliateID AS Cust ,   " & vbCrLf & _
                                          "   DeliveryDate = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(KM.Kanbandate,'')), 112) ,   " & vbCrLf & _
                                          "   CONVERT(CHAR(5), ISNULL(KM.kanbantime,'00:00:00')) AS TIME ,   " & vbCrLf

                        ls_sql = ls_sql + "   '' LocationID ,   " & vbCrLf & _
                                          "   KD.PONo ,   " & vbCrLf & _
                                          "   Barcode2 = @AffCode + ',' + RTRIM(KD.PONO) + ',' + RTRIM(KM.KanbanNo) + ',' + Rtrim(CONVERT(CHAR(10), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) + ',' + RTRIM(KD.PartNo) + ',' + RTRIM(MSP.PartCarMaker) + ',' + Replace(Rtrim(KD.POQtyBox),'.00','') , " & vbCrLf & _
                                          "   --Barcode2 = @AffCode + ',' + RTRIM(KD.PONO) + ',' + RTRIM(KM.KanbanNo) + ',' + Rtrim(CONVERT(CHAR(10), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) + ',' + RTRIM(KD.PartNo) + ',' + Replace(Rtrim(KD.POQtyBox),'.00','') , " & vbCrLf & _
                                          "   QtyBox = KD.POQtyBox , " & vbCrLf & _
                                          "   DockID = '', " & vbCrLf & _
                                          "   DeliveryLocationCode = ISNULL(KM.DeliveryLocationCode,''), " & vbCrLf & _
                                          "   ETAAffiliate = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(KM.Kanbandate,'')), 112) , " & vbCrLf

                        ls_sql = ls_sql + "   ETAPASI = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(ETDPASI,'')), 112) , " & vbCrLf & _
                                          "   BoxNo = '', " & vbCrLf & _
                                          "   Barcode = 'http://zxing.org/w/chart?cht=qr&chs=120x120&chld=L&choe=ISO-8859-1&chl=' + @AffCode + ',' + RTRIM(KD.PONO) + ',' + RTRIM(KM.KanbanNo) + ',' + Rtrim(CONVERT(CHAR(10), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) + ',' + RTRIM(KD.PartNo) + ',' + RTRIM(MSP.PartCarMaker) + ',' + Replace(Rtrim(KD.POQtyBox),'.00','')   " & vbCrLf & _
                                          "   --Barcode = 'http://zxing.org/w/chart?cht=qr&chs=120x120&chld=L&choe=ISO-8859-1&chl=' + @AffCode + ',' + RTRIM(KD.PONO) + ',' + RTRIM(KM.KanbanNo) + ',' + Rtrim(CONVERT(CHAR(10), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) + ',' + RTRIM(KD.PartNo) + ',' + Replace(Rtrim(KD.POQtyBox),'.00','')   " & vbCrLf & _
                                          "   ,Cycle = KM.KanbanCycle " & vbCrLf & _
                                          "   ,MSS.LabelCode " & vbCrLf & _
                                          "   FROM    dbo.Kanban_Master KM   " & vbCrLf & _
                                          "   LEFT JOIN dbo.Kanban_Detail KD ON KM.AffiliateID = KD.AffiliateID   " & vbCrLf & _
                                          "   AND KM.KanbanNo = KD.KanbanNo   " & vbCrLf & _
                                          "   AND KM.SupplierID = KD.SupplierID   " & vbCrLf & _
                                          "   --LEFT JOIN dbo.MS_DeliveryPlace MDP ON MDP.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
                                          "   LEFT JOIN dbo.MS_Supplier MSS ON MSS.SupplierID = KM.SupplierID   " & vbCrLf & _
                                          "   LEFT JOIN dbo.MS_Parts MSP ON MSP.PartNo = KD.PartNo   " & vbCrLf & _
                                          "   LEFT JOIN dbo.MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf

                        ls_sql = ls_sql + "   LEFT JOIN MS_ETD_PASI MEP ON MEP.AffiliateID = KM.AffiliateID " & vbCrLf & _
                                          "   AND CONVERT(CHAR(8), CONVERT(DATETIME, ETAAFfiliate),112) = CONVERT(CHAR(8), CONVERT(DATETIME, KanbanDate),112) " & vbCrLf & _
                                          "   WHERE KanbanQty <> 0 " & vbCrLf & _
                                          "   AND KD.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                                          "   AND KD.SupplierID  = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                                          "   AND KM.Kanbanno = '" & Trim(Session("FilterKanbanNo")) & "' " & vbCrLf

                        ls_sql = ls_sql + " OPEN cur_Print   " & vbCrLf & _
                                          "   FETCH NEXT FROM cur_Print   " & vbCrLf & _
                                          "      INTO @KanbanNo, @AffiliateID, @SupplierID, @SupplierName, @PartNo, @PartName,   " & vbCrLf & _
                                          "   		@Qty, @Cust, @DeliveryDate, @TIME, @Location, @PONo, @Barcode2, @QtyBox, @DockID,  " & vbCrLf & _
                                          "   		@DeliveryLocationCode, @ETAAffiliate, @ETAPASI, @BoxNo, @Barcode, @Cycle, @LabelCode " & vbCrLf & _
                                          "      " & vbCrLf & _
                                          "   WHILE @@Fetch_Status = 0  " & vbCrLf & _
                                          "     BEGIN   " & vbCrLf & _
                                          "         SET @StartNo = 0 " & vbCrLf & _
                                          "         SET @total = 0   " & vbCrLf & _
                                          "         WHILE @Total < @Qty  " & vbCrLf

                        ls_sql = ls_sql + "             BEGIN   " & vbCrLf & _
                                          "                 BEGIN   " & vbCrLf & _
                                          "                      SET @StartNo = @StartNo + 1  " & vbCrLf & _
                                          "                      SET @sequence = @sequence + 1      " & vbCrLf & _
                                          "                      INSERT  INTO #Data  " & vbCrLf & _
                                          "                      VALUES  ( @KanbanNo, @AffiliateID, @SupplierID,  " & vbCrLf & _
                                          "                                @SupplierName, @PartNo, @PartName, @Qty, @Cust,  " & vbCrLf & _
                                          "                                @DeliveryDate, @Time, @Location, @PONo,  " & vbCrLf & _
                                          "                                ( Rtrim(@Barcode2) + ',' + @LabelCode + RIGHT(RTRIM('00000' + REPLACE(CONVERT(NUMERIC, @sequence), '.00', '')), 5))   " & vbCrLf & _
                                          "                               , @QtyBox,  " & vbCrLf & _
                                          "                                RTRIM(CONVERT(NUMERIC, @StartNo)),  " & vbCrLf & _
                                          "                                RTRIM(CONVERT(NUMERIC, ( @Qty / @QtyBox ))),  "

                        ls_sql = ls_sql + "                                ( @Qty / @QtyBox ), @dockID,  " & vbCrLf & _
                                          "                                @DeliveryLocationCode, @ETAPASI, @EtaAffiliate,  " & vbCrLf & _
                                          "                                @LabelCode + RIGHT(RTRIM('00000' + REPLACE(CONVERT(NUMERIC, @sequence), '.00', '')), 5),  " & vbCrLf & _
                                          "                                Rtrim(@Barcode) + ',' + @LabelCode + RIGHT(RTRIM('00000' + REPLACE(CONVERT(NUMERIC, @sequence), '.00', '')), 5) , @Cycle, @LabelCode )  " & vbCrLf & _
                                          "                      SET @Total = @Total + @QtyBox   " & vbCrLf & _
                                          "                 END   " & vbCrLf & _
                                          "             END " & vbCrLf & _
                                          "         FETCH NEXT FROM cur_Print   " & vbCrLf & _
                                          "   		 INTO @KanbanNo, @AffiliateID, @SupplierID, @SupplierName, @PartNo, @PartName,   " & vbCrLf & _
                                          "   			@Qty, @Cust, @DeliveryDate, @TIME, @Location, @PONo, @Barcode2, @QtyBox,@DockID,  " & vbCrLf & _
                                          "   		@DeliveryLocationCode, @ETAAffiliate, @ETAPASI, @BoxNo, @Barcode, @Cycle, @LabelCode   " & vbCrLf

                        ls_sql = ls_sql + "      " & vbCrLf & _
                                          "     END   " & vbCrLf & _
                                          "   CLOSE cur_Print   " & vbCrLf & _
                                          "   DEALLOCATE cur_Print    " & vbCrLf & _
                                          "   INSERT    INTO Kanban_Barcode " & vbCrLf & _
                                          "             SELECT  AffiliateID , " & vbCrLf & _
                                          "                     SupplierID , " & vbCrLf & _
                                          "                     DockID , " & vbCrLf & _
                                          "                     Location , " & vbCrLf & _
                                          "                     EtaAffiliate , " & vbCrLf & _
                                          "                     ETAPasi , " & vbCrLf

                        ls_sql = ls_sql + "                     POno , " & vbCrLf &
                                          "                     KanbanNo , " & vbCrLf &
                                          "                     Cycle , " & vbCrLf &
                                          "                     partNo , " & vbCrLf &
                                          "                     BoxNo,  " & vbCrLf &
                                          "                     Startno , " & vbCrLf &
                                          "                     EndNo , " & vbCrLf &
                                          "                     QtyBox , " & vbCrLf &
                                          "                     Barcode , " & vbCrLf &
                                          "                     DeliveryLocationCode , " & vbCrLf &
                                          "                     Barcode2 "

                        ls_sql = ls_sql + "             FROM    #data    " & vbCrLf

                        ls_sql = ls_sql + "             " & vbCrLf & _
                                          "             DROP TABLE  #data    " & vbCrLf
                    ElseIf Trim(ds.Tables(0).Rows(0)("y")) = "" Then
                        ls_sql = "UPDATE Kanban_Master " & vbCrLf & _
                            "SET AffiliateApproveUser = '" & Session("UserID").ToString & "', " & vbCrLf & _
                            "AffiliateApproveDate = GETDATE(), " & vbCrLf & _
                            "ExcelCls = '1' " & vbCrLf & _
                            "WHERE AffiliateID = '" & Session("AffiliateID") & " ' " & vbCrLf & _
                            "AND SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                            "AND KanbanNo = '" & Trim(Session("FilterKanbanNo")) & "'" & vbCrLf
                    Else
                        ls_sql = "UPDATE Kanban_Master " & vbCrLf & _
                            "SET ExcelCls = '1' " & vbCrLf & _
                            "WHERE AffiliateID = '" & Session("AffiliateID") & " ' " & vbCrLf & _
                            "AND SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf & _
                            "AND KanbanNo = '" & Trim(Session("FilterKanbanNo")) & "'" & vbCrLf
                    End If

                    sqlComm = New SqlCommand(ls_sql, cn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    Call clsMsg.DisplayMessage(lblerrmessage, "1006", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Session("msgapprove") = lblerrmessage.Text
                    status = "approve"

                    sqlComm.Dispose()
                    sqlTran.Commit()

                    xKanbanNo = Session("xKanbanNo")
                Else
                    'data ga ada
                    Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    status = "nothing"
                End If
            End Using
        End Using

        Call fillHeaderAfterApprove()

        If status = "approve" Then
            Call clsMsg.DisplayMessage(lblerrmessage, "1006", clsMessage.MsgType.InformationMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        ElseIf status = "unapprove" Then
            Call clsMsg.DisplayMessage(lblerrmessage, "1009", clsMessage.MsgType.InformationMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        ElseIf status = "notupdate" Then
            Call clsMsg.DisplayMessage(lblerrmessage, "6027", clsMessage.MsgType.ErrorMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Else
            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        End If
        Session("msgapprove") = lblerrmessage.Text
    End Sub
#End Region

#Region "Excel"
    Private Sub Excel()
        Dim strFileSize As String = ""

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer

        Dim fileTocopyKanban As String
        Dim NewFileCopyKanban As String
        Dim fileTocopyDelivery As String
        Dim NewFileCopyDelivery As String

        Dim lspath As String

        On Error GoTo ErrHandler

        lspath = "D:\Template\"
        fileTocopyKanban = Server.MapPath("~\Template\Template Kanban.xlsm")
        NewFileCopyKanban = lspath & "Template Kanban.xlsm"

        'copy file from server to local
        If System.IO.File.Exists(fileTocopyKanban) = True Then
            System.IO.File.Delete(NewFileCopyKanban)
            System.IO.File.Copy(fileTocopyKanban, NewFileCopyKanban)
        Else
            System.IO.File.Copy(fileTocopyKanban, NewFileCopyKanban)
        End If
        'copy file from server to local

        Dim xlApp = New Excel.Application
        Dim ls_file As String
        ls_file = NewFileCopyKanban
        ExcelBook = xlApp.Workbooks.Open(ls_file)
        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

        If Grid.VisibleRowCount > 0 Then
            Dim ls_sql As String
            Dim CCemail As String
            Dim receiptemail As String
            Dim x As Long

            ls_sql = ""
            CCemail = ""
            receiptemail = ""

            ls_sql = " select kanbanCC =  Rtrim(convert(char,kanbanCC)) + ';' + Rtrim(convert(char,KanbanTo)),kanbanTo = '' from MS_emailAffiliate where AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                 " union ALL " & vbCrLf & _
                 " select kanbanCC, kanbanTo = '' from MS_EmailPasi  " & vbCrLf & _
                 " UNION ALL " & vbCrLf & _
                 " select kanbanCC, kanbanTo = kanbanTo from MS_EmailSupplier where supplierID = '" & Trim(cbosupplier.Text) & "' "

            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                receiptemail = ds.Tables(0).Rows(2)("kanbanTo")
                For x = 0 To ds.Tables(0).Rows.Count - 1
                    If CCemail = "" Then
                        CCemail = ds.Tables(0).Rows(x)("kanbanCC")
                    Else
                        CCemail = CCemail & ";" & ds.Tables(0).Rows(x)("kanbanCC")
                    End If
                Next
            End Using

            ExcelSheet.Range("H2").Value = receiptemail
            ExcelSheet.Range("Y2").Value = CCemail
            ExcelSheet.Range("H3").Value = Session("AffiliateID")
            ExcelSheet.Range("H4").Value = Trim(cbolocation.Text)
            ExcelSheet.Range("H5").Value = Trim(cbosupplier.Text)

            ExcelSheet.Range("I9").Value = Session("KAffiliateName")
            ExcelSheet.Range("I11").Value = Trim(txtlocation.Text)
            ExcelSheet.Range("I13").Value = Trim(txtsuppliername.Text)
            ExcelSheet.Range("I18").Value = Format(dtkanban.Value, "dd MMM yyyy")

            ExcelSheet.Range("AE35").Value = Trim(txtkanban1.Text)
            ExcelSheet.Range("AI35").Value = Trim(txtkanban2.Text)
            ExcelSheet.Range("AM35").Value = Trim(txtkanban3.Text)
            ExcelSheet.Range("AQ35").Value = Trim(txtkanban4.Text)

            ExcelSheet.Range("AE36").Value = Trim(txttime1.Text)
            ExcelSheet.Range("AI36").Value = Trim(txttime2.Text)
            ExcelSheet.Range("AM36").Value = Trim(txttime3.Text)
            ExcelSheet.Range("AQ36").Value = Trim(txttime4.Text)


            For i = 0 To Grid.VisibleRowCount - 1
                If Grid.GetRowValues(i, "cols") = "1" Then
                    'ExcelSheet.Range("B" & i + 39 & ": C" & i + 39).Merge()
                    'ExcelSheet.Range("D" & i + 39 & ": H" & i + 39).Merge()
                    'ExcelSheet.Range("i" & i + 39 & ": P" & i + 39).Merge()
                    'ExcelSheet.Range("Q" & i + 39 & ": V" & i + 39).Merge()
                    'ExcelSheet.Range("W" & i + 39 & ": Y" & i + 39).Merge()
                    'ExcelSheet.Range("Y" & i + 39 & ": Z" & i + 39).Merge()
                    'ExcelSheet.Range("AA" & i + 39 & ": AD" & i + 39).Merge()
                    'ExcelSheet.Range("AE" & i + 39 & ": AH" & i + 39).Merge()
                    'ExcelSheet.Range("AI" & i + 39 & ": AL" & i + 39).Merge()
                    'ExcelSheet.Range("AM" & i + 39 & ": AP" & i + 39).Merge()
                    'ExcelSheet.Range("AQ" & i + 39 & ": AT" & i + 39).Merge()
                    'ExcelSheet.Range("AU" & i + 39 & ": AW" & i + 39).Merge()

                    ExcelSheet.Range("B" & i + 39 & ": C" & i + 39).Value = Trim(Grid.GetRowValues(i, "colno"))
                    ExcelSheet.Range("D" & i + 39 & ": H" & i + 39).Value = Trim(Grid.GetRowValues(i, "colpartno"))
                    ExcelSheet.Range("i" & i + 39 & ": P" & i + 39).Value = Trim(Grid.GetRowValues(i, "coldescription"))
                    ExcelSheet.Range("Q" & i + 39 & ": V" & i + 39).Value = Trim(Grid.GetRowValues(i, "colpono"))
                    ExcelSheet.Range("W" & i + 39 & ": Y" & i + 39).Value = Trim(Grid.GetRowValues(i, "coluom"))
                    ExcelSheet.Range("Y" & i + 39 & ": Z" & i + 39).Value = Trim(Grid.GetRowValues(i, "colqty"))
                    ExcelSheet.Range("AA" & i + 39 & ": AD" & i + 39).Value = Trim(Grid.GetRowValues(i, "colkanbanqty"))
                    ExcelSheet.Range("AE" & i + 39 & ": AH" & i + 39).Value = Trim(Grid.GetRowValues(i, "colcycle1"))
                    ExcelSheet.Range("AI" & i + 39 & ": AL" & i + 39).Value = Trim(Grid.GetRowValues(i, "colcycle2"))
                    ExcelSheet.Range("AM" & i + 39 & ": AP" & i + 39).Value = Trim(Grid.GetRowValues(i, "colcycle3"))
                    ExcelSheet.Range("AQ" & i + 39 & ": AT" & i + 39).Value = Trim(Grid.GetRowValues(i, "colcycle4"))
                    ExcelSheet.Range("AU" & i + 39 & ": AW" & i + 39).Value = Trim(Grid.GetRowValues(i, "colbox"))

                End If
            Next
            ExcelSheet.Range("B40").Interior.Color = Color.White
            ExcelSheet.Range("B40").Font.Color = Color.Black
            ExcelSheet.Range("B" & i + 39).Value = "E"
            ExcelSheet.Range("B" & i + 39).Interior.Color = Color.Black
            ExcelSheet.Range("B" & i + 39).Font.Color = Color.White

            DrawAllBorders(ExcelSheet.Range("B39" & ": AW" & i + 38))

            'System.IO.File.Delete(lspath & "Kanban.xlsm")
            ExcelBook.SaveAs(lspath & "Kanban.xlsm")
            ExcelBook.SaveAs(Server.MapPath("~\Result\Kanban.xlsm"))
            'System.IO.File.Delete(NewFileCopyKanban)


            'Dim fServer As String
            'Dim fLocal As String

            'fServer = Server.MapPath("~\Result\Kanban.xlsm")
            'fLocal = "D:\Template"
            ''copy file from server to local
            'If System.IO.File.Exists(fLocal) = True Then
            '    System.IO.File.Copy(fLocal, fServer)
            'Else
            '    System.IO.File.Copy(fLocal, fServer)
            'End If
            ''copy file from server to local

            xlApp.Workbooks.Close()
            xlApp.Quit()

            '----------------------------------------------------------------------------------------------------------
            'copy file from server to local
            fileTocopyDelivery = Server.MapPath("~\Template\Template Delivery.xlsm")
            NewFileCopyDelivery = lspath & "Template Delivery.xlsm"

            If System.IO.File.Exists(fileTocopyDelivery) = True Then
                System.IO.File.Delete(NewFileCopyDelivery)
                System.IO.File.Copy(fileTocopyDelivery, NewFileCopyDelivery)
            Else
                System.IO.File.Copy(fileTocopyDelivery, NewFileCopyDelivery)
            End If
            'copy file from server to local


            ls_file = NewFileCopyDelivery
            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)
            Dim j As Long
            Dim jKanbanno As String
            i = 0
            j = 0

            ExcelSheet.Range("H2").Value = receiptemail
            ExcelSheet.Range("Y2").Value = CCemail
            ExcelSheet.Range("H3").Value = Session("AffiliateID")
            ExcelSheet.Range("H4").Value = Trim(cbolocation.Text)
            ExcelSheet.Range("H5").Value = Trim(cbosupplier.Text)

            'ExcelSheet.Range("I11:X11").Merge()
            ExcelSheet.Range("I11").Value = Trim(txtsuppliername.Text)
            'ExcelSheet.Range("I12:X12").Merge()
            'ExcelSheet.Range("I12").Value = Trim(txtaddress.Text)
            'ExcelSheet.Range("I16:X16").Merge()
            ExcelSheet.Range("I16").Value = Session("KAffiliateName")
            'ExcelSheet.Range("I17:X19").Merge()
            ExcelSheet.Range("I17").Value = Session("KAlamat")



            For j = 0 To Grid.VisibleRowCount - 1
                For i = 0 To 3
                    If i = 0 Then jKanbanno = Trim(txtkanban1.Text)
                    If i = 1 Then jKanbanno = Trim(txtkanban2.Text)
                    If i = 2 Then jKanbanno = Trim(txtkanban3.Text)
                    If i = 3 Then jKanbanno = Trim(txtkanban4.Text)

                    If Grid.GetRowValues(j, "cols") = "1" Then
                        'ExcelSheet.Range("B" & i + 39 & ": C" & i + 39).Merge()
                        'ExcelSheet.Range("D" & i + 39 & ": H" & i + 39).Merge()
                        'ExcelSheet.Range("i" & i + 39 & ": K" & i + 39).Merge()
                        'ExcelSheet.Range("L" & i + 39 & ": O" & i + 39).Merge()
                        'ExcelSheet.Range("P" & i + 39 & ": T" & i + 39).Merge()
                        'ExcelSheet.Range("U" & i + 39 & ": AC" & i + 39).Merge()
                        'ExcelSheet.Range("AD" & i + 39 & ": AE" & i + 39).Merge()
                        'ExcelSheet.Range("AF" & i + 39 & ": AG" & i + 39).Merge()
                        'ExcelSheet.Range("AH" & i + 39 & ": AJ" & i + 39).Merge()
                        'ExcelSheet.Range("AK" & i + 39 & ": AN" & i + 39).Merge()
                        'ExcelSheet.Range("AO" & i + 39 & ": AR" & i + 39).Merge()
                        'ExcelSheet.Range("AS" & i + 39 & ": AV" & i + 39).Merge()
                        'ExcelSheet.Range("AW" & i + 39 & ": AZ" & i + 39).Merge()

                        ExcelSheet.Range("B" & i + 39 & ": C" & i + 39).Value = i + 1
                        ExcelSheet.Range("D" & i + 39 & ": H" & i + 39).Value = Trim(Grid.GetRowValues(j, "colpono"))
                        'ExcelSheet.Range("i" & i + 39 & ": K" & i + 39).Value = Trim(Grid.GetRowValues(i, "colpokanban"))
                        ExcelSheet.Range("L" & i + 39 & ": O" & i + 39).Value = jKanbanno
                        ExcelSheet.Range("P" & i + 39 & ": T" & i + 39).Value = Trim(Grid.GetRowValues(j, "colpartno"))
                        ExcelSheet.Range("U" & i + 39 & ": AC" & i + 39).Value = Trim(Grid.GetRowValues(j, "coldescription"))
                        ExcelSheet.Range("AD" & i + 39 & ": AE" & i + 39).Value = Trim(Grid.GetRowValues(j, "coluom"))
                        ExcelSheet.Range("AF" & i + 39 & ": AG" & i + 39).Value = Trim(Grid.GetRowValues(j, "colqty"))
                        ExcelSheet.Range("AH" & i + 39 & ": AJ" & i + 39).Value = "0"
                        ExcelSheet.Range("AK" & i + 39 & ": AN" & i + 39).Value = Trim(Grid.GetRowValues(j, "colkanbanqty"))
                        ExcelSheet.Range("AO" & i + 39 & ": AR" & i + 39).Value = Trim(Grid.GetRowValues(j, "colkanbanqty"))
                        ExcelSheet.Range("AS" & i + 39 & ": AV" & i + 39).Value = ""
                        ExcelSheet.Range("AW" & i + 39 & ": AZ" & i + 39).Value = ""

                    End If
                Next
            Next
            ExcelSheet.Range("B40").Interior.Color = Color.White
            ExcelSheet.Range("B40").Font.Color = Color.Black
            ExcelSheet.Range("B" & i + 39).Value = "E"
            ExcelSheet.Range("B" & i + 39).Interior.Color = Color.Black
            ExcelSheet.Range("B" & i + 39).Font.Color = Color.White

            DrawAllBorders(ExcelSheet.Range("B39" & ": AZ" & i + 38))

            'System.IO.File.Delete(lspath & "Delivery.xlsm")
            ExcelBook.SaveAs(lspath & "Delivery.xlsm")
            ExcelBook.SaveAs(Server.MapPath("~\Result\Delivery.xlsm"))
            'System.IO.File.Delete(NewFileCopyDelivery)

            ''copy file from server to local
            'fServer = Server.MapPath("~\Result\Delivery.xlsm")
            'fLocal = "D:\Template"
            'If System.IO.File.Exists(fLocal) = True Then
            '    System.IO.File.Copy(fLocal, fServer)
            'Else
            '    System.IO.File.Copy(fLocal, fServer)
            'End If
            ''copy file from server to local

            xlApp.Workbooks.Close()

            xlApp.Quit()
            '----------------------------------------------------------------------------------------------------------

            'Call sendEmail()
        End If
        Exit Sub
ErrHandler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
        xlApp.Workbooks.Close()
        xlApp.Quit()
        Kill(NewFileCopyDelivery)
        Kill(NewFileCopyKanban)
    End Sub

    Private Sub sendEmail()
        Dim TempFilePath As String
        Dim TempFileName1 As String
        Dim TempFileName2 As String
        Dim FileExtStr As String
        Dim OutApp As Object
        Dim OutMail As Object
        Dim receiptEmail As String
        Dim ls_sql As String
        Dim j As Long
        Dim CCemail As String
        Dim fromemail As String
        Dim cemail1 As String = ""
        Dim cemail2 As String = ""
        Dim cemail3 As String = ""
        Dim cemail4 As String = ""

        CCemail = ""
        receiptEmail = ""
        fromemail = ""

        TempFilePath = "D:\Template\"
        TempFileName1 = "Kanban.xlsm"
        TempFileName2 = "Delivery.xlsm"

        On Error Resume Next
        Dim smtpserver As New SmtpClient
        Dim smtp As New SmtpClient
        Dim mail As New MailMessage
        Dim fileAtt1 As String = ""
        Dim fileAtt2 As String = ""

        fileAtt1 = TempFilePath & TempFileName1
        fileAtt2 = TempFilePath & TempFileName2

        ls_sql = " select kanbanCC = isnull(kanbanCC,'') ,kanbanTo = '', kanbanFrom = isnull(KanbanTo,'') from MS_emailAffiliate where AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                 " union ALL " & vbCrLf & _
                 " select kanbanCC = isnull(kanbanCC,'') , kanbanTo = '', kanbanFrom = '' from MS_EmailPasi  " & vbCrLf & _
                 " UNION ALL " & vbCrLf & _
                 " select kanbanCC = isnull(kanbanCC,'') , kanbanTo = isnull(kanbanTo,''), KanbanFrom = '' from MS_EmailSupplier where supplierID = '" & Trim(cbosupplier.Text) & "' "

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()
            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            fromemail = Trim(ds.Tables(0).Rows(0)("KanbanFrom"))
            receiptEmail = ds.Tables(0).Rows(2)("kanbanTo")
            Dim Email As New Mail.MailMessage(fromemail, receiptEmail)
            For j = 0 To ds.Tables(0).Rows.Count - 1
                CCemail = ds.Tables(0).Rows(j)("kanbanCC")
                If CCemail <> "" Then
                    Dim mailAddressCC As New MailAddress(CCemail)
                    Email.CC.Add(mailAddressCC)
                End If
            Next


            Email.Subject = "Kanban Template Testing"
            Email.Body = "test email via smtp"
            Dim filename1 As String = TempFilePath & TempFileName1
            Dim filename2 As String = TempFilePath & TempFileName2
            Email.Attachments.Add(New Attachment(filename1))
            Email.Attachments.Add(New Attachment(filename2))
            Email.IsBodyHtml = False
            smtp.Host = "smtp.atisicloud.com"
            'smtp.Host = "mail.tos.co.id"
            smtp.EnableSsl = False
            smtp.UseDefaultCredentials = True
            smtp.Port = 25
            If receiptEmail = "" Then
                MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
                Exit Sub
            Else
                smtp.Send(Email)
            End If

            ''Delete the file
            'Kill(filename1)
            'Kill(filename2)

        End Using

    End Sub

    Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
        With Rg
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
    End Sub

#End Region

    Private Sub Approve_Callback(source As Object, e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles Approve.Callback
        up_ApproveData()
        fillHeaderAfterApprove()

        Call up_GridLoad(dtkanban.Value, cbosupplier.Text, cboseq.Text)

        lblerrmessage.Text = Session("msgapprove")
        Approve.JSProperties("cpMessage") = Session("msgapprove")

        If Left(Session("msgapprove"), 5) = "[1006" Then
            Approve.JSProperties("cpButton") = "UNAPPROVE"
        ElseIf Left(Session("msgapprove"), 5) = "[1009" Then
            Approve.JSProperties("cpButton") = "APPROVE"
        End If

        Session("KNEW") = "FALSE"
        Session("ALREADY") = "YES"
    End Sub
End Class