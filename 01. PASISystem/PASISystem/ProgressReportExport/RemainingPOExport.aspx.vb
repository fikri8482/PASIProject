Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO

Public Class RemainingPOExport
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = ""

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "O01"
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim dtExcel As DataTable

#End Region

#Region "CONTROL EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_FillCombo()
                dtPeriodFrom.Text = Format(Now, "yyyy-MM")
                dtPeriodTo.Text = Format(Now, "yyyy-MM")
                rdrRAll.Checked = True
                rdrEAll.Checked = True
                lblInfo.Text = ""
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("G01Msg")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If
                Case "excel"
                    GetExcel()
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

#End Region
    
#Region "PROCEDURE"

    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        'Combo Affiliate
        With cboAffiliate
            ls_SQL = "SELECT AffiliateID = '==ALL==', AffiliateName = '==ALL=='" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     "SELECT AffiliateID = RTRIM(AffiliateID), AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 90
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240
                .TextField = "AffiliateID"
                .DataBind()
                .Text = "==ALL=="
                txtAffiliate.Text = "==ALL=="
            End Using
        End With

        'Combo Supplier
        With cbosupplier
            ls_SQL = "SELECT SupplierID = '==ALL==', SupplierName = '==ALL=='" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     "SELECT SupplierID = RTRIM(SupplierID), SupplierName = RTRIM(SupplierName) FROM dbo.MS_Supplier"
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
                .Columns(1).Width = 120

                .TextField = "SupplierID"
                .DataBind()
                .Text = "==ALL=="
                txtsupplier.Text = "==ALL=="
            End Using
        End With

        'Combo Parts
        With cboPart
            ls_SQL = "SELECT PartNo = '==ALL==', PartName = '==ALL=='" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     "SELECT PartNo = RTRIM(PartNo), PartName = RTRIM(PartName) FROM dbo.MS_Parts"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 90
                .Columns.Add("PartName")
                .Columns(1).Width = 240

                .TextField = "PartNo"
                .DataBind()
                .Text = "==ALL=="
                txtPartName.Text = "==ALL=="
            End Using
        End With
    End Sub

    Private Sub up_GridLoad()
        Dim ls_sql As String = ""
        Dim ls_Filter As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            If rdrRYes.Value = True Then
                ls_Filter = ls_Filter + " AND ISNULL(Remaining,0) > 0 " & vbCrLf
            ElseIf rdrRNo.Value = True Then
                ls_Filter = ls_Filter + " AND ISNULL(Remaining,0) = 0 " & vbCrLf
            End If

            If cboPart.Text <> "==ALL==" And cboPart.Text <> "" Then
                ls_Filter = ls_Filter + " AND PartNo = '" & Trim(cboPart.Text) & "' "
            End If

            If cbosupplier.Text <> "==ALL==" And cbosupplier.Text <> "" Then
                ls_Filter = ls_Filter + " AND SupplierID = '" & Trim(cbosupplier.Text) & "' "
            End If

            If cboAffiliate.Text <> "==ALL==" And cboAffiliate.Text <> "" Then
                ls_Filter = ls_Filter + " AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' "
            End If

            If txtpono.Text <> "==ALL==" And txtpono.Text <> "" Then
                ls_Filter = ls_Filter + " AND OrderNo Like '%" & Trim(txtpono.Text) & "%' "
            End If

            If rdrEyes.Value = True Then
                ls_Filter = ls_Filter + " AND EmergencyCls = 'M' " & vbCrLf
            ElseIf rdrENo.Value = True Then
                ls_Filter = ls_Filter + " AND EmergencyCls = 'E' " & vbCrLf
            End If

            If txtboxno.Text <> "" Then
                ls_Filter = ls_Filter + "  AND '" & Microsoft.VisualBasic.Right(Trim(txtboxno.Text), 7) & "' between Right(rtrim(BoxNo),7) and Right(rtrim(boxNo),7)  " & vbCrLf
            End If

            ls_sql = " Select No = ROW_NUMBER() over (order by PartNo,OrderNo, SupplierID, AffiliateID), * from( " & vbCrLf & _
                  " Select distinct  " & vbCrLf & _
                  "   --SuratJalanno = isnull(RM.SuratJalanno,''),   " & vbCrLf & _
                  "   Period = Right(Convert(char(11),Convert(datetime,POM.Period),106),8),   " & vbCrLf & _
                  "   OrderNo = POM.OrderNo1,    " & vbCrLf & _
                  "   AffiliateID = POM.AffiliateID,    " & vbCrLf & _
                  "   AffiliateName = AffiliateName,   " & vbCrLf & _
                  "   ETDVendor = Convert(char(11),convert(datetime,POM.ETDVendor1),106),   " & vbCrLf & _
                  "   ETDPort =  Convert(char(11),convert(datetime,POM.ETDPort1),106),   " & vbCrLf & _
                  "   ETAPort =  Convert(char(11),convert(datetime,POM.ETAPORT1),106),    ETAFactory =  Convert(char(11),convert(datetime,POM.ETAFACTORY1),106),   " & vbCrLf & _
                  "   SupplierID = POM.SupplierID,   " & vbCrLf

            ls_sql = ls_sql + "   SupplierName = MS.SupplierName,   " & vbCrLf & _
                              "   PartNo = POD.PartNo,    " & vbCrLf & _
                              "   PartName = MP.PartName,    " & vbCrLf & _
                              "   UOM = isnull(MU.Description,''),    " & vbCrLf & _
                              "   QtyBox = ISNULL(POD.POQtyBox,MPM.QtyBox),    " & vbCrLf & _
                              "   BoxNo = (Case when isnull(RD.suratjalanno,'') <> '' then isnull(Rtrim(RB.label1) + '-' + Rtrim(RB.label2),'')  " & vbCrLf & _
                              " 			when Isnull(DSM.SuratJalanNo,'') <> '' then isnull(Rtrim(DSB1.BoxNo) + '-' + Rtrim(DSB1.BoxNoMax),'')  " & vbCrLf & _
                              " 			Else isnull(Rtrim(PL1.BoxNo) + '-' + Rtrim(PL1.BoxNoMax),'') END), " & vbCrLf & _
                              "   --SJNo = Isnull(DSM.SuratJalanNo,''), " & vbCrLf & _
                              "   POQty = Replace(Week1,'.00',''),   " & vbCrLf & _
                              "   DOQty = Replace(isnull(Doqty,0),'.00',''), " & vbCrLf

            ls_sql = ls_sql + "   GoodRecQty = Replace(isnull(GoodrecQty,0),'.00',''),   " & vbCrLf & _
                              "   DefectRecQty = 0,   " & vbCrLf & _
                              "   Remaining = Replace((isnull(Doqty,0) - (isnull(Rb.Box,0)*ISNULL(POD.POQtyBox,MPM.QtyBox))),'.0',''),   " & vbCrLf & _
                              "   EmergencyCls = Case when POM.EmergencyCls = 'M' then 'No' else 'Yes' END,  " & vbCrLf & _
                              "   BoxQty = RB.box " & vbCrLf & _
                              " From PO_Master_Export POM LEFT JOIN PO_Detail_Export POD " & vbCrLf & _
                              " ON POM.PONo = POD.PONo " & vbCrLf & _
                              " AND POM.OrderNo1 = POD.OrderNo1 " & vbCrLf & _
                              " AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " LEFT JOIN (Select PONo,OrderNo,AffiliateID,SupplierID,PartNo,BoxNo = Min(LabelNo),BoxNoMax = Max(LabelNo)  " & vbCrLf

            ls_sql = ls_sql + " 			from PrintLabelExport group By PONo,OrderNo,AffiliateID,SupplierID,PartNo) PL1 ON PL1.PONo = POM.PONo " & vbCrLf & _
                              " AND PL1.OrderNo = POM.OrderNo1  " & vbCrLf & _
                              " AND PL1.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " AND PL1.SupplierID = POM.SupplierID " & vbCrLf & _
                              " AND PL1.PartNo = POD.PartNo " & vbCrLf & _
                              " /*LEFT JOIN (Select PONo,OrderNo,AffiliateID,SupplierID,PartNo,BoxNo = Max(LabelNo)  " & vbCrLf & _
                              " 			from PrintLabelExport group By PONo,OrderNo,AffiliateID,SupplierID,PartNo) PL2 ON PL2.PONo = POM.PONo " & vbCrLf & _
                              " AND PL2.OrderNo = POM.OrderNo1  " & vbCrLf & _
                              " AND PL2.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " AND PL2.SupplierID = POM.SupplierID " & vbCrLf & _
                              " AND PL2.PartNo = POD.PartNo*/ "

            ls_sql = ls_sql + " LEFT JOIN DOSupplier_Master_Export DSM ON DSM.PONo = POM.PONo " & vbCrLf & _
                              " AND DSM.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " AND DSM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " AND DSM.SupplierID = POM.SupplierID " & vbCrLf & _
                              " LWFT JOIN DOSupplier_Detail_Export DSD ON DSD.Suratjalanno = DSM.Suratjalanno " & vbCrLf & _
                              " AND DSD.PONo = DSM.PONo " & vbCrLf & _
                              " AND DSD.OrderNo = DSM.OrderNo " & vbCrLf & _
                              " AND DSD.AffiliateID = DSM.AffiliateID " & vbCrLf & _
                              " AND DSD.SupplierID = DSM.SupplierID " & vbCrLf & _
                              " AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " LEFT JOIN (Select Suratjalanno,PONo,OrderNo,AffiliateID,SupplierID,PartNo,BoxNo = Min(boxNo),BoxNoMax = Max(boxNo)  " & vbCrLf

            ls_sql = ls_sql + " 			from DOSupplier_DetailBox_Export Group By Suratjalanno,PONo,OrderNo,AffiliateID,SupplierID,PartNo) DSB1 ON DSB1.SuratJalanNo = DSD.Suratjalanno " & vbCrLf & _
                              " AND DSB1.PONo = DSD.PONo " & vbCrLf & _
                              " AND DSB1.OrderNo = DSD.OrderNo " & vbCrLf & _
                              " AND DSB1.AffiliateID = DSD.AffiliateID " & vbCrLf & _
                              " AND DSB1.SupplierID = DSD.SupplierID " & vbCrLf & _
                              " AND DSB1.PartNo = DSD.PartNo " & vbCrLf & _
                              " /*LEFT JOIN (Select Suratjalanno,PONo,OrderNo,AffiliateID,SupplierID,PartNo,BoxNo = Max(boxNo)  " & vbCrLf & _
                              " 			from DOSupplier_DetailBox_Export Group By Suratjalanno,PONo,OrderNo,AffiliateID,SupplierID,PartNo) DSB2 ON DSB1.SuratJalanNo = DSD.Suratjalanno " & vbCrLf & _
                              " AND DSB2.PONo = DSD.PONo " & vbCrLf & _
                              " AND DSB2.OrderNo = DSD.OrderNo " & vbCrLf & _
                              " AND DSB2.AffiliateID = DSD.AffiliateID " & vbCrLf

            ls_sql = ls_sql + " AND DSB2.SupplierID = DSD.SupplierID " & vbCrLf & _
                              " AND DSB2.PartNo = DSD.PartNo*/ " & vbCrLf & _
                              " LEFT JOIN ReceiveForwarder_Master RM ON RM.SuratJalanno = DSM.Suratjalanno " & vbCrLf & _
                              " AND RM.PONo = DSM.PONo " & vbCrLf & _
                              " AND RM.OrderNo = DSM.OrderNo " & vbCrLf & _
                              " AND RM.AffiliateID = DSM.AffiliateID " & vbCrLf & _
                              " AND RM.SupplierID = DSM.SupplierID " & vbCrLf & _
                              " LEFT JOIN ReceiveForwarder_Detail RD ON RD.Suratjalanno = RM.Suratjalanno " & vbCrLf & _
                              " AND RD.PONo = RM.PONo " & vbCrLf & _
                              " AND RD.OrderNo = RM.OrderNo " & vbCrLf & _
                              " AND RD.AffiliateID = RM.AffiliateID " & vbCrLf

            ls_sql = ls_sql + " AND RD.SupplierID = RM.SupplierID " & vbCrLf & _
                              " AND RD.PartNo = DSD.PartNo " & vbCrLf & _
                              " LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              " AND RB.PONo = RD.PONo " & vbCrLf & _
                              " AND RB.OrderNo = RD.OrderNo " & vbCrLf & _
                              " AND RB.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              " AND RB.SupplierID  = RD.SupplierID " & vbCrLf & _
                              " AND RB.PartNo = RD.PartNo " & vbCrLf & _
                              " LEFT JOIN ShippingInstruction_Master SHM ON SHM.AffiliateID = RM.AffiliateID " & vbCrLf & _
                              " AND RM.ForwarderID = POM.ForwarderID " & vbCrLf & _
                              " LEFT JOIN ShippingInstruction_Detail SHD ON SHD.ShippingInstructionNo = SHM.ShippingInstructionNo " & vbCrLf

            ls_sql = ls_sql + " AND SHD.AffiliateID = SHM.AffiliateID " & vbCrLf & _
                              " AND SHD.OrderNo = RM.OrderNo " & vbCrLf & _
                              " AND SHD.SupplierID = RM.SupplierID " & vbCrLf & _
                              " AND SHD.PartNo = RD.PartNo " & vbCrLf & _
                              " AND SHD.SuratJalanno = RD.SuratJalanno " & vbCrLf & _
                              " AND SHD.ForwarderID = SHM.ForwarderID " & vbCrLf & _
                              " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " LEFT JOIN MS_Supplier MS ON MS.SupplierID = POM.SupplierID " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                              " LEFT JOIN MS_PartMapping MPM ON MPM.AffiliateID = POM.AffiliateID AND MPM.SupplierID = POM.SupplierID AND MPM.PartNo = POD.PartNo " & vbCrLf & _
                              " LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.unitCls " & vbCrLf

            ls_sql = ls_sql + " where isnull(RB.StatusDefect,0) <> 1 and POD.week1 <> 0 and isnull(PL1.BoxNo,'') <>'' --and POM.Period between '" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtPeriodTo.Value, "yyyy-MM-dd") & "'" & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select distinct  " & vbCrLf & _
                              "   --SuratJalanno = isnull(RM.SuratJalanno,''),   " & vbCrLf & _
                              "   Period = Right(Convert(char(11),Convert(datetime,POM.Period),106),8),   " & vbCrLf & _
                              "   OrderNo = POM.OrderNo1,    " & vbCrLf & _
                              "   AffiliateID = POM.AffiliateID,    " & vbCrLf & _
                              "   AffiliateName = AffiliateName,   " & vbCrLf & _
                              "   ETDVendor = Convert(char(11),convert(datetime,POM.ETDVendor1),106),   " & vbCrLf & _
                              "   ETDPort =  Convert(char(11),convert(datetime,POM.ETDPort1),106),   " & vbCrLf & _
                              "   ETAPort =  Convert(char(11),convert(datetime,POM.ETAPORT1),106),    ETAFactory =  Convert(char(11),convert(datetime,POM.ETAFACTORY1),106),   " & vbCrLf

            ls_sql = ls_sql + "   SupplierID = POM.SupplierID,   " & vbCrLf & _
                              "   SupplierName = MS.SupplierName,   " & vbCrLf & _
                              "   PartNo = POD.PartNo,    " & vbCrLf & _
                              "   PartName = MP.PartName,    " & vbCrLf & _
                              "   UOM = isnull(MU.Description,''),    " & vbCrLf & _
                              "   QtyBox = ISNULL(POD.POQtyBox,MPM.QtyBox),    " & vbCrLf & _
                              "   BoxNo = (Case when isnull(RD.suratjalanno,'') <> '' then isnull(Rtrim(RB.label1) + '-' + Rtrim(RB.label2),'')  " & vbCrLf & _
                              " 			when Isnull(DSM.SuratJalanNo,'') <> '' then isnull(Rtrim(DSB1.BoxNo) + '-' + Rtrim(DSB1.BoxNoMax),'')  " & vbCrLf & _
                              " 			Else isnull(Rtrim(PL1.BoxNo) + '-' + Rtrim(PL1.BoxNoMax),'') END), " & vbCrLf & _
                              "   --SJNo = Isnull(DSM.SuratJalanNo,''),   " & vbCrLf & _
                              "   POQty = Replace(Week1,'.00',''),   " & vbCrLf

            ls_sql = ls_sql + "   DOQty = Replace(isnull(Doqty,0),'.00',''),     " & vbCrLf & _
                              "   GoodRecQty = 0,   " & vbCrLf & _
                              "   DefectRecQty = Replace(isnull(Rb.Box,0)*ISNULL(POD.POQtyBox,MPM.QtyBox),'.00',''),   " & vbCrLf & _
                              "   Remaining = isnull(POD.Week1,0)-(isnull(RB.Box,0)*ISNULL(POD.POQtyBox,MPM.QtyBox)),  " & vbCrLf & _
                              "   EmergencyCls = Case when POM.EmergencyCls = 'M' then 'No' else 'Yes' END,  " & vbCrLf & _
                              "   BoxQty = RB.box " & vbCrLf & _
                              " From PO_Master_Export POM LEFT JOIN PO_Detail_Export POD " & vbCrLf & _
                              " ON POM.PONo = POD.PONo " & vbCrLf & _
                              " AND POM.OrderNo1 = POD.OrderNo1 " & vbCrLf & _
                              " AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_sql = ls_sql + " LEFT JOIN (Select PONo,OrderNo,AffiliateID,SupplierID,PartNo,BoxNo = Min(LabelNo),BoxNoMax = Max(LabelNo)  " & vbCrLf & _
                              " 			from PrintLabelExport group By PONo,OrderNo,AffiliateID,SupplierID,PartNo) PL1 ON PL1.PONo = POM.PONo " & vbCrLf & _
                              " AND PL1.OrderNo = POM.OrderNo1  " & vbCrLf & _
                              " AND PL1.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " AND PL1.SupplierID = POM.SupplierID " & vbCrLf & _
                              " AND PL1.PartNo = POD.PartNo " & vbCrLf & _
                              " /*LEFT JOIN (Select PONo,OrderNo,AffiliateID,SupplierID,PartNo,BoxNo = Max(LabelNo)  " & vbCrLf & _
                              " 			from PrintLabelExport group By PONo,OrderNo,AffiliateID,SupplierID,PartNo) PL2 ON PL2.PONo = POM.PONo " & vbCrLf & _
                              " AND PL2.OrderNo = POM.OrderNo1  " & vbCrLf & _
                              " AND PL2.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " AND PL2.SupplierID = POM.SupplierID " & vbCrLf

            ls_sql = ls_sql + " AND PL2.PartNo = POD.PartNo*/ " & vbCrLf & _
                              " LEFT JOIN DOSupplier_Master_Export DSM ON DSM.PONo = POM.PONo " & vbCrLf & _
                              " AND DSM.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " AND DSM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " AND DSM.SupplierID = POM.SupplierID " & vbCrLf & _
                              " LEFT JOIN DOSupplier_Detail_Export DSD ON DSD.Suratjalanno = DSM.Suratjalanno " & vbCrLf & _
                              " AND DSD.PONo = DSM.PONo " & vbCrLf & _
                              " AND DSD.OrderNo = DSM.OrderNo " & vbCrLf & _
                              " AND DSD.AffiliateID = DSM.AffiliateID " & vbCrLf & _
                              " AND DSD.SupplierID = DSM.SupplierID " & vbCrLf & _
                              " AND DSD.PartNo = POD.PartNo " & vbCrLf

            ls_sql = ls_sql + " LEFT JOIN (Select Suratjalanno,PONo,OrderNo,AffiliateID,SupplierID,PartNo,BoxNo = Min(boxNo),BoxNoMax = Max(boxNo)  " & vbCrLf & _
                              " 			from DOSupplier_DetailBox_Export Group By Suratjalanno,PONo,OrderNo,AffiliateID,SupplierID,PartNo) DSB1 ON DSB1.SuratJalanNo = DSD.Suratjalanno " & vbCrLf & _
                              " AND DSB1.PONo = DSD.PONo " & vbCrLf & _
                              " AND DSB1.OrderNo = DSD.OrderNo " & vbCrLf & _
                              " AND DSB1.AffiliateID = DSD.AffiliateID " & vbCrLf & _
                              " AND DSB1.SupplierID = DSD.SupplierID " & vbCrLf & _
                              " AND DSB1.PartNo = DSD.PartNo " & vbCrLf & _
                              " /*LEFT JOIN (Select Suratjalanno,PONo,OrderNo,AffiliateID,SupplierID,PartNo,BoxNo = Max(boxNo)  " & vbCrLf & _
                              " 			from DOSupplier_DetailBox_Export Group By Suratjalanno,PONo,OrderNo,AffiliateID,SupplierID,PartNo) DSB2 ON DSB1.SuratJalanNo = DSD.Suratjalanno " & vbCrLf & _
                              " AND DSB2.PONo = DSD.PONo " & vbCrLf & _
                              " AND DSB2.OrderNo = DSD.OrderNo " & vbCrLf

            ls_sql = ls_sql + " AND DSB2.AffiliateID = DSD.AffiliateID " & vbCrLf & _
                              " AND DSB2.SupplierID = DSD.SupplierID " & vbCrLf & _
                              " AND DSB2.PartNo = DSD.PartNo*/ " & vbCrLf & _
                              " LEFT JOIN ReceiveForwarder_Master RM ON RM.SuratJalanno = DSM.Suratjalanno " & vbCrLf & _
                              " AND RM.PONo = DSM.PONo " & vbCrLf & _
                              " AND RM.OrderNo = DSM.OrderNo " & vbCrLf & _
                              " AND RM.AffiliateID = DSM.AffiliateID " & vbCrLf & _
                              " AND RM.SupplierID = DSM.SupplierID " & vbCrLf & _
                              " LEFT JOIN ReceiveForwarder_Detail RD ON RD.Suratjalanno = RM.Suratjalanno " & vbCrLf & _
                              " AND RD.PONo = RM.PONo " & vbCrLf & _
                              " AND RD.OrderNo = RM.OrderNo " & vbCrLf

            ls_sql = ls_sql + " AND RD.AffiliateID = RM.AffiliateID " & vbCrLf & _
                              " AND RD.SupplierID = RM.SupplierID " & vbCrLf & _
                              " AND RD.PartNo = DSD.PartNo " & vbCrLf & _
                              " LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              " AND RB.PONo = RD.PONo " & vbCrLf & _
                              " AND RB.OrderNo = RD.OrderNo " & vbCrLf & _
                              " AND RB.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              " AND RB.SupplierID  = RD.SupplierID " & vbCrLf & _
                              " AND RB.PartNo = RD.PartNo " & vbCrLf & _
                              " LEFT JOIN ShippingInstruction_Master SHM ON SHM.AffiliateID = RM.AffiliateID " & vbCrLf & _
                              " AND RM.ForwarderID = POM.ForwarderID " & vbCrLf

            ls_sql = ls_sql + " LEFT JOIN ShippingInstruction_Detail SHD ON SHD.ShippingInstructionNo = SHM.ShippingInstructionNo " & vbCrLf & _
                              " AND SHD.AffiliateID = SHM.AffiliateID " & vbCrLf & _
                              " AND SHD.OrderNo = RM.OrderNo " & vbCrLf & _
                              " AND SHD.SupplierID = RM.SupplierID " & vbCrLf & _
                              " AND SHD.PartNo = RD.PartNo " & vbCrLf & _
                              " AND SHD.SuratJalanno = RD.SuratJalanno " & vbCrLf & _
                              " AND SHD.ForwarderID = SHM.ForwarderID " & vbCrLf & _
                              " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " LEFT JOIN MS_Supplier MS ON MS.SupplierID = POM.SupplierID " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                              " LEFT JOIN MS_PartMapping MPM ON MPM.AffiliateID = POM.AffiliateID AND MPM.SupplierID = POM.SupplierID AND MPM.PartNo = POD.PartNo " & vbCrLf

            ls_sql = ls_sql + " LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.unitCls " & vbCrLf & _
                              " where RB.StatusDefect = '1' " & vbCrLf & _
                              " )x where and PartNo <> '' and Period between '" & Format(dtPeriodFrom.Value, "MMM yyyy") & "' and '" & Format(dtPeriodTo.Value, "MMM yyyy") & "'" & vbCrLf & _
                              "  " & vbCrLf

            ls_sql = ls_sql + ls_Filter
            ls_sql = ls_sql + " order By PartNo "

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            dtExcel = ds.Tables(0)
            sqlConn.Close()

        End Using
    End Sub

#End Region

#Region "Excel"

    Private Sub GetExcel()
        Call up_GridLoad()
        FileName = "REMAINING REPORT.xlsx"
        FilePath = Server.MapPath("~\Template\" & FileName)
        If grid.VisibleRowCount - 1 > 0 Then
            Call epplusExportHeaderExcel(FilePath, "", dtExcel, "D:3", "")
        Else
            Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblInfo.Text
        End If
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

    Private Sub epplusExportHeaderExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData1 As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try

            Dim NewFileName As String = Server.MapPath("~\ProgressReportExport\Remaining Report.xlsx")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet

            ws = exl.Workbook.Worksheets("REMAINING")
            Dim irow As Long = 0
            Dim iRowTmp As Long = 0
            Dim icol As Long = 0

            With ws
                ws.Cells("K4").Value = Format(dtPeriodFrom.Value, "MMM yyyy") + "-" + Format(dtPeriodTo.Value, "MMM yyyy")
                ws.Cells("K4").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                If rdrRAll.Value = 1 Then ws.Cells("K5").Value = "ALL"
                If rdrRYes.Value = 1 Then ws.Cells("K5").Value = "YES"
                If rdrRNo.Value = 1 Then ws.Cells("K5").Value = "NO"
                ws.Cells("K5").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ws.Cells("K6").Value = IIf(Trim(cboPart.Text) = "==ALL==", "ALL", Trim(cboPart.Text) + "-" + Trim(txtPartName.Text))
                ws.Cells("K6").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ws.Cells("K7").Value = IIf(Trim(txtboxno.Text) = "", "-", Trim(txtboxno.Text))
                ws.Cells("K7").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left

                ws.Cells("AA4").Value = IIf(Trim(cbosupplier.Text) = "==ALL==", "ALL", Trim(cbosupplier.Text) + "-" + Trim(txtsupplier.Text))
                ws.Cells("AA4").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ws.Cells("AA5").Value = IIf(Trim(cboAffiliate.Text) = "==ALL==", "ALL", Trim(cboAffiliate.Text) + "'-" + Trim(txtAffiliate.Text))
                ws.Cells("AA5").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                If rdrEAll.Value = 1 Then ws.Cells("AA6").Value = "ALL"
                If rdrEyes.Value = 1 Then ws.Cells("AA6").Value = "YES"
                If rdrENo.Value = 1 Then ws.Cells("AA6").Value = "NO"
                ws.Cells("AA6").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ws.Cells("AA7").Value = IIf(Trim(txtpono.Text) = "", "-", Trim(txtpono.Text))
                ws.Cells("AA7").Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
            End With

            iRowTmp = 11
            For irow = 0 To pData1.Rows.Count - 1
                If pData1.Rows.Count > 0 Then
                    ws.Cells("B" & iRowTmp).Value = irow + 1
                    ws.Cells("B" & iRowTmp & ":" & "C" & iRowTmp).Merge = True
                    ws.Cells("D" & iRowTmp).Value = pData1.Rows(irow)("period")

                    ws.Cells("D" & iRowTmp & ":" & "F" & iRowTmp).Merge = True
                    ws.Cells("G" & iRowTmp).Value = pData1.Rows(irow)("AffiliateID")

                    ws.Cells("K" & iRowTmp).Value = pData1.Rows(irow)("AffiliateName")
                    ws.Cells("K" & iRowTmp & ":" & "S" & iRowTmp).Merge = True

                    ws.Cells("T" & iRowTmp).Value = pData1.Rows(irow)("Orderno")
                    ws.Cells("T" & iRowTmp & ":" & "Y" & iRowTmp).Merge = True

                    ws.Cells("Z" & iRowTmp).Value = pData1.Rows(irow)("EmergencyCls")
                    ws.Cells("Z" & iRowTmp & ":" & "AC" & iRowTmp).Merge = True
                    ws.Cells("Z" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center

                    ws.Cells("AD" & iRowTmp).Value = pData1.Rows(irow)("SupplierID")
                    ws.Cells("AD" & iRowTmp & ":" & "AG" & iRowTmp).Merge = True

                    ws.Cells("AH" & iRowTmp).Value = pData1.Rows(irow)("SupplierName")
                    ws.Cells("AH" & iRowTmp & ":" & "AO" & iRowTmp).Merge = True

                    ws.Cells("AP" & iRowTmp).Value = pData1.Rows(irow)("ETDVendor")
                    ws.Cells("AP" & iRowTmp & ":" & "AS" & iRowTmp).Merge = True
                    ws.Cells("AP" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center

                    ws.Cells("AT" & iRowTmp).Value = pData1.Rows(irow)("ETDPORT")
                    ws.Cells("AT" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                    ws.Cells("AT" & iRowTmp & ":" & "AW" & iRowTmp).Merge = True

                    ws.Cells("AX" & iRowTmp).Value = pData1.Rows(irow)("ETAPORT")
                    ws.Cells("AX" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                    ws.Cells("AX" & iRowTmp & ":" & "BA" & iRowTmp).Merge = True

                    ws.Cells("BB" & iRowTmp).Value = pData1.Rows(irow)("ETAFACTORY")
                    ws.Cells("BB" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                    ws.Cells("BB" & iRowTmp & ":" & "BE" & iRowTmp).Merge = True

                    ws.Cells("BF" & iRowTmp).Value = pData1.Rows(irow)("PartNo")
                    ws.Cells("BF" & iRowTmp & ":" & "BJ" & iRowTmp).Merge = True

                    ws.Cells("BK" & iRowTmp).Value = pData1.Rows(irow)("PartName")
                    ws.Cells("BK" & iRowTmp & ":" & "BR" & iRowTmp).Merge = True

                    ws.Cells("BS" & iRowTmp).Value = Trim(pData1.Rows(irow)("UOm"))
                    ws.Cells("BS" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                    ws.Cells("BS" & iRowTmp & ":" & "BT" & iRowTmp).Merge = True

                    ws.Cells("BU" & iRowTmp).Value = pData1.Rows(irow)("QtyBox")
                    ws.Cells("BU" & iRowTmp & ":" & "BW" & iRowTmp).Merge = True

                    ws.Cells("BX" & iRowTmp).Value = pData1.Rows(irow)("BoxNo")
                    ws.Cells("BX" & iRowTmp & ":" & "CD" & iRowTmp).Merge = True

                    ws.Cells("CE" & iRowTmp).Value = pData1.Rows(irow)("POQty")
                    ws.Cells("CE" & iRowTmp & ":" & "CH" & iRowTmp).Merge = True
                    ws.Cells("CE" & iRowTmp).Style.Numberformat.Format = "###,##0"
                    ws.Cells("CE" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    ws.Cells("CI" & iRowTmp).Value = pData1.Rows(irow)("DOQty")
                    ws.Cells("CI" & iRowTmp & ":" & "CM" & iRowTmp).Merge = True
                    ws.Cells("Ci" & iRowTmp).Style.Numberformat.Format = "###,##0"
                    ws.Cells("CI" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    ws.Cells("CN" & iRowTmp).Value = pData1.Rows(irow)("GoodRecQty")
                    ws.Cells("CN" & iRowTmp & ":" & "CR" & iRowTmp).Merge = True
                    ws.Cells("CN" & iRowTmp).Style.Numberformat.Format = "###,##0"
                    ws.Cells("CN" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    ws.Cells("CS" & iRowTmp).Value = pData1.Rows(irow)("DefectRecQty")
                    ws.Cells("CS" & iRowTmp & ":" & "CW" & iRowTmp).Merge = True
                    ws.Cells("CS" & iRowTmp).Style.Numberformat.Format = "###,##0"
                    ws.Cells("CS" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    ws.Cells("CX" & iRowTmp).Value = pData1.Rows(irow)("Remaining")
                    ws.Cells("CX" & iRowTmp & ":" & "DB" & iRowTmp).Merge = True
                    ws.Cells("CX" & iRowTmp).Style.Numberformat.Format = "###,##0"
                    ws.Cells("CX" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'ws.Cells("C18" & ":J" & iRowTmp).Style.Numberformat.Format = "#,###"
                End If
                iRowTmp = iRowTmp + 1
            Next


            Dim rgAll As ExcelRange = ws.Cells(11, 2, iRowTmp - 1, 106)
            EpPlusDrawAllBorders(rgAll)

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

#End Region

End Class