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

Public Class ExportOrderHistory
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = ""

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "O01"

    Dim dtHeader As DataTable
    Dim dtDetail As DataTable
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""

#End Region

#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_FillCombo()
                cboAffiliateCode.Text = clsGlobal.gs_All
                cboPart.Text = clsGlobal.gs_All
                txtAffiliateName.Text = clsGlobal.gs_All
                txtPartName.Text = clsGlobal.gs_All
                period.Text = Format(Now, "MMM yyyy")
                'Call up_GridLoadWhenEventChange()
                'Call up_Initialize()
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
#End Region

#Region "PROCEDURE"
    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        'Combo Affiliate
        With cboAffiliateCode
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
            End Using
        End With
    End Sub

    Private Sub GridLoadExcel()
        Dim ds As New DataSet
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""


        ls_Filter = ""

        If cboAffiliateCode.Text <> clsGlobal.gs_All And cboAffiliateCode.Text <> "" Then
            ls_Filter = ls_Filter + " AND AFF = '" & cboAffiliateCode.Text & "'" & vbCrLf
        End If

        If cboPart.Text <> clsGlobal.gs_All And cboPart.Text <> "" Then
            ls_Filter = ls_Filter + " AND PART = '" & cboPart.Text & "'" & vbCrLf
        End If

        If txtorderno.Text <> "" Then
            ls_Filter = ls_Filter + " AND ORD = '" & Trim(txtorderno.Text) & "'" & vbCrLf
        End If

        ls_SQL = "  SELECT * FROM ( " & vbCrLf & _
              "  SELECT  DISTINCT " & vbCrLf & _
              " 		ColNo = CONVERT(char, CONVERT(Numeric, ROW_NUMBER() OVER (ORDER BY POM.AffiliateID, POM.OrderNo, POD.PartNo))), " & vbCrLf & _
              "         idx = '0' , " & vbCrLf & _
              "         AffiliateID = POM.AffiliateID , " & vbCrLf & _
              "         AffiliateName = MA.AffiliateName , " & vbCrLf & _
              "         PartNo = POD.PartNo , " & vbCrLf & _
              "         PartName = MP.PartName , " & vbCrLf & _
              "         MOQ = CONVERT(char,(convert(numeric(9,0),MP.MOQ))) , " & vbCrLf & _
              "         UOM = MU.Description , " & vbCrLf & _
              "         cls = 'BY PASI' , "

        ls_SQL = ls_SQL + "         OrderNo = ISNULL(POM.OrderNo, '') , " & vbCrLf & _
                          "         FirmQty = CONVERT(char,(convert(numeric(9,0), CASE POM.Week " & vbCrLf & _
                          "                 WHEN '1' THEN POD.Week1 " & vbCrLf & _
                          "                 WHEN '2' THEN POD.Week2 " & vbCrLf & _
                          "                 WHEN '3' THEN POD.Week3 " & vbCrLf & _
                          "                 WHEN '4' THEN POD.Week4 " & vbCrLf & _
                          "                 WHEN '5' THEN POD.Week5 " & vbCrLf & _
                          "               END ))), " & vbCrLf & _
                          "         Curr = ISNULL(MC.DESCRIPTION, '') , " & vbCrLf & _
                          "         Price = ISNULL(ID.Price, 0) , " & vbCrLf & _
                          "         Amount = ISNULL(ID.Amount, 0), "

        ls_SQL = ls_SQL + "         AFF = POM.AffiliateID, " & vbCrLf & _
                          "         ORD = POM.OrderNo, " & vbCrLf & _
                          "         PART = POD.PartNo " & vbCrLf & _
                          "  FROM   ( SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo1 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor1 , " & vbCrLf & _
                          "                     ETAPort = ETAPort1 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory1 , " & vbCrLf & _
                          "                     week = 1 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL "

        ls_SQL = ls_SQL + "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo2 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor2 , " & vbCrLf & _
                          "                     ETAPort = ETAPort2 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory2 , " & vbCrLf & _
                          "                     week = 2 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo3 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor3 , "

        ls_SQL = ls_SQL + "                     ETAPort = ETAPort3 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory3 , " & vbCrLf & _
                          "                     week = 3 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo4 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor4 , " & vbCrLf & _
                          "                     ETAPort = ETAPort4 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory4 , " & vbCrLf & _
                          "                     week = 4 "

        ls_SQL = ls_SQL + "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo5 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor5 , " & vbCrLf & _
                          "                     ETAPort = ETAPort5 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory5 , " & vbCrLf & _
                          "                     week = 5 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "         ) POM " & vbCrLf & _
                          "         LEFT JOIN PO_Detail_Export POD ON POM.PONO = POD.PONO "

        ls_SQL = ls_SQL + "                                           AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                                           AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "         LEFT JOIN ( SELECT TOP 1 " & vbCrLf & _
                          "                             * , " & vbCrLf & _
                          "                             OrderNO = OrderNo1 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                          "                             ETAPort = ETAPort1 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory1 , " & vbCrLf & _
                          "                             week = 1 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     ORDER BY PORevNo "

        ls_SQL = ls_SQL + "                     UNION ALL " & vbCrLf & _
                          "                     SELECT TOP 1 " & vbCrLf & _
                          "                             * , " & vbCrLf & _
                          "                             OrderNO = OrderNo2 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                          "                             ETAPort = ETAPort2 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory2 , " & vbCrLf & _
                          "                             week = 2 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     ORDER BY PORevNo " & vbCrLf & _
                          "                     UNION ALL "

        ls_SQL = ls_SQL + "                     SELECT TOP 1 " & vbCrLf & _
                          "                             * , " & vbCrLf & _
                          "                             OrderNO = OrderNo3 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                          "                             ETAPort = ETAPort3 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory3 , " & vbCrLf & _
                          "                             week = 3 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     ORDER BY PORevNo " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT TOP 1 "

        ls_SQL = ls_SQL + "                             * , " & vbCrLf & _
                          "                             OrderNO = OrderNo4 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                          "                             ETAPort = ETAPort4 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory4 , " & vbCrLf & _
                          "                             week = 4 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     ORDER BY PORevNo " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT TOP 1 " & vbCrLf & _
                          "                             * , "

        ls_SQL = ls_SQL + "                             OrderNO = OrderNo5 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                          "                             ETAPort = ETAPort5 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory5 , " & vbCrLf & _
                          "                             week = 5 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     ORDER BY PORevNo " & vbCrLf & _
                          "                   ) PRM ON PRM.PONO = POD.PONO " & vbCrLf & _
                          "                            AND PRM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                            AND PRM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                            AND PRM.OrderNo = POM.OrderNo "

        ls_SQL = ls_SQL + "         LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO " & vbCrLf & _
                          "                                              AND PRD.AffiliateID = PRM.AffiliateID " & vbCrLf & _
                          "                                              AND PRD.SupplierID = PRM.SupplierID " & vbCrLf & _
                          "                                              AND PRD.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN ReceiveForwarder_Detail RD ON POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                          "                                                 AND POD.SupplierID = RD.SupplierID " & vbCrLf & _
                          "                                                 AND POD.POno = RD.POno " & vbCrLf & _
                          "                                                 AND POM.OrderNo = RD.OrderNo " & vbCrLf & _
                          "                                                 AND POD.PartNo = RD.PartNo " & vbCrLf & _
                          "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                          "                                                       AND IM.AffiliateID = POD.AffiliateID "

        ls_SQL = ls_SQL + "                                                       AND IM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                                                       AND IM.POno = POD.POno " & vbCrLf & _
                          "                                                       AND IM.OrderNo = POM.OrderNo " & vbCrLf & _
                          "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                          "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                          "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                          "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                          "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                          "                                                       AND ID.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo "

        ls_SQL = ls_SQL + "         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                          "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                          "  WHERE  ISNULL(POM.OrderNO, '') <> ''  " & vbCrLf & _
                          " -------------------------------------------------------------------------------------------------------- " & vbCrLf & _
                          "  UNION ALL --SUPPLIER " & vbCrLf & _
                          "  SELECT   DISTINCT " & vbCrLf & _
                          " 		ColNo = '', " & vbCrLf & _
                          "         idx = '1' , " & vbCrLf & _
                          "         AffiliateID = '' , " & vbCrLf & _
                          "         AffiliateName = '' , " & vbCrLf & _
                          "         PartNo = '' , "

        ls_SQL = ls_SQL + "         PartName ='' , " & vbCrLf & _
                          "         MOQ = '' , " & vbCrLf & _
                          "         UOM = '' , " & vbCrLf & _
                          "         cls = 'BY SUPPLIER' , " & vbCrLf & _
                          "         OrderNo = '' , " & vbCrLf & _
                          "         FirmQty = CONVERT(char,(convert(numeric(9,0), isnull(CASE POM.Week " & vbCrLf & _
                          "                 WHEN '1' THEN PDU.Week1 " & vbCrLf & _
                          "                 WHEN '2' THEN PDU.Week2 " & vbCrLf & _
                          "                 WHEN '3' THEN PDU.Week3 " & vbCrLf & _
                          "                 WHEN '4' THEN PDU.Week4 " & vbCrLf & _
                          "                 WHEN '5' THEN PDU.Week5 "

        ls_SQL = ls_SQL + "               END,0) ))), " & vbCrLf & _
                          "         Curr = ISNULL(MC.DESCRIPTION, '') , " & vbCrLf & _
                          "         Price = ISNULL(ID.Price, 0) , " & vbCrLf & _
                          "         Amount = ISNULL(ID.Amount, 0), " & vbCrLf & _
                          "         AFF = POM.AffiliateID, " & vbCrLf & _
                          "         ORD = POM.OrderNo, " & vbCrLf & _
                          "         PART = POD.PartNo " & vbCrLf & _
                          "  FROM   ( SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo1 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor1 , " & vbCrLf & _
                          "                     ETAPort = ETAPort1 , "

        ls_SQL = ls_SQL + "                     ETAFactory = ETAFactory1 , " & vbCrLf & _
                          "                     week = 1 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo2 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor2 , " & vbCrLf & _
                          "                     ETAPort = ETAPort2 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory2 , " & vbCrLf & _
                          "                     week = 2 " & vbCrLf & _
                          "           FROM      Po_Master_Export "

        ls_SQL = ls_SQL + "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo3 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor3 , " & vbCrLf & _
                          "                     ETAPort = ETAPort3 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory3 , " & vbCrLf & _
                          "                     week = 3 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo4 , "

        ls_SQL = ls_SQL + "                     ETDVendor = ETDVendor4 , " & vbCrLf & _
                          "                     ETAPort = ETAPort4 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory4 , " & vbCrLf & _
                          "                     week = 4 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo5 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor5 , " & vbCrLf & _
                          "                     ETAPort = ETAPort5 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory5 , "

        ls_SQL = ls_SQL + "                     week = 5 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "         ) POM " & vbCrLf & _
                          "         LEFT JOIN PO_Detail_Export POD ON POM.PONO = POD.PONO " & vbCrLf & _
                          "                                           AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                                           AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "         LEFT JOIN ( SELECT  * , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                          "                             week = 1 " & vbCrLf & _
                          "                     FROM    PO_masterUpload_export " & vbCrLf & _
                          "                     UNION ALL "

        ls_SQL = ls_SQL + "                     SELECT  * , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                          "                             week = 2 " & vbCrLf & _
                          "                     FROM    PO_masterUpload_export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                          "                             week = 3 " & vbCrLf & _
                          "                     FROM    PO_masterUpload_export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , "

        ls_SQL = ls_SQL + "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                          "                             week = 4 " & vbCrLf & _
                          "                     FROM    PO_masterUpload_export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                          "                             week = 5 " & vbCrLf & _
                          "                     FROM    PO_masterUpload_export " & vbCrLf & _
                          "                   ) PMU ON PMU.PONO = POD.PONO " & vbCrLf & _
                          "                            AND PMU.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                            AND PMU.SupplierID = POD.SupplierID "

        ls_SQL = ls_SQL + "         LEFT JOIN PO_DetailUpload_export PDU ON PDU.PONO = PMU.PONO " & vbCrLf & _
                          "                                                 AND PDU.AffiliateID = PMU.AffiliateID " & vbCrLf & _
                          "                                                 AND PDU.SupplierID = PMU.SupplierID " & vbCrLf & _
                          "                                                 AND PDU.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN ReceiveForwarder_Detail RD ON POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                          "                                                 AND POD.SupplierID = RD.SupplierID " & vbCrLf & _
                          "                                                 AND POD.POno = RD.POno " & vbCrLf & _
                          "                                                 AND POM.OrderNo = RD.OrderNo " & vbCrLf & _
                          "                                                 AND POD.PartNo = RD.PartNo " & vbCrLf & _
                          "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                          "                                                       AND IM.AffiliateID = POD.AffiliateID "

        ls_SQL = ls_SQL + "                                                       AND IM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                                                       AND IM.POno = POD.POno " & vbCrLf & _
                          "                                                       AND IM.OrderNo = POM.OrderNo " & vbCrLf & _
                          "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                          "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                          "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                          "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                          "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                          "                                                       AND ID.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo "

        ls_SQL = ls_SQL + "         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                          "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                          "  WHERE  ISNULL(POM.OrderNO, '') <> ''  " & vbCrLf & _
                          " -------------------------------------------------------------------------------------------------------- " & vbCrLf & _
                          "  UNION ALL -- REVISION " & vbCrLf & _
                          "  SELECT  DISTINCT " & vbCrLf & _
                          " 		ColNo = '', " & vbCrLf & _
                          "         idx = '2' , " & vbCrLf & _
                          "         AffiliateID = '' , " & vbCrLf & _
                          "         AffiliateName = '' , " & vbCrLf & _
                          "         PartNo = '' , "

        ls_SQL = ls_SQL + "         PartName = '' , " & vbCrLf & _
                          "         MOQ = '' , " & vbCrLf & _
                          "         UOM = '' , " & vbCrLf & _
                          "         cls = 'REVISION' , " & vbCrLf & _
                          "         OrderNo = '' , " & vbCrLf & _
                          "         FirmQty = CONVERT(char,(convert(numeric(9,0), isnull(CASE POM.Week " & vbCrLf & _
                          "                 WHEN '1' THEN PRD.Week1 " & vbCrLf & _
                          "                 WHEN '2' THEN PRD.Week2 " & vbCrLf & _
                          "                 WHEN '3' THEN PRD.Week3 " & vbCrLf & _
                          "                 WHEN '4' THEN PRD.Week4 " & vbCrLf & _
                          "                 WHEN '5' THEN PRD.Week5 "

        ls_SQL = ls_SQL + "               END,0) ))), " & vbCrLf & _
                          "         Curr = ISNULL(MC.DESCRIPTION, '') , " & vbCrLf & _
                          "         Price = ISNULL(ID.Price, 0) , " & vbCrLf & _
                          "         Amount = ISNULL(ID.Amount, 0), " & vbCrLf & _
                          "         AFF = POM.AffiliateID, " & vbCrLf & _
                          "         ORD = POM.OrderNo, " & vbCrLf & _
                          "         PART = POD.PartNo " & vbCrLf & _
                          "  FROM   ( SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo1 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor1 , " & vbCrLf & _
                          "                     ETAPort = ETAPort1 , "

        ls_SQL = ls_SQL + "                     ETAFactory = ETAFactory1 , " & vbCrLf & _
                          "                     week = 1 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo2 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor2 , " & vbCrLf & _
                          "                     ETAPort = ETAPort2 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory2 , " & vbCrLf & _
                          "                     week = 2 " & vbCrLf & _
                          "           FROM      Po_Master_Export "

        ls_SQL = ls_SQL + "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo3 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor3 , " & vbCrLf & _
                          "                     ETAPort = ETAPort3 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory3 , " & vbCrLf & _
                          "                     week = 3 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo4 , "

        ls_SQL = ls_SQL + "                     ETDVendor = ETDVendor4 , " & vbCrLf & _
                          "                     ETAPort = ETAPort4 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory4 , " & vbCrLf & _
                          "                     week = 4 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo5 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor5 , " & vbCrLf & _
                          "                     ETAPort = ETAPort5 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory5 , "

        ls_SQL = ls_SQL + "                     week = 5 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "         ) POM " & vbCrLf & _
                          "         LEFT JOIN PO_Detail_Export POD ON POM.PONO = POD.PONO " & vbCrLf & _
                          "                                           AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                                           AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "         LEFT JOIN ( SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo1 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                          "                             ETAPort = ETAPort1 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory1 , "

        ls_SQL = ls_SQL + "                             week = 1 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo2 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                          "                             ETAPort = ETAPort2 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory2 , " & vbCrLf & _
                          "                             week = 2 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     UNION ALL "

        ls_SQL = ls_SQL + "                     SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo3 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                          "                             ETAPort = ETAPort3 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory3 , " & vbCrLf & _
                          "                             week = 3 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo4 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor4 , "

        ls_SQL = ls_SQL + "                             ETAPort = ETAPort4 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory4 , " & vbCrLf & _
                          "                             week = 4 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo5 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                          "                             ETAPort = ETAPort5 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory5 , " & vbCrLf & _
                          "                             week = 5 "

        ls_SQL = ls_SQL + "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                   ) PRM ON PRM.PONO = POD.PONO " & vbCrLf & _
                          "                            AND PRM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                            AND PRM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                            AND PRM.OrderNo = POM.OrderNo " & vbCrLf & _
                          "         LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO " & vbCrLf & _
                          "                                              AND PRD.AffiliateID = PRM.AffiliateID " & vbCrLf & _
                          "                                              AND PRD.SupplierID = PRM.SupplierID " & vbCrLf & _
                          "                                              AND PRD.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN ReceiveForwarder_Detail RD ON POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                          "                                                 AND POD.SupplierID = RD.SupplierID "

        ls_SQL = ls_SQL + "                                                 AND POD.POno = RD.POno " & vbCrLf & _
                          "                                                 AND POM.OrderNo = RD.OrderNo " & vbCrLf & _
                          "                                                 AND POD.PartNo = RD.PartNo " & vbCrLf & _
                          "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                          "                                                       AND IM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                                                       AND IM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                                                       AND IM.POno = POD.POno " & vbCrLf & _
                          "                                                       AND IM.OrderNo = POM.OrderNo " & vbCrLf & _
                          "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                          "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                          "                                                       AND ID.SupplierID = IM.SupplierID "

        ls_SQL = ls_SQL + "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                          "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                          "                                                       AND ID.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                          "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                          "  WHERE  ISNULL(PRM.OrderNO, '') <> ''  " & vbCrLf & _
                          "  ------------------------------------------------------------------------------------------------------ " & vbCrLf & _
                          "  UNION ALL -- DIFF " & vbCrLf & _
                          "  SELECT  DISTINCT "

        ls_SQL = ls_SQL + " 		ColNo = '', " & vbCrLf & _
                          "         idx = '3' , " & vbCrLf & _
                          "         AffiliateID = '', " & vbCrLf & _
                          "         AffiliateName = '', " & vbCrLf & _
                          "         PartNo ='', " & vbCrLf & _
                          "         PartName = '', " & vbCrLf & _
                          "         MOQ = '', " & vbCrLf & _
                          "         UOM = '', " & vbCrLf & _
                          "         cls = 'DIFFERENCE' , " & vbCrLf & _
                          "         OrderNo ='', " & vbCrLf & _
                          "         FirmQty = 0 , "

        ls_SQL = ls_SQL + "         Curr = '',  " & vbCrLf & _
                          "         Price = ISNULL(ID.Price, 0) , " & vbCrLf & _
                          "         Amount = ISNULL(ID.Amount, 0), " & vbCrLf & _
                          "         AFF = POM.AffiliateID, " & vbCrLf & _
                          "         ORD = POM.OrderNo, " & vbCrLf & _
                          "         PART = POD.PartNo " & vbCrLf & _
                          "  FROM   ( SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo1 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor1 , " & vbCrLf & _
                          "                     ETAPort = ETAPort1 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory1 , "

        ls_SQL = ls_SQL + "                     week = 1 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo2 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor2 , " & vbCrLf & _
                          "                     ETAPort = ETAPort2 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory2 , " & vbCrLf & _
                          "                     week = 2 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL "

        ls_SQL = ls_SQL + "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo3 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor3 , " & vbCrLf & _
                          "                     ETAPort = ETAPort3 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory3 , " & vbCrLf & _
                          "                     week = 3 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo4 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor4 , "

        ls_SQL = ls_SQL + "                     ETAPort = ETAPort4 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory4 , " & vbCrLf & _
                          "                     week = 4 " & vbCrLf & _
                          "           FROM      Po_Master_Export " & vbCrLf & _
                          "           UNION ALL " & vbCrLf & _
                          "           SELECT    * , " & vbCrLf & _
                          "                     OrderNO = OrderNo5 , " & vbCrLf & _
                          "                     ETDVendor = ETDVendor5 , " & vbCrLf & _
                          "                     ETAPort = ETAPort5 , " & vbCrLf & _
                          "                     ETAFactory = ETAFactory5 , " & vbCrLf & _
                          "                     week = 5 "

        ls_SQL = ls_SQL + "           FROM      Po_Master_Export " & vbCrLf & _
                          "         ) POM " & vbCrLf & _
                          "         LEFT JOIN PO_Detail_Export POD ON POM.PONO = POD.PONO " & vbCrLf & _
                          "                                           AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                                           AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "         LEFT JOIN ( SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo1 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                          "                             ETAPort = ETAPort1 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory1 , " & vbCrLf & _
                          "                             week = 1 "

        ls_SQL = ls_SQL + "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo2 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                          "                             ETAPort = ETAPort2 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory2 , " & vbCrLf & _
                          "                             week = 2 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , "

        ls_SQL = ls_SQL + "                             OrderNO = OrderNo3 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                          "                             ETAPort = ETAPort3 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory3 , " & vbCrLf & _
                          "                             week = 3 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo4 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                          "                             ETAPort = ETAPort4 , "

        ls_SQL = ls_SQL + "                             ETAFactory = ETAFactory4 , " & vbCrLf & _
                          "                             week = 4 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export " & vbCrLf & _
                          "                     UNION ALL " & vbCrLf & _
                          "                     SELECT  * , " & vbCrLf & _
                          "                             OrderNO = OrderNo5 , " & vbCrLf & _
                          "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                          "                             ETAPort = ETAPort5 , " & vbCrLf & _
                          "                             ETAFactory = ETAFactory5 , " & vbCrLf & _
                          "                             week = 5 " & vbCrLf & _
                          "                     FROM    PoRev_Master_Export "

        ls_SQL = ls_SQL + "                   ) PRM ON PRM.PONO = POD.PONO " & vbCrLf & _
                          "                            AND PRM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                            AND PRM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                            AND PRM.OrderNo = POM.OrderNo " & vbCrLf & _
                          "         LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO " & vbCrLf & _
                          "                                              AND PRD.AffiliateID = PRM.AffiliateID " & vbCrLf & _
                          "                                              AND PRD.SupplierID = PRM.SupplierID " & vbCrLf & _
                          "                                              AND PRD.PartNo = POD.PartNo " & vbCrLf & _
                          " 		LEFT JOIN PO_DetailUpload_export PDU ON PDU.PONO = POM.PONO " & vbCrLf & _
                          "                                                 AND PDU.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "                                                 AND PDU.SupplierID = POM.SupplierID "

        ls_SQL = ls_SQL + "                                                 AND PDU.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN ReceiveForwarder_Detail RD ON POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                          "                                                 AND POD.SupplierID = RD.SupplierID " & vbCrLf & _
                          "                                                 AND POD.POno = RD.POno " & vbCrLf & _
                          "                                                 AND POM.OrderNo = RD.OrderNo " & vbCrLf & _
                          "                                                 AND POD.PartNo = RD.PartNo " & vbCrLf & _
                          "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                          "                                                       AND IM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          "                                                       AND IM.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                                                       AND IM.POno = POD.POno " & vbCrLf & _
                          "                                                       AND IM.OrderNo = POM.OrderNo "

        ls_SQL = ls_SQL + "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                          "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                          "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                          "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                          "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                          "                                                       AND ID.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                          "         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                          "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                          "  WHERE  ISNULL(POM.OrderNO, '') <> '' " & vbCrLf & _
                          "   )x " & vbCrLf & _
                          " -------------------------------------------------------------------------------------------------------  " & vbCrLf & _
                          " WHERE ORD <> '' " & vbCrLf

        ls_SQL = ls_SQL + ls_Filter

        ls_SQL = ls_SQL + " ORDER BY AFF ,ORD, PART ,idx  "


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using
        dtHeader = ds.Tables(0)
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_Filter = ""

            If cboAffiliateCode.Text <> clsGlobal.gs_All And cboAffiliateCode.Text <> "" Then
                ls_Filter = ls_Filter + " AND AFF = '" & cboAffiliateCode.Text & "'" & vbCrLf
            End If

            If cboPart.Text <> clsGlobal.gs_All And cboPart.Text <> "" Then
                ls_Filter = ls_Filter + " AND PART = '" & cboPart.Text & "'" & vbCrLf
            End If

            If txtorderno.Text <> "" Then
                ls_Filter = ls_Filter + " AND ORD like '%" & Trim(txtorderno.Text) & "%'" & vbCrLf
            End If


            ls_SQL = "  SELECT * FROM ( " & vbCrLf & _
                  "  SELECT  DISTINCT " & vbCrLf & _
                  " 		ColNo = CONVERT(char, CONVERT(Numeric, ROW_NUMBER() OVER (ORDER BY POM.AffiliateID, POM.OrderNo, POD.PartNo))), " & vbCrLf & _
                  "         idx = '0' , " & vbCrLf & _
                  "         AffiliateID = POM.AffiliateID , " & vbCrLf & _
                  "         AffiliateName = MA.AffiliateName , " & vbCrLf & _
                  "         PartNo = POD.PartNo , " & vbCrLf & _
                  "         PartName = MP.PartName , " & vbCrLf & _
                  "         MOQ = CONVERT(char,(convert(numeric(9,0),MP.MOQ))) , " & vbCrLf & _
                  "         UOM = MU.Description , " & vbCrLf & _
                  "         cls = 'BY PASI' , "

            ls_SQL = ls_SQL + "         OrderNo = ISNULL(POM.OrderNo, '') , " & vbCrLf & _
                              "         FirmQty = CONVERT(char,(convert(numeric(9,0), CASE POM.Week " & vbCrLf & _
                              "                 WHEN '1' THEN POD.Week1 " & vbCrLf & _
                              "                 WHEN '2' THEN POD.Week2 " & vbCrLf & _
                              "                 WHEN '3' THEN POD.Week3 " & vbCrLf & _
                              "                 WHEN '4' THEN POD.Week4 " & vbCrLf & _
                              "                 WHEN '5' THEN POD.Week5 " & vbCrLf & _
                              "               END ))), " & vbCrLf & _
                              "         Curr = ISNULL(MC.DESCRIPTION, '') , " & vbCrLf & _
                              "         Price = ISNULL(ID.Price, 0) , " & vbCrLf & _
                              "         Amount = ISNULL(ID.Amount, 0), "

            ls_SQL = ls_SQL + "         AFF = POM.AffiliateID, " & vbCrLf & _
                              "         ORD = POM.OrderNo, " & vbCrLf & _
                              "         PART = POD.PartNo " & vbCrLf & _
                              "         ,PERIOD = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) " & vbCrLf & _
                              "  FROM   ( SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo1 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                     ETAPort = ETAPort1 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory1 , " & vbCrLf & _
                              "                     week = 1 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL "

            ls_SQL = ls_SQL + "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo2 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                     ETAPort = ETAPort2 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory2 , " & vbCrLf & _
                              "                     week = 2 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo3 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor3 , "

            ls_SQL = ls_SQL + "                     ETAPort = ETAPort3 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory3 , " & vbCrLf & _
                              "                     week = 3 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo4 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                     ETAPort = ETAPort4 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory4 , " & vbCrLf & _
                              "                     week = 4 "

            ls_SQL = ls_SQL + "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo5 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                     ETAPort = ETAPort5 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory5 , " & vbCrLf & _
                              "                     week = 5 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "         ) POM " & vbCrLf & _
                              "         LEFT JOIN PO_Detail_Export POD ON POM.PONO = POD.PONO "

            ls_SQL = ls_SQL + "                                           AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                           AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN ( SELECT TOP 1 " & vbCrLf & _
                              "                             * , " & vbCrLf & _
                              "                             OrderNO = OrderNo1 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                             ETAPort = ETAPort1 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory1 , " & vbCrLf & _
                              "                             week = 1 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo "

            ls_SQL = ls_SQL + "                     UNION ALL " & vbCrLf & _
                              "                     SELECT TOP 1 " & vbCrLf & _
                              "                             * , " & vbCrLf & _
                              "                             OrderNO = OrderNo2 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                             ETAPort = ETAPort2 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory2 , " & vbCrLf & _
                              "                             week = 2 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo " & vbCrLf & _
                              "                     UNION ALL "

            ls_SQL = ls_SQL + "                     SELECT TOP 1 " & vbCrLf & _
                              "                             * , " & vbCrLf & _
                              "                             OrderNO = OrderNo3 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                             ETAPort = ETAPort3 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory3 , " & vbCrLf & _
                              "                             week = 3 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT TOP 1 "

            ls_SQL = ls_SQL + "                             * , " & vbCrLf & _
                              "                             OrderNO = OrderNo4 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                             ETAPort = ETAPort4 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory4 , " & vbCrLf & _
                              "                             week = 4 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT TOP 1 " & vbCrLf & _
                              "                             * , "

            ls_SQL = ls_SQL + "                             OrderNO = OrderNo5 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                             ETAPort = ETAPort5 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory5 , " & vbCrLf & _
                              "                             week = 5 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo " & vbCrLf & _
                              "                   ) PRM ON PRM.PONO = POD.PONO " & vbCrLf & _
                              "                            AND PRM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                            AND PRM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                            AND PRM.OrderNo = POM.OrderNo "

            ls_SQL = ls_SQL + "         LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO " & vbCrLf & _
                              "                                              AND PRD.AffiliateID = PRM.AffiliateID " & vbCrLf & _
                              "                                              AND PRD.SupplierID = PRM.SupplierID " & vbCrLf & _
                              "                                              AND PRD.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN ReceiveForwarder_Detail RD ON POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                 AND POD.SupplierID = RD.SupplierID " & vbCrLf & _
                              "                                                 AND POD.POno = RD.POno " & vbCrLf & _
                              "                                                 AND POM.OrderNo = RD.OrderNo " & vbCrLf & _
                              "                                                 AND POD.PartNo = RD.PartNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              "                                                       AND IM.AffiliateID = POD.AffiliateID "

            ls_SQL = ls_SQL + "                                                       AND IM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                       AND IM.POno = POD.POno " & vbCrLf & _
                              "                                                       AND IM.OrderNo = POM.OrderNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                              "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                              "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                              "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                              "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                              "                                                       AND ID.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo "

            ls_SQL = ls_SQL + "         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                              "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                              "  WHERE  ISNULL(POM.OrderNO, '') <> ''  " & vbCrLf & _
                              "  AND RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) = '" & period.Text & "'" & vbCrLf

            If cboAffiliateCode.Text <> clsGlobal.gs_All And cboAffiliateCode.Text <> "" Then
                ls_SQL = ls_SQL + " AND POM.AffiliateID = '" & cboAffiliateCode.Text & "'" & vbCrLf
            End If

            If cboPart.Text <> clsGlobal.gs_All And cboPart.Text <> "" Then
                ls_SQL = ls_SQL + " AND POD.PartNo = '" & cboPart.Text & "'" & vbCrLf
            End If

            If txtorderno.Text <> "" Then
                ls_SQL = ls_SQL + " AND POM.OrderNo like '%" & Trim(txtorderno.Text) & "%'" & vbCrLf
            End If


            ls_SQL = ls_SQL + " -------------------------------------------------------------------------------------------------------- " & vbCrLf & _
                              "  UNION ALL --SUPPLIER " & vbCrLf & _
                              "  SELECT   DISTINCT " & vbCrLf & _
                              " 		ColNo = '', " & vbCrLf & _
                              "         idx = '1' , " & vbCrLf & _
                              "         AffiliateID = '' , " & vbCrLf & _
                              "         AffiliateName = '' , " & vbCrLf & _
                              "         PartNo = '' , "

            ls_SQL = ls_SQL + "         PartName ='' , " & vbCrLf & _
                              "         MOQ = '' , " & vbCrLf & _
                              "         UOM = '' , " & vbCrLf & _
                              "         cls = 'BY SUPPLIER' , " & vbCrLf & _
                              "         OrderNo = '' , " & vbCrLf & _
                              "         FirmQty = CONVERT(char,(convert(numeric(9,0), isnull(CASE POM.Week " & vbCrLf & _
                              "                 WHEN '1' THEN PDU.Week1 " & vbCrLf & _
                              "                 WHEN '2' THEN PDU.Week2 " & vbCrLf & _
                              "                 WHEN '3' THEN PDU.Week3 " & vbCrLf & _
                              "                 WHEN '4' THEN PDU.Week4 " & vbCrLf & _
                              "                 WHEN '5' THEN PDU.Week5 "

            ls_SQL = ls_SQL + "               END,0) ))), " & vbCrLf & _
                              "         Curr = ISNULL(MC.DESCRIPTION, '') , " & vbCrLf & _
                              "         Price = ISNULL(ID.Price, 0) , " & vbCrLf & _
                              "         Amount = ISNULL(ID.Amount, 0), " & vbCrLf & _
                              "         AFF = POM.AffiliateID, " & vbCrLf & _
                              "         ORD = POM.OrderNo, " & vbCrLf & _
                              "         PART = POD.PartNo " & vbCrLf & _
                              "         ,PERIOD = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) " & vbCrLf & _
                              "  FROM   ( SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo1 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                     ETAPort = ETAPort1 , "

            ls_SQL = ls_SQL + "                     ETAFactory = ETAFactory1 , " & vbCrLf & _
                              "                     week = 1 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo2 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                     ETAPort = ETAPort2 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory2 , " & vbCrLf & _
                              "                     week = 2 " & vbCrLf & _
                              "           FROM      Po_Master_Export "

            ls_SQL = ls_SQL + "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo3 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                     ETAPort = ETAPort3 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory3 , " & vbCrLf & _
                              "                     week = 3 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo4 , "

            ls_SQL = ls_SQL + "                     ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                     ETAPort = ETAPort4 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory4 , " & vbCrLf & _
                              "                     week = 4 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo5 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                     ETAPort = ETAPort5 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory5 , "

            ls_SQL = ls_SQL + "                     week = 5 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "         ) POM " & vbCrLf & _
                              "         LEFT JOIN PO_Detail_Export POD ON POM.PONO = POD.PONO " & vbCrLf & _
                              "                                           AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                           AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN ( SELECT  * , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                             week = 1 " & vbCrLf & _
                              "                     FROM    PO_masterUpload_export " & vbCrLf & _
                              "                     UNION ALL "

            ls_SQL = ls_SQL + "                     SELECT  * , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                             week = 2 " & vbCrLf & _
                              "                     FROM    PO_masterUpload_export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                             week = 3 " & vbCrLf & _
                              "                     FROM    PO_masterUpload_export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , "

            ls_SQL = ls_SQL + "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                             week = 4 " & vbCrLf & _
                              "                     FROM    PO_masterUpload_export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                             week = 5 " & vbCrLf & _
                              "                     FROM    PO_masterUpload_export " & vbCrLf & _
                              "                   ) PMU ON PMU.PONO = POD.PONO " & vbCrLf & _
                              "                            AND PMU.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                            AND PMU.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "         LEFT JOIN PO_DetailUpload_export PDU ON PDU.PONO = PMU.PONO " & vbCrLf & _
                              "                                                 AND PDU.AffiliateID = PMU.AffiliateID " & vbCrLf & _
                              "                                                 AND PDU.SupplierID = PMU.SupplierID " & vbCrLf & _
                              "                                                 AND PDU.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN ReceiveForwarder_Detail RD ON POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                 AND POD.SupplierID = RD.SupplierID " & vbCrLf & _
                              "                                                 AND POD.POno = RD.POno " & vbCrLf & _
                              "                                                 AND POM.OrderNo = RD.OrderNo " & vbCrLf & _
                              "                                                 AND POD.PartNo = RD.PartNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              "                                                       AND IM.AffiliateID = POD.AffiliateID "

            ls_SQL = ls_SQL + "                                                       AND IM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                       AND IM.POno = POD.POno " & vbCrLf & _
                              "                                                       AND IM.OrderNo = POM.OrderNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                              "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                              "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                              "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                              "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                              "                                                       AND ID.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo "

            ls_SQL = ls_SQL + "         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                              "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                              "  WHERE  ISNULL(POM.OrderNO, '') <> ''  " & vbCrLf & _
                              "  AND RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) = '" & period.Text & "'" & vbCrLf & _
                              " -------------------------------------------------------------------------------------------------------- " & vbCrLf & _
                              "  UNION ALL -- REVISION " & vbCrLf & _
                              "  SELECT  DISTINCT " & vbCrLf & _
                              " 		ColNo = '', " & vbCrLf & _
                              "         idx = '2' , " & vbCrLf & _
                              "         AffiliateID = '' , " & vbCrLf & _
                              "         AffiliateName = '' , " & vbCrLf & _
                              "         PartNo = '' , "

            ls_SQL = ls_SQL + "         PartName = '' , " & vbCrLf & _
                              "         MOQ = '' , " & vbCrLf & _
                              "         UOM = '' , " & vbCrLf & _
                              "         cls = 'REVISION' , " & vbCrLf & _
                              "         OrderNo = '' , " & vbCrLf & _
                              "         FirmQty = CONVERT(char,(convert(numeric(9,0), isnull(CASE POM.Week " & vbCrLf & _
                              "                 WHEN '1' THEN PRD.Week1 " & vbCrLf & _
                              "                 WHEN '2' THEN PRD.Week2 " & vbCrLf & _
                              "                 WHEN '3' THEN PRD.Week3 " & vbCrLf & _
                              "                 WHEN '4' THEN PRD.Week4 " & vbCrLf & _
                              "                 WHEN '5' THEN PRD.Week5 "

            ls_SQL = ls_SQL + "               END,0) ))), " & vbCrLf & _
                              "         Curr = ISNULL(MC.DESCRIPTION, '') , " & vbCrLf & _
                              "         Price = ISNULL(ID.Price, 0) , " & vbCrLf & _
                              "         Amount = ISNULL(ID.Amount, 0), " & vbCrLf & _
                              "         AFF = POM.AffiliateID, " & vbCrLf & _
                              "         ORD = POM.OrderNo, " & vbCrLf & _
                              "         PART = POD.PartNo " & vbCrLf & _
                              "         ,PERIOD = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) " & vbCrLf & _
                              "  FROM   ( SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo1 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                     ETAPort = ETAPort1 , "

            ls_SQL = ls_SQL + "                     ETAFactory = ETAFactory1 , " & vbCrLf & _
                              "                     week = 1 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo2 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                     ETAPort = ETAPort2 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory2 , " & vbCrLf & _
                              "                     week = 2 " & vbCrLf & _
                              "           FROM      Po_Master_Export "

            ls_SQL = ls_SQL + "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo3 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                     ETAPort = ETAPort3 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory3 , " & vbCrLf & _
                              "                     week = 3 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo4 , "

            ls_SQL = ls_SQL + "                     ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                     ETAPort = ETAPort4 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory4 , " & vbCrLf & _
                              "                     week = 4 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo5 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                     ETAPort = ETAPort5 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory5 , "

            ls_SQL = ls_SQL + "                     week = 5 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "         ) POM " & vbCrLf & _
                              "         LEFT JOIN PO_Detail_Export POD ON POM.PONO = POD.PONO " & vbCrLf & _
                              "                                           AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                           AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN ( SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo1 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                             ETAPort = ETAPort1 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory1 , "

            ls_SQL = ls_SQL + "                             week = 1 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo2 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                             ETAPort = ETAPort2 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory2 , " & vbCrLf & _
                              "                             week = 2 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     UNION ALL "

            ls_SQL = ls_SQL + "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo3 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                             ETAPort = ETAPort3 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory3 , " & vbCrLf & _
                              "                             week = 3 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo4 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor4 , "

            ls_SQL = ls_SQL + "                             ETAPort = ETAPort4 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory4 , " & vbCrLf & _
                              "                             week = 4 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo5 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                             ETAPort = ETAPort5 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory5 , " & vbCrLf & _
                              "                             week = 5 "

            ls_SQL = ls_SQL + "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                   ) PRM ON PRM.PONO = POD.PONO " & vbCrLf & _
                              "                            AND PRM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                            AND PRM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                            AND PRM.OrderNo = POM.OrderNo " & vbCrLf & _
                              "         LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO " & vbCrLf & _
                              "                                              AND PRD.AffiliateID = PRM.AffiliateID " & vbCrLf & _
                              "                                              AND PRD.SupplierID = PRM.SupplierID " & vbCrLf & _
                              "                                              AND PRD.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN ReceiveForwarder_Detail RD ON POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                 AND POD.SupplierID = RD.SupplierID "

            ls_SQL = ls_SQL + "                                                 AND POD.POno = RD.POno " & vbCrLf & _
                              "                                                 AND POM.OrderNo = RD.OrderNo " & vbCrLf & _
                              "                                                 AND POD.PartNo = RD.PartNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              "                                                       AND IM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                       AND IM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                       AND IM.POno = POD.POno " & vbCrLf & _
                              "                                                       AND IM.OrderNo = POM.OrderNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                              "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                              "                                                       AND ID.SupplierID = IM.SupplierID "

            ls_SQL = ls_SQL + "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                              "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                              "                                                       AND ID.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                              "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                              "  WHERE  ISNULL(PRM.OrderNO, '') <> ''  " & vbCrLf & _
                              "  AND RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) = '" & period.Text & "'" & vbCrLf & _
                              "  ------------------------------------------------------------------------------------------------------ " & vbCrLf & _
                              "  UNION ALL -- DIFF " & vbCrLf & _
                              "  SELECT  DISTINCT "

            ls_SQL = ls_SQL + " 		ColNo = '', " & vbCrLf & _
                              "         idx = '3' , " & vbCrLf & _
                              "         AffiliateID = '', " & vbCrLf & _
                              "         AffiliateName = '', " & vbCrLf & _
                              "         PartNo ='', " & vbCrLf & _
                              "         PartName = '', " & vbCrLf & _
                              "         MOQ = '', " & vbCrLf & _
                              "         UOM = '', " & vbCrLf & _
                              "         cls = 'DIFFERENCE' , " & vbCrLf & _
                              "         OrderNo ='', " & vbCrLf & _
                              "         FirmQty = 0 , "

            ls_SQL = ls_SQL + "         Curr = '',  " & vbCrLf & _
                              "         Price = ISNULL(ID.Price, 0) , " & vbCrLf & _
                              "         Amount = ISNULL(ID.Amount, 0), " & vbCrLf & _
                              "         AFF = POM.AffiliateID, " & vbCrLf & _
                              "         ORD = POM.OrderNo, " & vbCrLf & _
                              "         PART = POD.PartNo " & vbCrLf & _
                              "         ,PERIOD = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) " & vbCrLf & _
                              "  FROM   ( SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo1 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                     ETAPort = ETAPort1 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory1 , "

            ls_SQL = ls_SQL + "                     week = 1 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo2 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                     ETAPort = ETAPort2 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory2 , " & vbCrLf & _
                              "                     week = 2 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL "

            ls_SQL = ls_SQL + "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo3 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                     ETAPort = ETAPort3 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory3 , " & vbCrLf & _
                              "                     week = 3 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo4 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor4 , "

            ls_SQL = ls_SQL + "                     ETAPort = ETAPort4 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory4 , " & vbCrLf & _
                              "                     week = 4 " & vbCrLf & _
                              "           FROM      Po_Master_Export " & vbCrLf & _
                              "           UNION ALL " & vbCrLf & _
                              "           SELECT    * , " & vbCrLf & _
                              "                     OrderNO = OrderNo5 , " & vbCrLf & _
                              "                     ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                     ETAPort = ETAPort5 , " & vbCrLf & _
                              "                     ETAFactory = ETAFactory5 , " & vbCrLf & _
                              "                     week = 5 "

            ls_SQL = ls_SQL + "           FROM      Po_Master_Export " & vbCrLf & _
                              "         ) POM " & vbCrLf & _
                              "         LEFT JOIN PO_Detail_Export POD ON POM.PONO = POD.PONO " & vbCrLf & _
                              "                                           AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                           AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN ( SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo1 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                             ETAPort = ETAPort1 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory1 , " & vbCrLf & _
                              "                             week = 1 "

            ls_SQL = ls_SQL + "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo2 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                             ETAPort = ETAPort2 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory2 , " & vbCrLf & _
                              "                             week = 2 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , "

            ls_SQL = ls_SQL + "                             OrderNO = OrderNo3 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                             ETAPort = ETAPort3 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory3 , " & vbCrLf & _
                              "                             week = 3 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo4 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                             ETAPort = ETAPort4 , "

            ls_SQL = ls_SQL + "                             ETAFactory = ETAFactory4 , " & vbCrLf & _
                              "                             week = 4 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo5 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor5 , " & vbCrLf & _
                              "                             ETAPort = ETAPort5 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory5 , " & vbCrLf & _
                              "                             week = 5 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export "

            ls_SQL = ls_SQL + "                   ) PRM ON PRM.PONO = POD.PONO " & vbCrLf & _
                              "                            AND PRM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                            AND PRM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                            AND PRM.OrderNo = POM.OrderNo " & vbCrLf & _
                              "         LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO " & vbCrLf & _
                              "                                              AND PRD.AffiliateID = PRM.AffiliateID " & vbCrLf & _
                              "                                              AND PRD.SupplierID = PRM.SupplierID " & vbCrLf & _
                              "                                              AND PRD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		LEFT JOIN PO_DetailUpload_export PDU ON PDU.PONO = POM.PONO " & vbCrLf & _
                              "                                                 AND PDU.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                 AND PDU.SupplierID = POM.SupplierID "

            ls_SQL = ls_SQL + "                                                 AND PDU.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN ReceiveForwarder_Detail RD ON POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                 AND POD.SupplierID = RD.SupplierID " & vbCrLf & _
                              "                                                 AND POD.POno = RD.POno " & vbCrLf & _
                              "                                                 AND POM.OrderNo = RD.OrderNo " & vbCrLf & _
                              "                                                 AND POD.PartNo = RD.PartNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              "                                                       AND IM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                       AND IM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                       AND IM.POno = POD.POno " & vbCrLf & _
                              "                                                       AND IM.OrderNo = POM.OrderNo "

            ls_SQL = ls_SQL + "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                              "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                              "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                              "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                              "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                              "                                                       AND ID.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                              "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                              "  WHERE  ISNULL(POM.OrderNO, '') <> '' " & vbCrLf & _
                              "  AND RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) = '" & period.Text & "'" & vbCrLf & _
                              "   )x " & vbCrLf & _
                              " -------------------------------------------------------------------------------------------------------  " & vbCrLf & _
                              " WHERE ORD <> '' " & vbCrLf & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + " ORDER BY AFF ,ORD, PART ,idx  "


            'If cboAffiliateCode.Text <> clsGlobal.gs_All And cboAffiliateCode.Text <> "" Then
            '    ls_SQL = ls_SQL + " and POM.AffiliateID = '" & cboAffiliateCode.Text & "'" & vbCrLf
            'End If

            'If cboPart.Text <> clsGlobal.gs_All And cboPart.Text <> "" Then
            '    ls_SQL = ls_SQL + " and POD.PartNo = '" & cboPart.Text & "'" & vbCrLf
            'End If

            'If txtorderno.Text <> "" Then
            '    ls_SQL = ls_SQL + " And POM.OrderNo = '" & Trim(txtorderno.Text) & "'" & vbCrLf
            'End If


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub Excel()
        Call GridLoadExcel()
        FileName = "TemplateExportOrderHistory.xlsx"
        FilePath = Server.MapPath("~\Template\" & FileName)
        Call epplusExportHeaderExcel(FilePath, "", dtHeader, "A:17", "")
    End Sub

    Private Sub epplusExportHeaderExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try

            Dim NewFileName As String = Server.MapPath("~\ProgressReportExport\TemplateExportOrderHistory.xlsx")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet

            ws = exl.Workbook.Worksheets("ExportOrderHistory")
            Dim irow As Integer = 0
            Dim iRowTmp As Integer = 0
            Dim icol As Integer = 0

            iRowTmp = 5
            For irow = 0 To pData.Rows.Count - 1
                If pData.Rows.Count > 0 Then
                    ws.Cells("A" & iRowTmp).Value = pData.Rows(irow)("ColNo")
                    ws.Cells("B" & iRowTmp).Value = pData.Rows(irow)("AffiliateID")
                    ws.Cells("C" & iRowTmp).Value = pData.Rows(irow)("AffiliateName")
                    ws.Cells("D" & iRowTmp).Value = pData.Rows(irow)("Partno")
                    ws.Cells("E" & iRowTmp).Value = pData.Rows(irow)("PartName")
                    ws.Cells("F" & iRowTmp).Value = pData.Rows(irow)("MOQ")
                    ws.Cells("G" & iRowTmp).Value = pData.Rows(irow)("UOM")
                    ws.Cells("H" & iRowTmp).Value = pData.Rows(irow)("Cls")
                    ws.Cells("I" & iRowTmp).Value = pData.Rows(irow)("orderno")
                    ws.Cells("J" & iRowTmp).Value = pData.Rows(irow)("firmqty")
                    ws.Cells("K" & iRowTmp).Value = pData.Rows(irow)("curr")
                    ws.Cells("L" & iRowTmp).Value = pData.Rows(irow)("price")
                    ws.Cells("M" & iRowTmp).Value = pData.Rows(irow)("Amount")

                    'ALIGNMENT
                    ws.Cells("A" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                    ws.Cells("B" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("C" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("D" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("E" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("F" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("G" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("H" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("I" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("J" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("K" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("L" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("M" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right


                    'WIDTH
                    ws.Column(2).Width = 11
                End If
                iRowTmp = iRowTmp + 1
            Next

            Dim rgAll As ExcelRange = ws.Cells(5, 1, iRowTmp - 1, 13)
            EpPlusDrawAllBorders(rgAll)

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

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

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)

        Select Case pAction
            Case "gridload"
                Call up_GridLoad()

                If grid.VisibleRowCount = 0 Then
                    Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                End If
            Case "excel"
                Call Excel()

        End Select
    End Sub
End Class