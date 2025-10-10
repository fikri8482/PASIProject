Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing
Imports System.Transactions
Imports OfficeOpenXml
Imports System.IO

Public Class DeliveryToAffListExport
    Inherits System.Web.UI.Page

    '-----------------------------------------------------
    Private grid_Renamed As ASPxGridView
    Private mergedCells As New Dictionary(Of GridViewDataColumn, TableCell)()
    Private cellRowSpans As New Dictionary(Of TableCell, Integer)()
    '-----------------------------------------------------

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pOrderNo As String
    Dim pSupplierCode As String
    Dim pSupplierName As String

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim param As String = ""
            If (Not IsPostBack) AndAlso (Not IsCallback) Then

                Session("MenuDesc") = "DELIVERY, RECEIVING AND SHIPPING CONFIRMATION"
                If IsNothing(Request.QueryString("prm")) Then
                    param = ""
                Else
                    param = Request.QueryString("prm").ToString()
                End If

                If param = "" Then
                    Call up_fillcombo()
                    lblerrmessage.Text = ""

                    grid.JSProperties("cpdtfrom") = Format(Now, "01 MMM yyyy")
                    grid.JSProperties("cpdtto") = Format(Now, "dd MMM yyyy")
                    grid.JSProperties("cpdt1") = Format(Now, "01 MMM yyyy")
                    grid.JSProperties("cpreceive") = "ALL"

                ElseIf param <> "" And Session("GOTOStatus") = "4" Then
                    lblerrmessage.Text = ""
                    pAffiliateCode = Split(param, "|")(1)
                    pAffiliateName = Split(param, "|")(2)
                    pSupplierCode = Split(param, "|")(3)
                    pSupplierName = Split(param, "|")(4)
                    pOrderNo = Split(param, "|")(5)

                    If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"

                    cboaffiliate.Text = pAffiliateCode
                    txtaffiliate.Text = pAffiliateName
                    cbosupplier.Text = pSupplierCode
                    txtsupplier.Text = pSupplierName
                    txtorderno.Text = pOrderNo

                    Call up_GridLoad()
                    Session("pCheckError") = "1"

                    Session.Remove("EmergencyUrl")
                    btnsubmenu.Text = "BACK"

                ElseIf param <> "" And Session("GOTOStatus") = "5" Then
                    lblerrmessage.Text = ""
                    pAffiliateCode = Split(param, "|")(1)
                    pAffiliateName = Split(param, "|")(2)
                    pSupplierCode = Split(param, "|")(3)
                    pSupplierName = Split(param, "|")(4)
                    pOrderNo = Split(param, "|")(5)

                    If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"

                    cboaffiliate.Text = pAffiliateCode
                    txtaffiliate.Text = pAffiliateName
                    cbosupplier.Text = pSupplierCode
                    txtsupplier.Text = pSupplierName
                    txtorderno.Text = pOrderNo

                    Call up_GridLoad()
                    Session("pCheckError") = "1"

                    Session.Remove("EmergencyUrl")
                    btnsubmenu.Text = "BACK"

                ElseIf param <> "" And Session("GOTOStatus") = "empat" Then
                    lblerrmessage.Text = ""
                    pAffiliateCode = Split(param, "|")(1)
                    pAffiliateName = Split(param, "|")(2)
                    pSupplierCode = Split(param, "|")(3)
                    pSupplierName = Split(param, "|")(4)
                    pOrderNo = Split(param, "|")(5)

                    If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"

                    cboaffiliate.Text = pAffiliateCode
                    txtaffiliate.Text = pAffiliateName
                    cbosupplier.Text = pSupplierCode
                    txtsupplier.Text = pSupplierName
                    txtorderno.Text = pOrderNo

                    Call up_GridLoad()
                    Session("pCheckError") = "1"

                    Session.Remove("EmergencyUrl")
                    btnsubmenu.Text = "BACK"

                ElseIf param <> "" And Session("GOTOStatus") = "lima" Then
                    lblerrmessage.Text = ""
                    pAffiliateCode = Split(param, "|")(1)
                    pAffiliateName = Split(param, "|")(2)
                    pSupplierCode = Split(param, "|")(3)
                    pSupplierName = Split(param, "|")(4)
                    pOrderNo = Split(param, "|")(5)

                    If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"

                    cboaffiliate.Text = pAffiliateCode
                    txtaffiliate.Text = pAffiliateName
                    cbosupplier.Text = pSupplierCode
                    txtsupplier.Text = pSupplierName
                    txtorderno.Text = pOrderNo

                    Call up_GridLoad()
                    Session("pCheckError") = "1"

                    Session.Remove("EmergencyUrl")
                    btnsubmenu.Text = "BACK"
                End If
            End If



        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

    '#Region "PROCEDURE"
    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'AFFILIATE
        ls_sql = "SELECT distinct SupplierID = '" & clsGlobal.gs_All & "', SupplierName = '" & clsGlobal.gs_All & "' from ms_supplier " & vbCrLf & _
                 "UNION Select SupplierID = RTRIM(SupplierID) ,SupplierName = RTRIM(SupplierName) FROM dbo.ms_supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbosupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierID")
                .Columns(0).Width = 70
                .Columns.Add("SupplierName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtsupplier.Text = clsGlobal.gs_All
                .TextField = "SupplierID"
                .DataBind()
            End With
            sqlConn.Close()

            'PartNo
            ls_sql = "SELECT distinct PartNo = '" & clsGlobal.gs_All & "', PartName = '" & clsGlobal.gs_All & "' from MS_Parts " & vbCrLf & _
                "Union all SELECT PartNo = RTRIM(PartNo) ,PartName = RTRIM(PartName) FROM MS_Parts " & vbCrLf
            sqlConn.Open()

            Dim sqlDAA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds1 As New DataSet
            sqlDAA.Fill(ds1)

            With cbopart
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds1.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 70
                .Columns.Add("PartName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtpart.Text = clsGlobal.gs_All
                .TextField = "Partno"
                .DataBind()
            End With

            sqlConn.Close()

            'AFFILIATE
            sqlConn.Open()
            ls_sql = "SELECT distinct AffiliateID = '" & clsGlobal.gs_All & "', AffiliateName = '" & clsGlobal.gs_All & "' from MS_Affiliate " & vbCrLf & _
                     "UNION Select AffiliateID = RTRIM(AffiliateID) ,AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate  where isnull(overseascls, '0') = '1'" & vbCrLf
            Dim sqlDA2 As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds2 As New DataSet
            sqlDA2.Fill(ds2)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds2.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 70
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtaffiliate.Text = clsGlobal.gs_All
                .TextField = "AffiliateID"
                .DataBind()
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            'PLAN DELIVERY DATE
            If checkbox1.Checked = True Then
                ls_Filter = ls_Filter + " AND POM.ETDVendor <= '" & Format(dt1.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            'Supplier Already Deliver
            If rbdeliver.Value = "YES" Then
                ls_Filter = ls_Filter + " AND ISNULL(RM.SuratJalanNo,'') <> '' " & vbCrLf
            ElseIf rbdeliver.Value = "NO" Then
                ls_Filter = ls_Filter + " AND ISNULL(RM.SuratJalanNo,'') = '' " & vbCrLf
            End If

            'remaining Receiving Qty
            If rbreceiving.Value = "YES" Then
                ls_Filter = ls_Filter + " AND ISNULL(DSD.DOQty, 0) - ( ISNULL(RD.GoodRecQty, 0) + ISNULL(RD.DefectRecQty, 0) ) > 0 " & vbCrLf
            ElseIf rbreceiving.Value = "NO" Then
                ls_Filter = ls_Filter + " AND ISNULL(DSD.DOQty, 0) - ( ISNULL(RD.GoodRecQty, 0) + ISNULL(RD.DefectRecQty, 0) ) = 0 " & vbCrLf
            End If

            'DIFF QTY
            If rbdiff.Value = "YES" Then
                ls_Filter = ls_Filter + " AND ISNULL(RD.DOQty, 0) <> ( ISNULL(RD.GoodRecQty, 0) + ISNULL(RD.DefectRecQty, 0)) " & vbCrLf
            ElseIf rbdiff.Value = "NO" Then
                ls_Filter = ls_Filter + " AND ISNULL(RD.DOQty, 0) = ( ISNULL(RD.GoodRecQty, 0) + ISNULL(RD.DefectRecQty, 0)) " & vbCrLf
            End If

            'ALREADY SHIPPING
            If rbshipping.Value = "YES" Then
                ls_Filter = ls_Filter + " AND ISNULL(SH.ShippingInstructionNo,'') <> '' " & vbCrLf
            ElseIf rbshipping.Value = "NO" Then
                ls_Filter = ls_Filter + " AND ISNULL(SH.ShippingInstructionNo,'') = '' " & vbCrLf
            End If

            'PARTNO
            If cbopart.Text <> clsGlobal.gs_All And cbopart.Text <> "" Then
                ls_Filter = ls_Filter + "AND RD.PartNo = '" & Trim(cbopart.Text) & "'" & vbCrLf
            End If

            'AFF
            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_Filter = ls_Filter + " AND RD.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf
            End If

            'SUPP
            If cbosupplier.Text <> clsGlobal.gs_All And cbosupplier.Text <> "" Then
                ls_Filter = ls_Filter + " AND RD.SupplierID = '" & Trim(cbosupplier.Text) & "'" & vbCrLf
            End If

            'ORDERNO
            If txtorderno.Text <> "" Then
                ls_Filter = ls_Filter + "and RD.PONo LIKE '%" & txtorderno.Text & "%'" & vbCrLf
            End If

            'RECEIVE DATE
            If checkbox2.Checked = True Then
                ls_Filter = ls_Filter + " AND CONVERT(date,RM.RECEIVEDATE) between '" & Format(dtfrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtto.Value, "yyyy-MM-dd") & "'" & vbCrLf
            End If

            
            'ls_SQL = "   SELECT * FROM (   " & vbCrLf & _
            '      "   SELECT coldetail2 = coldetail, coldetailname2 = coldetailname, coldetail = colprint,coldetailname = colprintname,colno=CONVERT(char,ROW_NUMBER() OVER(ORDER BY H_ORDERNO,H_SJRECEIVING,H_SJ,H_PARTNO,URUT)),   " & vbCrLf & _
            '      "   		ACT = ACT, colperiod,colaffiliatecode,colaffiliatename,   " & vbCrLf & _
            '      "           coldeliverylocationcode,coldeliverylocationname,colorderno,colsuppliercode,   " & vbCrLf & _
            '      "           colsuppliername,colplandeldate,coldeldate,colsj,colpartno,colpartname,coluom,   " & vbCrLf & _
            '      "           coldeliveryqty,colgood,coldefect,colremaining,colreceivedate,H_SJRECEIVING,   " & vbCrLf & _
            '      "           H_SJ,H_ORDERNO,H_PARTNO,URUT,S7,S8,S9, FWDID, LabelNo   " & vbCrLf & _
            '      "   FROM (   " & vbCrLf & _
            '      "   SELECT DISTINCT   " & vbCrLf & _
            '      "   colprint = 'GoodReceivingReportExport.aspx?prm=' + RTRIM(ISNULL(RM.AffiliateID,'')) + '|' + RTRIM(ISNULL(RM.SupplierID,'')) + '|' + RTRIM(ISNULL(RM.SuratJalanNo,'')),  " & vbCrLf & _
            '      "   colprintName = CASE WHEN ISNULL(RM.SuratJalanNo,'') = '' THEN '' ELSE 'PRINT' END,  "

            'ls_SQL = ls_SQL + "           coldetail = 'ReceivingEntryExport.aspx?prm=' + Rtrim(RM.OrderNo)    " & vbCrLf & _
            '                  "   					+ '|' + RTRIM(ISNULL(DSM.SuratJalanNo, ''))   " & vbCrLf & _
            '                  "   					+ '|' + RTRIM(RM.OrderNo)   " & vbCrLf & _
            '                  "   					+ '|' + RTRIM(DSM.AffiliateID) + '|'    " & vbCrLf & _
            '                  "   					+ RTRIM(DSM.SupplierID) + '|' + Rtrim(POM.PONO),     " & vbCrLf & _
            '                  "           coldetailname = CASE WHEN ISNULL(RM.SuratJalanNo,'') = '' THEN 'RECEIVE' ELSE 'DETAIL RECEIVE' END ,   " & vbCrLf & _
            '                  "           colno = '' ,   " & vbCrLf & _
            '                  "           ACT = (CASE WHEN isnull(RM.ExcelCls,0) = 0 THEN 0 ELSE 0 END),  " & vbCrLf & _
            '                  "           colperiod = (CONVERT(CHAR(7),CONVERT(DATETIME,isnull(POM.Period,'')),121)) ,   " & vbCrLf & _
            '                  "           colaffiliatecode = ISNULL(DSM.AffiliateID,'') ,   " & vbCrLf & _
            '                  "           colaffiliatename = ISNULL(MA.AffiliateName,'') ,   "

            'ls_SQL = ls_SQL + "           coldeliverylocationcode = POM.ForwarderID ,   " & vbCrLf & _
            '                  "           coldeliverylocationname = ISNULL(MF.ForwarderName,'') ,   " & vbCrLf & _
            '                  "           colorderno = RM.OrderNo ,   " & vbCrLf & _
            '                  "           colsuppliercode = ISNULL(DSM.SupplierID,'') ,   " & vbCrLf & _
            '                  "           colsuppliername = ISNULL(MS.SupplierName,'') ,   " & vbCrLf & _
            '                  "           colplandeldate = Convert(Char(12), convert(Datetime, POM.ETDVendor1),121) ,   " & vbCrLf & _
            '                  "           coldeldate = Convert(Char(12), convert(Datetime, isnull(DSM.DeliveryDate,'')),121) ,   " & vbCrLf & _
            '                  "           colsj = ISNULL(DSD.SuratJalanNo,'') ,   " & vbCrLf & _
            '                  "           colpartno = '' ,   " & vbCrLf & _
            '                  "           colpartname = '',   " & vbCrLf & _
            '                  "           coluom = '' ,   "

            'ls_SQL = ls_SQL + "           coldeliveryqty = '' ,   " & vbCrLf & _
            '                  "           colgood = '' ,   " & vbCrLf & _
            '                  "           coldefect = '' ,   " & vbCrLf & _
            '                  "           colremaining = '' ,   " & vbCrLf & _
            '                  "           colreceivedate = Convert(Char(12), convert(Datetime, isnull(RM.ReceiveDate,'')),121),   " & vbCrLf & _
            '                  "           H_SJRECEIVING = isnull(RM.SuratjalanNo,''),   " & vbCrLf & _
            '                  "           H_SJ = isnull(DSD.SuratJalanNo,''),   " & vbCrLf & _
            '                  "           H_ORDERNO = RM.OrderNo,   " & vbCrLf & _
            '                  "           H_PARTNO = '',   " & vbCrLf & _
            '                  "           URUT = '0'   " & vbCrLf & _
            '                  "           ,S7 = CASE WHEN ISNULL(DSM.SuratJalanNo,'') <> '' THEN Convert(Char(16), convert(Datetime, DSM.DeliveryDate),20) + '   ' +  CONVERT(char(3),RTRIM(isnull(DSM.EntryUser,''))) ELSE '' end  "

            'ls_SQL = ls_SQL + "           ,S8 = CASE WHEN ISNULL(RM.SuratJalanNo,'') <> '' THEN Convert(Char(16), convert(Datetime, RM.ReceiveDate),20) + '   ' + CONVERT(char(3),RTRIM(isnull(RM.EntryUser,''))) ELSE '' end  " & vbCrLf & _
            '                  "           ,S9 = CASE WHEN ISNULL(SH.ShippingInstructionNo,'') <> '' THEN Convert(Char(16), convert(Datetime, SH.EntryDate),20) + '   ' + CONVERT(char(3),RTRIM(isnull(SH.EntryUser,''))) ELSE '' end  " & vbCrLf & _
            '                  " 	         ,FWDID = Isnull(RM.ForwarderID,''), LabelNo = ''  " & vbCrLf & _
            '                  "   FROM    DOSupplier_Detail_Export DSD   " & vbCrLf & _
            '                  "           INNER JOIN DOSupplier_DetailBox_Export DDB ON DDB.SuratJalanNo = DSD.SuratJalanNo and DDB.AffiliateID = DSD.AffiliateID and DDB.SupplierID = DSD.SupplierID " & vbCrLf & _
            '                  "           LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo   " & vbCrLf & _
            '                  "                                                     AND DSM.AffiliateID = DSD.AffiliateID   " & vbCrLf & _
            '                  "                                                     AND DSM.SupplierID = DSD.SupplierID   " & vbCrLf & _
            '                  "                                                     AND DSM.PONO = DSD.PONO   " & vbCrLf & _
            '                  "                                                     AND DSM.OrderNo = DSD.OrderNo " & vbCrLf & _
            '                  "           LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO   " & vbCrLf & _
            '                  "                                             AND POD.AffiliateID = DSM.AffiliateID   "

            'ls_SQL = ls_SQL + "                                             AND POD.SupplierID = DSM.SupplierID   " & vbCrLf & _
            '                  "                                             AND POD.PartNo = DSD.PartNo   " & vbCrLf & _
            '                  "   		LEFT JOIN PO_Master_Export POM ON POM.PONO = POD.PONO   " & vbCrLf & _
            '                  "   										  AND POM.AffiliateID = POD.AffiliateID   " & vbCrLf & _
            '                  "   										  AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
            '                  "   		                                  AND POM.OrderNo1 = DSM.OrderNo   " & vbCrLf & _
            '                  "   		LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo   " & vbCrLf & _
            '                  "                                                   AND DSD.affiliateID = RM.affiliateID   " & vbCrLf & _
            '                  "                                                   AND DSD.SupplierID = RM.SupplierID  " & vbCrLf & _
            '                  "                                                   AND DSD.OrderNo = RM.OrderNo   " & vbCrLf & _
            '                  "           LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo   "

            'ls_SQL = ls_SQL + "                                                   AND RM.AffiliateID = RD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND RM.SupplierID = RD.SupplierID   " & vbCrLf & _
            '                  "                                                   AND RM.PONO = RD.PONO   " & vbCrLf & _
            '                  "                                                   AND DSD.PartNo = RD.PartNo   " & vbCrLf & _
            '                  "                                                   AND DSD.PONO = RD.PONO  " & vbCrLf & _
            '                  "                                                   AND RM.OrderNo = RD.OrderNo 	  " & vbCrLf & _
            '                  "   		LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID    " & vbCrLf & _
            '                  "   		LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID   " & vbCrLf & _
            '                  "   		LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID   " & vbCrLf & _
            '                  "   		LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo   " & vbCrLf & _
            '                  "   		LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls   " & vbCrLf
            'ls_SQL = ls_SQL + "         LEFT JOIN (SELECT DISTINCT A.ShippingInstructionNo, A.affiliateID, B.OrderNo, B.SupplierID, A.ShippingInstructionDate, A.EntryUser, b.SuratJalanNo   " & vbCrLf & _
            '                  " 					      ,EntryDate = CONVERT(VARCHAR,A.ShippingInstructionDate) + ' ' + CONVERT(varchar(20),A.EntryDate,114)--(SELECT TOP 1 CONVERT(varchar(20),EntryDate,114) FROM ShippingInstruction_Detail WHERE A.affiliateID = AffiliateID AND A.ForwarderID = ForwarderID  AND A.ShippingInstructionNo = ShippingInstructionNo ORDER BY EntryDate DESC) " & vbCrLf & _
            '                  " 				     FROM ShippingInstruction_Master A  " & vbCrLf & _
            '                  " 					      LEFT JOIN ShippingInstruction_Detail B  ON A.affiliateID = B.AffiliateID  AND A.ForwarderID = B.ForwarderID  AND A.ShippingInstructionNo = B.ShippingInstructionNo " & vbCrLf & _
            '                  " 				  ) SH ON SH.AffiliateID = DSM.AffiliateID " & vbCrLf & _
            '                  "   						AND SH.SupplierID = DSM.SupplierID " & vbCrLf & _
            '                  "   						AND SH.OrderNo = DSM.OrderNo " & vbCrLf & _
            '                  "   						AND SH.SuratJalanNo = DSM.SuratJalanNo " & vbCrLf
            'ls_SQL = ls_SQL + "          /*LEFT JOIN (SELECT DISTINCT A.ShippingInstructionNo, A.affiliateID, B.OrderNo, B.SupplierID, A.ShippingInstructionDate, A.EntryUser, b.SuratJalanNo  " & vbCrLf & _
            '                  "  					FROM ShippingInstruction_Master A LEFT JOIN ShippingInstruction_Detail B  " & vbCrLf & _
            '                  "  					ON A.affiliateID = B.AffiliateID  " & vbCrLf & _
            '                  "  					AND A.ForwarderID = B.ForwarderID  " & vbCrLf & _
            '                  "  					AND A.ShippingInstructionNo = B.ShippingInstructionNo) SH ON SH.AffiliateID = DSM.AffiliateID  " & vbCrLf & _
            '                  "  					AND SH.SupplierID = DSM.SupplierID  " & vbCrLf & _
            '                  "  					AND SH.OrderNo = DSM.OrderNo  " & vbCrLf & _
            '                  "  					AND SH.SuratJalanNo = DSM.SuratJalanNo*/  " & vbCrLf & _
            '                  "   WHERE isnull(RM.SuratjalanNo,'') <> ''   " & vbCrLf & _
            '                  "  /*AND ISNULL(SH.ShippingInstructionNo,'') = ''*/  " & vbCrLf

            'ls_SQL = ls_SQL + ls_Filter

            'ls_SQL = ls_SQL + "   ) HEADER   " & vbCrLf & _
            '                  "   --=================== DETAIL ====================-   " & vbCrLf & _
            '                  "  )X  ORDER BY H_ORDERNO,H_SJRECEIVING,H_SJ,H_PARTNO,URUT, LabelNo"

            ls_SQL = "    SELECT * FROM (    " & vbCrLf & _
                     "    SELECT coldetail2 = coldetail, coldetailname2 = coldetailname, coldetail = colprint,coldetailname = colprintname,colno=CONVERT(char,ROW_NUMBER() OVER(ORDER BY H_ORDERNO,H_SJRECEIVING,H_SJ,H_PARTNO,URUT)),    " & vbCrLf & _
                     "    		  ACT = ACT, colperiod,colaffiliatecode,colaffiliatename,    " & vbCrLf & _
                     "           coldeliverylocationcode,coldeliverylocationname,colorderno,colsuppliercode,    " & vbCrLf & _
                     "           colsuppliername,colplandeldate,coldeldate,colsj,colpartno,colpartname,coluom,    " & vbCrLf & _
                     "           coldeliveryqty,colgood,coldefect,colremaining,colreceivedate,H_SJRECEIVING,    " & vbCrLf & _
                     "           H_SJ,H_ORDERNO,H_PARTNO,URUT,S7,S8,S9, S10,FWDID, LabelNo    " & vbCrLf & _
                     "    FROM (    " & vbCrLf & _
                     "    SELECT DISTINCT    " & vbCrLf & _
                     "    colprint = 'GoodReceivingReportExport.aspx?prm=' + RTRIM(ISNULL(RM.AffiliateID,'')) + '|' + RTRIM(ISNULL(RM.SupplierID,'')) + '|' + RTRIM(ISNULL(RM.SuratJalanNo,'')),   " & vbCrLf & _
                     "    colprintName = CASE WHEN ISNULL(RM.SuratJalanNo,'') = '' THEN '' ELSE 'PRINT' END,  " & vbCrLf

            ls_SQL = ls_SQL + "    coldetail = 'ReceivingEntryExport.aspx?prm=' + Rtrim(RM.OrderNo)     " & vbCrLf & _
                              "    					+ '|' + RTRIM(ISNULL(RM.SuratJalanNo, ''))    " & vbCrLf & _
                              "    					+ '|' + RTRIM(RM.OrderNo)    " & vbCrLf & _
                              "    					+ '|' + RTRIM(RM.AffiliateID) + '|'     " & vbCrLf & _
                              "    					+ RTRIM(RM.SupplierID) + '|' + Rtrim(POM.PONO),      " & vbCrLf & _
                              "            coldetailname = CASE WHEN ISNULL(RM.SuratJalanNo,'') = '' THEN 'RECEIVE' ELSE 'DETAIL RECEIVE' END ,    " & vbCrLf & _
                              "            colno = '' ,    " & vbCrLf & _
                              "            ACT = (CASE WHEN isnull(RM.ExcelCls,0) = 0 THEN 0 ELSE 0 END),   " & vbCrLf & _
                              "            colperiod = (CONVERT(CHAR(7),CONVERT(DATETIME,isnull(POM.Period,'')),121)) ,    " & vbCrLf & _
                              "            colaffiliatecode = ISNULL(RM.AffiliateID,'') ,   " & vbCrLf & _
                              "            colaffiliatename = ISNULL(MA.AffiliateName,'') , " & vbCrLf

            ls_SQL = ls_SQL + " 		   coldeliverylocationcode = POM.ForwarderID ,    " & vbCrLf & _
                              "            coldeliverylocationname = ISNULL(MF.ForwarderName,'') ,    " & vbCrLf & _
                              "            colorderno = RM.OrderNo ,    " & vbCrLf & _
                              "            colsuppliercode = ISNULL(RM.SupplierID,'') ,    " & vbCrLf & _
                              "            colsuppliername = ISNULL(MS.SupplierName,'') ,    " & vbCrLf & _
                              "            colplandeldate = Convert(Char(12), convert(Datetime, POM.ETDVendor1),121) ,    " & vbCrLf & _
                              "            coldeldate = Convert(Char(12), convert(Datetime, isnull(RM.ReceiveDate,'')),121) ,    " & vbCrLf & _
                              "            colsj = ISNULL(RD.SuratJalanNo,'') ,    " & vbCrLf & _
                              "            colpartno = '' ,    " & vbCrLf & _
                              "            colpartname = '',    " & vbCrLf & _
                              "            coluom = '' ,  " & vbCrLf

            ls_SQL = ls_SQL + " 		   coldeliveryqty = '' ,    " & vbCrLf & _
                              "            colgood = '' ,    " & vbCrLf & _
                              "            coldefect = '' ,    " & vbCrLf & _
                              "            colremaining = '' ,    " & vbCrLf & _
                              "            colreceivedate = Convert(Char(12), convert(Datetime, isnull(RM.ReceiveDate,'')),121),    " & vbCrLf & _
                              "            H_SJRECEIVING = isnull(RM.SuratjalanNo,''),    " & vbCrLf & _
                              "            H_SJ = isnull(RD.SuratJalanNo,''),    " & vbCrLf & _
                              "            H_ORDERNO = RM.OrderNo,    " & vbCrLf & _
                              "            H_PARTNO = '',    " & vbCrLf & _
                              "            URUT = '0'    " & vbCrLf & _
                              "            ,S7 = CASE WHEN ISNULL(RM.SuratJalanNo,'') <> '' THEN Convert(Char(16), convert(Datetime, RM.ReceiveDate),20) + '   ' +  CONVERT(char(3),RTRIM(isnull(RM.EntryUser,''))) ELSE '' end " & vbCrLf

            ls_SQL = ls_SQL + " 		   ,S8 = CASE WHEN ISNULL(RM.SuratJalanNo,'') <> '' THEN Convert(Char(16), convert(Datetime, RM.ReceiveDate),20) + '   ' + CONVERT(char(3),RTRIM(isnull(RM.EntryUser,''))) ELSE '' end   " & vbCrLf & _
                              "            ,S9 = CASE WHEN ISNULL(SH.ShippingInstructionNo,'') <> '' THEN Convert(Char(16), convert(Datetime, SH.EntryDate),20) + '   ' + CONVERT(char(3),RTRIM(isnull(SH.EntryUser,''))) ELSE '' end   " & vbCrLf & _
                              "            ,S10 = CASE WHEN ISNULL(TD.ShippingInstructionNo,'') <> '' THEN Convert(Char(16), convert(Datetime, SH.EntryDate),20) + '   ' + CONVERT(char(3),RTRIM(isnull(SH.EntryUser,''))) ELSE '' end   " & vbCrLf & _
                              "  	       ,FWDID = Isnull(RM.ForwarderID,''), LabelNo = ''   " & vbCrLf & _
                              "    FROM     " & vbCrLf & _
                              " 		   ReceiveForwarder_Master RM  " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo   " & vbCrLf & _
                              " 												   AND RM.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                    AND RM.SupplierID = RD.SupplierID   " & vbCrLf & _
                              "                                                    AND RM.PONO = RD.PONO    " & vbCrLf & _
                              "                                                    AND RM.OrderNo = RD.OrderNo 	   " & vbCrLf & _
                              "            LEFT JOIN po_detail_Export POD ON POD.PONO = RM.PONO    " & vbCrLf

            ls_SQL = ls_SQL + "                                              AND POD.AffiliateID = RM.AffiliateID  " & vbCrLf & _
                              " 											 AND POD.SupplierID = RM.SupplierID    " & vbCrLf & _
                              "                                              AND POD.PartNo = RD.PartNo    " & vbCrLf & _
                              "    		LEFT JOIN PO_Master_Export POM ON POM.PONO = POD.PONO    " & vbCrLf & _
                              "    										  AND POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "    										  AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "    		                                  AND POM.OrderNo1 = RM.OrderNo   	   " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = RM.AffiliateID     " & vbCrLf & _
                              "    		LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID    " & vbCrLf & _
                              "    		LEFT JOIN ms_supplier MS ON MS.SupplierID = RM.SupplierID    " & vbCrLf & _
                              "    		LEFT JOIN MS_Parts MP ON MP.PartNo = RD.PartNo    " & vbCrLf

            ls_SQL = ls_SQL + "    		LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls    " & vbCrLf & _
                              "          LEFT JOIN (SELECT DISTINCT A.ShippingInstructionNo, A.affiliateID, B.OrderNo, B.SupplierID, A.ShippingInstructionDate, A.EntryUser, b.SuratJalanNo    " & vbCrLf & _
                              "  					      ,EntryDate = CONVERT(VARCHAR,A.ShippingInstructionDate) + ' ' + CONVERT(varchar(20),A.EntryDate,114) " & vbCrLf & _
                              "  				     FROM ShippingInstruction_Master A   " & vbCrLf & _
                              "  					      LEFT JOIN ShippingInstruction_Detail B  ON A.affiliateID = B.AffiliateID  AND A.ForwarderID = B.ForwarderID  AND A.ShippingInstructionNo = B.ShippingInstructionNo  " & vbCrLf & _
                              "  				  ) SH ON SH.AffiliateID = RM.AffiliateID  " & vbCrLf & _
                              "    						AND SH.SupplierID = RM.SupplierID  " & vbCrLf & _
                              "    						AND SH.OrderNo = RM.OrderNo  " & vbCrLf & _
                              "    						AND SH.SuratJalanNo = RM.SuratJalanNo  " & vbCrLf & _
                              "			LEFT JOIN dbo.Tally_Detail TD ON TD.AffiliateID = MA.AffiliateID " & vbCrLf & _
                              "				AND TD.ForwarderID = MF.ForwarderID " & vbCrLf & _
                              "				AND TD.OrderNo = RD.OrderNo " & vbCrLf & _
                              "				AND TD.PartNo = MP.PartNo " & vbCrLf & _
                              "				AND TD.ShippingInstructionNo = SH.ShippingInstructionNo " & vbCrLf & _
                              " LEFT JOIN dbo.Tally_Master TM ON TM.AffiliateID = MA.AffiliateID " & vbCrLf & _
                              "				AND	TM.ForwarderID = MF.ForwarderID " & vbCrLf & _
                              "				AND	TM.ShippingInstructionNo = TD.ShippingInstructionNo " & vbCrLf & _
                              "    WHERE isnull(RM.SuratjalanNo,'') <> ''     " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + "    ) HEADER    " & vbCrLf
            ls_SQL = ls_SQL + "    --=================== DETAIL ====================-    " & vbCrLf & _
                              "   )X  ORDER BY H_ORDERNO,H_SJRECEIVING,H_SJ,H_PARTNO,URUT, LabelNo " & vbCrLf & _
                              "  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    '#End Region

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        If btnsubmenu.Text = "BACK" And Session("GOTOStatus") = "4" Then
            Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
        ElseIf btnsubmenu.Text = "BACK" And Session("GOTOStatus") = "5" Then
            Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
        ElseIf btnsubmenu.Text = "BACK" And Session("GOTOStatus") = "empat" Then
            Response.Redirect("~/PurchaseOrderExport/POExportFinalApprovalList.aspx")
        ElseIf btnsubmenu.Text = "BACK" And Session("GOTOStatus") = "lima" Then
            Response.Redirect("~/PurchaseOrderExport/POExportFinalApprovalList.aspx")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If

        Session.Remove("GOTOStatus")
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_sql As String
        Dim ls_affiliateID As String
        Dim ls_SjNo As String
        Dim ls_supplierID As String
        Dim ls_pono As String
        Dim ls_testing As String

        If HF.Get("hfTest") = "1" Then 'send good receive
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                If grid.VisibleRowCount = 0 Then Exit Sub
                Session.Remove("msgDelivery")
                Using sqlTran As New TransactionScope
                    If e.UpdateValues.Count = 0 Then
                        Session("MsgDelivery") = "NOT SAVED"
                        Exit Sub
                    End If

                    For iRow = 0 To e.UpdateValues.Count - 1
                        ls_affiliateID = e.UpdateValues(iRow).NewValues("colaffiliatecode").ToString()
                        ls_SjNo = e.UpdateValues(iRow).NewValues("colsj").ToString()
                        ls_supplierID = e.UpdateValues(iRow).NewValues("colsuppliercode").ToString()
                        ls_pono = e.UpdateValues(iRow).NewValues("colorderno").ToString()

                        'delete data detail
                        ls_sql = " Update  dbo.ReceiveForwarder_Master set Excelcls = '1' WHERE SuratJalanNo = '" & ls_SjNo & "' AND SupplierID = '" & ls_supplierID & "' " & vbCrLf & _
                                 " AND AffiliateID = '" & ls_affiliateID & "' AND Orderno = '" & ls_pono & "' "
                        Dim sqlComm3 As New SqlCommand(ls_sql, sqlConn)
                        sqlComm3.ExecuteNonQuery()
                        sqlComm3.Dispose()

                    Next iRow
                    sqlTran.Complete()

                End Using
                sqlConn.Close()
                Call up_GridLoad()
            End Using
        ElseIf HF.Get("hfTest") = "0" Then 'shipping
            If e.UpdateValues.Count = 0 Then Exit Sub
            Dim pAffiliateID As String = ""
            Dim pOrderNo As String = ""
            Dim pSupplier As String = ""
            Dim pFwd As String = ""
            Dim ls_General As String = ""
            Dim ls_SuratJalanNo As String = ""
            Dim isSJ As Boolean = True

            'pAffiliateID = Trim(e.UpdateValues(0).NewValues("colaffiliatecode").ToString())
            'pOrderNo = Trim(e.UpdateValues(0).NewValues("colorderno").ToString())
            'pSupplier = Trim(e.UpdateValues(0).NewValues("colsuppliercode").ToString())
            'pFwd = Trim(e.UpdateValues(0).NewValues("FWDID").ToString())

            If e.UpdateValues.Count > 0 Then
                For i = 0 To e.UpdateValues.Count - 1
                    'If (e.UpdateValues(i).NewValues("ACT").ToString()) = 1 Then
                    '    pAffiliateID = pAffiliateID + ",'" & Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) & "'"
                    '    pOrderNo = pOrderNo + ",'" & Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) & "'"
                    '    pSupplier = pSupplier + ",'" & Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) & "'"
                    'End If
                    If ls_SuratJalanNo <> "" Then
                        'If ls_SuratJalanNo <> Trim(e.UpdateValues(i).NewValues("H_SJRECEIVING").ToString()) Then isSJ = False : ls_SuratJalanNo = ""
                        If ls_SuratJalanNo <> Trim(e.UpdateValues(i).NewValues("H_SJRECEIVING").ToString()) Then ls_SuratJalanNo = ls_SuratJalanNo & ",'" & Trim(e.UpdateValues(i).NewValues("H_SJRECEIVING").ToString()) & "'"
                    Else
                        ls_SuratJalanNo = "'" & Trim(e.UpdateValues(i).NewValues("H_SJRECEIVING").ToString()) & "'"
                    End If

                    If pFwd = "" Then
                        pFwd = Trim(e.UpdateValues(i).NewValues("FWDID").ToString())
                    Else
                        If Trim(e.UpdateValues(i).NewValues("FWDID").ToString()) <> pFwd Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "6041", clsMessage.MsgType.ErrorMessage)
                            grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("AA220Msg") = lblerrmessage.Text
                            Exit Sub
                        End If
                    End If
                    'colsj
                    If pAffiliateID = "" Then
                        pAffiliateID = Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString())
                        If ls_General = "" Then
                            ls_General = "'" & Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) + Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        Else
                            ls_General = ls_General + ",'" + Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) + Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        End If
                    Else
                        If Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) <> pAffiliateID Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "6040", clsMessage.MsgType.ErrorMessage)
                            grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("AA220Msg") = lblerrmessage.Text
                            Exit Sub
                        End If
                        If ls_General = "" Then
                            ls_General = "'" & Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) + Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        Else
                            ls_General = ls_General + ",'" + Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) + Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        End If
                    End If
                Next
            End If

            If pAffiliateID <> "" Then
                Session("SHAFFILIATEID") = pAffiliateID
                Session("SHORDERNO") = pOrderNo
                Session("SHSUPPLIERID") = pSupplier
                Session("SHFWD") = pFwd
                Session("isSJ") = ls_SuratJalanNo
                Session("SHGENERAL") = ls_General
                DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~/ShippingInstruction/ShippingInstructionToForwarder.aspx")
                'Response.Redirect("~/MainMenu.aspx")
            End If
        ElseIf HF.Get("hfTest") = "2" Then
            If e.UpdateValues.Count = 0 Then Exit Sub
            Dim pAffiliateID As String = ""
            Dim pOrderNo As String = ""
            Dim pSupplier As String = ""
            Dim pFwd As String = ""
            Dim ls_General As String = ""
            Dim ls_SuratJalanNo As String = ""
            Dim isSJ As Boolean = True

            If e.UpdateValues.Count > 0 Then
                For i = 0 To e.UpdateValues.Count - 1
                    If ls_SuratJalanNo <> "" Then
                        'If ls_SuratJalanNo <> Trim(e.UpdateValues(i).NewValues("H_SJRECEIVING").ToString()) Then isSJ = False : ls_SuratJalanNo = ""
                        If ls_SuratJalanNo <> Trim(e.UpdateValues(i).NewValues("H_SJRECEIVING").ToString()) Then ls_SuratJalanNo = ls_SuratJalanNo & ",'" & Trim(e.UpdateValues(i).NewValues("H_SJRECEIVING").ToString()) & "'"
                    Else
                        ls_SuratJalanNo = "'" & Trim(e.UpdateValues(i).NewValues("H_SJRECEIVING").ToString()) & "'"
                    End If

                    If pFwd = "" Then
                        pFwd = Trim(e.UpdateValues(i).NewValues("FWDID").ToString())
                    Else
                        If Trim(e.UpdateValues(i).NewValues("FWDID").ToString()) <> pFwd Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "6041", clsMessage.MsgType.ErrorMessage)
                            grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("AA220Msg") = lblerrmessage.Text
                            Exit Sub
                        End If
                    End If

                    If pAffiliateID = "" Then
                        pAffiliateID = Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString())
                        If ls_General = "" Then
                            ls_General = "'" & Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) + Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        Else
                            ls_General = ls_General + ",'" + Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) + Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        End If
                    Else
                        If Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) <> pAffiliateID Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "6040", clsMessage.MsgType.ErrorMessage)
                            grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("AA220Msg") = lblerrmessage.Text
                            Exit Sub
                        End If
                        If ls_General = "" Then
                            ls_General = "'" & Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) + Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        Else
                            ls_General = ls_General + ",'" + Trim(e.UpdateValues(i).NewValues("colorderno").ToString()) + Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        End If
                    End If
                Next
            End If

            If pAffiliateID <> "" Then
                Session("SHGENERAL") = ls_General
                'DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~/ShippingInstruction/ShippingInstructionToForwarder.aspx")
                'Response.Redirect("~/MainMenu.aspx")
            End If
        End If
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)

            If pAction = "gridExcel" Then GoTo keluar

            If pAction <> "send" Or pAction <> "gridExcel" Then
                Dim pPlan As String = Split(e.Parameters, "|")(1)
                Dim pSupplier As String = Split(e.Parameters, "|")(2)
                Dim pRemaining As String = Split(e.Parameters, "|")(3)
                Dim psj As String = Split(e.Parameters, "|")(4)
                Dim pDateFrom As String = Split(e.Parameters, "|")(5)
                Dim pDateTo As String = Split(e.Parameters, "|")(6)
                Dim pPart As String = Split(e.Parameters, "|")(7)
                Dim pOrder As String = Split(e.Parameters, "|")(8)
            End If
keluar:
            Select Case pAction
                Case "gridload"
                    'Call up_GridLoad(pPlan, pSupplierDeliver, pRemaining, psj, pDateFrom, pDateTo, pSupplier, pPart, pPoNo, pKanban)
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "update"
                    grid.UpdateEdit()
                Case "send"
                    Call up_GridLoad()
                Case "gridExcel"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = GetSummaryOutStanding()
                    FileName = "TemplateShipping.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:8", psERR)
                    End If
            End Select

EndProcedure:
            'Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Shipping Intruction" & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\ProgressReport\Import\" & tempFile & "")
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
                '.Cells(3, 4).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
                '.Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                .Cells("A8").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(8, 1, pData.Rows.Count + 7, 10).AutoFitColumns()

                .Cells(8, 6, pData.Rows.Count + 7, 6).Style.Numberformat.Format = "dd-mmm-yy"

                .Cells(8, 5, pData.Rows.Count + 7, 5).Style.Numberformat.Format = "#,##0"

                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 10)
                EpPlusDrawAllBorders(rgAll)

                'For irow = 0 To pData.Rows.Count - 1
                '    For icol = 1 To pData.Columns.Count
                '        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                '        If icol = 7 Or icol = 8 Or icol = 14 Or icol = 15 Or icol = 16 Or icol = 20 Or icol = 23 Or icol = 26 Or icol = 29 Then
                '            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "dd-mmm-yy"
                '        End If
                '        If icol = 11 Or icol = 13 Or icol = 18 Or icol = 19 Or icol = 21 Or icol = 28 Or icol = 30 Or icol = 25 Or icol = 34 Then
                '            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "#,##0"
                '        End If
                '    Next
                'Next

                'Dim rgAll As ExcelRange = .Cells(8, 1, irow + 8, 34)
                'EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

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

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If Not (e.DataColumn.FieldName = "coldetail" Or e.DataColumn.FieldName = "ACT") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        End If

        If e.DataColumn.FieldName = "ACT" Then
            If (e.GetValue("colaffiliatecode") = "") Then
                e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
                e.Cell.Controls("0").Controls.Clear()
            End If
        End If
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub

    Private Sub btndeliver_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btndeliver.Click
        Response.Redirect("~/DeliveryExport/DeliveryExportUpload.aspx")
    End Sub

    Private Function GetSummaryOutStanding() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()

                ls_sql = "  Select * from (  " & vbCrLf & _
                      " 	SELECT distinct   " & vbCrLf & _
                      " 		RTRIM(RD.OrderNo)OrderNo, RTRIM(RD.PartNo)PartNo, RTRIM(PartName)PartName, QtyBox = ISNULL(PMD.POQtyBox,PMP.QtyBox), GoodRecQty = Replace(CONVERT(char,isnull(RB.Box,0) * ISNULL(PMD.POQtyBox,ISNULL(PMP.QtyBox,0))),'.00','') ,    " & vbCrLf & _
                      " 		ReceiveDate, RM.AffiliateID, " & vbCrLf & _
                      " 		RTRIM(RM.SupplierID)SupplierID,  " & vbCrLf & _
                      " 		RM.SuratJalanNo,  " & vbCrLf & _
                      " 		LabelNo = isnull(Rtrim(RB.Label1) + '-' + Rtrim(RB.Label2),'')  " & vbCrLf & _
                      " 	FROM dbo.ReceiveForwarder_Master RM   " & vbCrLf & _
                      " 	LEFT JOIN dbo.ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratJalanNo   "

                ls_sql = ls_sql + " 		AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                                  " 		AND RM.AffiliateID = RD.AffiliateID   " & vbCrLf & _
                                  " 		AND RM.PONo = RD.PONo   " & vbCrLf & _
                                  " 	LEFT JOIN dbo.MS_Parts MP ON RD.PartNo = MP.PartNo   " & vbCrLf & _
                                  " 	LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo  " & vbCrLf & _
                                  " 		AND RB.SupplierID = RD.SupplierID   " & vbCrLf & _
                                  " 		AND RB.AffiliateID = RD.AffiliateID   " & vbCrLf & _
                                  " 		AND RB.PONo = RD.PONo   " & vbCrLf & _
                                  " 		AND RB.OrderNo = RD.OrderNo   " & vbCrLf & _
                                  " 		AND RB.PartNo = RD.PartNo   " & vbCrLf & _
                                  " 		AND RB.StatusDefect = '0' "

                ls_sql = ls_sql + " 	LEFT JOIN dbo.MS_PartMapping PMP ON RD.PartNo = PMP.PartNo and RM.AffiliateID = PMP.AffiliateID and RM.SupplierID = PMP.SupplierID   " & vbCrLf & _
                                  " 	LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
                                  " 	LEFT JOIN dbo.MS_Supplier MSS ON RM.SupplierID = MSS.SupplierID   " & vbCrLf & _
                                  " 	LEFT JOIN dbo.PO_Master_Export PME ON PME.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                                  " 		AND (RM.OrderNo =  PME.OrderNo1 or RM.OrderNo =  PME.OrderNo2  " & vbCrLf & _
                                  " 		or RM.OrderNo =  PME.OrderNo3 or RM.OrderNo =  PME.OrderNo4 or RM.OrderNo =  PME.OrderNo5)  "

                ls_sql = ls_sql + " 		and PME.SupplierID = RM.SupplierID AND PME.PONo = RM.PONo  " & vbCrLf & _
                                  "     LEFT JOIN dbo.PO_Detail_Export PMD on PMD.PONo = PME.PONo and PME.AffiliateID = PMD.AffiliateID and PME.SupplierID = PMD.SupplierID and PMD.PartNo = RD.PartNo " & vbCrLf & _
                                  " 	WHERE ISNULL(RD.OrderNo, '') <> ''    " & vbCrLf & _
                                  " 		and RTrim(RD.SuratJalanNo) + Rtrim(RD.AffiliateID) + Rtrim(RD.OrderNo)+RTRIM(RD.SupplierID)+RTRIM(RD.PartNo)  " & vbCrLf & _
                                  " 	NOT IN (SELECT DISTINCT RTrim(SuratJalanNo) + Rtrim(AffiliateID) + Rtrim(OrderNo)+RTRIM(SupplierID)+RTRIM(PartNo) From ShippingInstruction_Detail where  " & vbCrLf & _
                                  " 	suratjalanno = RD.SuratJalanNo and AffiliateID = RD.AffiliateID AND SupplierID = RD.SupplierID and partno = RD.Partno and orderno = RD.OrderNo)  " & vbCrLf

                If Session("SHGENERAL") <> "" Then
                    ls_sql = ls_sql + _
                        "           AND RTRIM(RM.OrderNo) + RTRIM(RM.AffiliateID) + RTRIM(RM.SupplierID) + RTRIM(RM.SuratJalanNo) in (" & Session("SHGENERAL") & ") " & vbCrLf
                End If

                ls_sql = ls_sql + ")x  Order By AffiliateID, LabelNo"

                Dim Cmd As New SqlCommand(ls_sql, cn)
                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 200
                da.Fill(dt)

                Return dt
            End Using


        Catch ex As Exception
            Return Nothing
        End Try
    End Function


End Class