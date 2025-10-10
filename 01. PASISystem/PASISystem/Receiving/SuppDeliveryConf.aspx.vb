Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView

Public Class SuppDeliveryConf
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
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_fillcombo()
                lblerrmessage.Text = ""
                grid.JSProperties("cpdtfrom") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdtto") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdt1") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdeliver") = "ALL"
                grid.JSProperties("cpreceive") = "ALL"
                grid.JSProperties("cpkanban") = "ALL"

            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())

        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)

        End Try
    End Sub

#Region "PROCEDURE"
    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'SupplierCode
        ls_sql = "SELECT distinct supplier_Code = '" & clsGlobal.gs_All & "', Supplier_Name = '" & clsGlobal.gs_All & "' from MS_Supplier " & vbCrLf & _
                 "UNION ALL Select Supplier_Code = RTRIM(supplierID) ,Supplier_Name = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbosupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Supplier_Code")
                .Columns(0).Width = 70
                .Columns.Add("Supplier_Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtsupplier.Text = clsGlobal.gs_All
                .TextField = "Supplier_Code"
                .DataBind()
            End With
            sqlConn.Close()

            'PartNo
            ls_sql = "SELECT distinct Partno = '" & clsGlobal.gs_All & "', Partname = '" & clsGlobal.gs_All & "' from MS_Parts " & vbCrLf & _
                "Union all SELECT Parno = RTRIM(PartNo) ,Partname = RTRIM(PartName) FROM MS_Parts " & vbCrLf
            sqlConn.Open()

            Dim sqlDAA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds1 As New DataSet
            sqlDAA.Fill(ds1)

            With cbopart
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds1.Tables(0)
                .Columns.Add("Partno")
                .Columns(0).Width = 70
                .Columns.Add("Partname")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtpart.Text = clsGlobal.gs_All
                .TextField = "Partno"
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoadOLD(ByVal pdateplan As Date, ByVal psupplierdeliver As String, ByVal premaining As String, ByVal psuratjalanno As String, ByVal pFrom As Date, ByVal pTo As Date, ByVal psupplier As String, ByVal ppart As String, ByVal ppono As String, ByVal pkanban As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "  SELECT * FROM (  " & vbCrLf & _
                  "  	SELECT coldetail, coldetailname, CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY h_poorder, h_KanbanCls,h_kanbanorder, H_SURATJALAN, h_idxorder)) AS colno,  " & vbCrLf & _
                  "  		colperiod, colaffiliatecode, colaffiliatename ,  " & vbCrLf & _
                  "  		colpono, colsuppliercode,colsuppliername, colpokanban, colkanbanno,  " & vbCrLf & _
                  "  		colplandeldate, coldeldate, colsj,colpartno,colpartname,  " & vbCrLf & _
                  "  		coluom,coldeliveryqty,colreceiveqty,coldefect,colremaining,colreceivedate= (CASE WHEN colreceivedate = '01 Jan 1900' THEN '' ELSE colreceivedate END) ,  " & vbCrLf & _
                  "  		colreceiveby,H_POORDER, H_IDXORDER ,H_KANBANORDER ,H_AFFILIATEORDER, H_KANBANCLS, " & vbCrLf & _
                  "  		H_PLANDELDATE,H_ALREADYDELIVER,H_REMAINING,H_SUPPLIER,H_RECEIVEDATE,H_SURATJALAN " & vbCrLf & _
                  "  	FROM (  " & vbCrLf & _
                  "  		SELECT DISTINCT    " & vbCrLf & _
                  "   				 coldetail = 'ReceivingEntry.aspx?prm='+CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(RM.entryDate,'')),106)   " & vbCrLf

            ls_SQL = ls_SQL + "                                                  + '|' +Rtrim(POM.SupplierID) + '|' +Rtrim(MS.SupplierName)   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(ISNULL(DOM.SuratJalanNo,''))   " & vbCrLf & _
                              "                                                  + '|' +Rtrim( CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106))   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),106))   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(ISNULL(KD.KanbanNo,''))   --+ '|' +Rtrim(CASE WHEN POD.KanbanCls = '0' THEN '-' ELSE ISNULL(KD.KanbanNo,'') END)  " & vbCrLf & _
                              "                                                  + '|' +Rtrim(POM.AffiliateID) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOD.PONO) , " & vbCrLf & _
                              "   				 coldetailname = CASE WHEN ISNULL(RD.GoodRecQty,0)=0 THEN 'RECEIVE' ELSE 'DETAIL' END,    " & vbCrLf & _
                              "  				 colno = '' ,   " & vbCrLf & _
                              "  				 colperiod = Right(Convert(char(11),Convert(datetime,POM.Period),106),8),   " & vbCrLf & _
                              "  				 colaffiliatecode = POM.AffiliateID ,    " & vbCrLf & _
                              "  				 colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
                              "  				 colpono = POM.PONo ,    				 colsuppliercode = POM.SupplierID ,    " & vbCrLf

            ls_SQL = ls_SQL + "  				 colsuppliername = MS.SupplierName ,    " & vbCrLf & _
                              "  				 colpokanban = CASE WHEN ISNULL(POD.KanbanCls,'0')='0' THEN 'NO' ELSE 'YES' END ,    " & vbCrLf & _
                              "  				 colkanbanno = '', --ISNULL(KD.KanbanNo,'') ,    " & vbCrLf & _
                              "  				 colplandeldate = '', --CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106) ,    " & vbCrLf & _
                              "  				 coldeldate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),106) ,    " & vbCrLf & _
                              "  				 colsj = ISNULL(DOM.SuratJalanNo,'') ,    " & vbCrLf & _
                              "  				 colpartno = '' , " & vbCrLf & _
                              "  				 colpartname = '' ,   " & vbCrLf & _
                              "  				 coluom = '' ,   " & vbCrLf & _
                              "  				 coldeliveryqty = '' ,   				 colreceiveqty = '' ,   " & vbCrLf & _
                              "                   coldefect = '' ,  " & vbCrLf

            ls_SQL = ls_SQL + "  				 colremaining = '' ,   " & vbCrLf & _
                              "  				 colreceivedate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(RM.entryDate,'')),106) ,   " & vbCrLf & _
                              "  				 colreceiveby = ISNULL(RM.EntryUser,''),  " & vbCrLf & _
                              "  				 pom.PONo H_POORDER, H_IDXORDER = 0, H_KANBANORDER = ISNULL(KD.KanbanNo,'-') ,   " & vbCrLf & _
                              "  				 H_AFFILIATEORDER = POM.AffiliateID , H_KANBANCLS = pod.KanbanCls,  " & vbCrLf & _
                              "  				 H_PLANDELDATE =  CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106), " & vbCrLf & _
                              "  				 H_ALREADYDELIVER = '', " & vbCrLf & _
                              "  				 H_REMAINING = CONVERT(CHAR,(SELECT SUM(ISNULL(A.DOQty,0)) " & vbCrLf & _
                              "                                                 From dbo.DOSupplier_Detail A " & vbCrLf & _
                              "                                                 WHERE(A.AffiliateID = DOM.AffiliateID) " & vbCrLf & _
                              "                                                 AND A.SupplierID = DOM.SupplierID " & vbCrLf & _
                              "                                                 AND A.KanbanNo = KM.KanbanNo " & vbCrLf & _
                              "                                                 AND A.PONo = POM.PONo) - " & vbCrLf & _
                              "                                             (SELECT ISNULL(SUM(ISNULL(A.GoodRecQty,0)) + SUM(ISNULL(A.DefectRecQty,0)),0) " & vbCrLf & _
                              "                                                 From dbo.ReceivePASI_Detail A  " & vbCrLf & _
                              "                                                 WHERE(A.AffiliateID =  DOM.AffiliateID) " & vbCrLf & _
                              "                                                 AND A.SupplierID = DOM.SupplierID " & vbCrLf & _
                              "                                                 AND A.KanbanNo = KM.KanbanNo " & vbCrLf & _
                              "                                                 AND A.PONo = POM.PONo )), " & vbCrLf & _
                              "  				 H_SUPPLIER = POM.SupplierID, " & vbCrLf & _
                              "  				 H_RECEIVEDATE = CONVERT(CHAR(10), CONVERT(DATETIME,ISNULL(RM.ReceiveDate,'')),120),  " & vbCrLf & _
                              "  				 H_SURATJALAN = ISNULL(DOM.SuratJalanNo,'') " & vbCrLf

            ls_SQL = ls_SQL + "  				 --,H_PARTCODE = pod.PartNo " & vbCrLf & _
                              "  				  " & vbCrLf & _
                              "  		 FROM    dbo.PO_Master POM   " & vbCrLf & _
                              "  				 INNER JOIN dbo.PO_Detail POD ON POD.PONo = POM.PONo    " & vbCrLf & _
                              "  												  AND POD.SupplierID = POM.SupplierID                                              " & vbCrLf & _
                              "  												  AND POD.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                              "  												  AND POM.DeliveryByPasiCls = 1    				INNER JOIN dbo.DOSupplier_Detail DOD ON DOD.PONo = POD.PONo    " & vbCrLf & _
                              "  													 AND DOD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "  													 AND DOD.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "  													 AND DOD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  				INNER JOIN dbo.DOSupplier_Master DOM ON DOM.SuratJalanNo =  DOD.SuratJalanNo    " & vbCrLf

            ls_SQL = ls_SQL + "  													 AND DOM.SupplierID = DOD.SupplierID    " & vbCrLf & _
                              "  													 AND DOM.AffiliateID = DOD.AffiliateID    " & vbCrLf & _
                              "  				LEFT JOIN dbo.ReceivePASI_Detail RD ON RD.SuratJalanNo = DOM.SuratJalanNo    " & vbCrLf & _
                              "    													AND RD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "    													AND RD.PartNo = POD.PartNo    											  " & vbCrLf & _
                              "    													AND RD.PONo = POD.PONo " & vbCrLf & _
                              "                                                     AND RD.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                              "                                                     AND RD.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
                              "                 LEFT JOIN dbo.ReceivePASI_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo    " & vbCrLf & _
                              "    													AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                              "                                                     --AND isnull(RM.HT_CLS,0) = 0  " & vbCrLf & _
                              "  				LEFT JOIN dbo.Kanban_Detail KD ON KD.PONo = POM.PONo    " & vbCrLf & _
                              "  													 AND KD.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "  													 AND KD.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  													 AND KD.PartNo = POD.PartNo   " & vbCrLf

            ls_SQL = ls_SQL + "  													 AND KD.KanbanNo = DOD.KanbanNo  " & vbCrLf & _
                              "  				LEFT JOIN dbo.Kanban_Master KM ON KM.KanbanNo = KD.KanbanNo   " & vbCrLf & _
                              "  													 AND KM.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  													 AND KM.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  													 AND KM.KanbanNo = DOD.KanbanNo  				INNER JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                              "  				INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                              "  				INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "     WHERE ISNULL(RM.SuratJalanNo,'')<> '' " & vbCrLf & _
                              "  		 --DATA DO YG SUDAH DI RECEIVE " & vbCrLf
            ls_SQL = ls_SQL + " UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "  		SELECT DISTINCT    " & vbCrLf & _
                              "   				 coldetail = 'ReceivingEntry.aspx?prm='+CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(RM.entryDate,'')),106)   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(POM.SupplierID)                                                   + '|' +Rtrim(MS.SupplierName)   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(ISNULL(DOM.SuratJalanNo,''))   " & vbCrLf & _
                              "                                                  + '|' +Rtrim( CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106))   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),106))   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(ISNULL(KD.KanbanNo,''))   --+ '|' +Rtrim(CASE WHEN POD.KanbanCls = '0' THEN '-' ELSE ISNULL(KD.KanbanNo,'') END)  " & vbCrLf & _
                              "                                                  + '|' +Rtrim(POM.AffiliateID) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOD.PONo) ,  " & vbCrLf & _
                              "   				 coldetailname = CASE WHEN ISNULL(RD.GoodRecQty,0)=0 THEN 'RECEIVE' ELSE 'DETAIL' END,    " & vbCrLf & _
                              "  				 colno = '' ,   " & vbCrLf & _
                              "  				 colperiod = Right(Convert(char(11),Convert(datetime,POM.Period),106),8) ,   " & vbCrLf & _
                              "  				 colaffiliatecode = POM.AffiliateID ,    " & vbCrLf & _
                              "  				 colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
                              "  				 colpono = POM.PONo ,    				 colsuppliercode = POM.SupplierID ,    " & vbCrLf

            ls_SQL = ls_SQL + "  				 colsuppliername = MS.SupplierName ,    " & vbCrLf & _
                              "  				 colpokanban = CASE WHEN ISNULL(POD.KanbanCls,'0')='0' THEN 'NO' ELSE 'YES' END ,    " & vbCrLf & _
                              "  				 colkanbanno = '', --ISNULL(KD.KanbanNo,'') ,    " & vbCrLf & _
                              "  				 colplandeldate = '', --CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106) ,    " & vbCrLf & _
                              "  				 coldeldate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),106) ,    " & vbCrLf & _
                              "  				 colsj = ISNULL(DOM.SuratJalanNo,'') ,    " & vbCrLf & _
                              "  				 colpartno = '' , " & vbCrLf & _
                              "  				 colpartname = '' ,   " & vbCrLf & _
                              "  				 coluom = '' ,   " & vbCrLf & _
                              "  				 coldeliveryqty = '' ,   				 colreceiveqty = '' ,   " & vbCrLf & _
                              "                   coldefect = '' ,  " & vbCrLf

            ls_SQL = ls_SQL + "  				 colremaining = '' ,   " & vbCrLf & _
                              "  				 colreceivedate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(RM.entryDate,'')),106) ,   " & vbCrLf & _
                              "  				 colreceiveby = ISNULL(RM.EntryUser,''),  " & vbCrLf & _
                              "  				 pom.PONo H_POORDER, H_IDXORDER = 0, H_KANBANORDER = ISNULL(KD.KanbanNo,'-') ,   " & vbCrLf & _
                              "  				 H_AFFILIATEORDER = POM.AffiliateID , H_KANBANCLS = pod.KanbanCls,  " & vbCrLf & _
                              "  				 H_PLANDELDATE =  CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106), " & vbCrLf & _
                              "  				 H_ALREADYDELIVER = '', " & vbCrLf & _
                              "  				 H_REMAINING = CONVERT(CHAR,(SELECT SUM(ISNULL(A.DOQty,0)) " & vbCrLf & _
                              "                                                 From dbo.DOSupplier_Detail A " & vbCrLf & _
                              "                                                 WHERE(A.AffiliateID = DOM.AffiliateID) " & vbCrLf & _
                              "                                                 AND A.SupplierID = DOM.SupplierID " & vbCrLf & _
                              "                                                 AND A.KanbanNo = KM.KanbanNo " & vbCrLf & _
                              "                                                 AND A.PONo = POM.PONo) - " & vbCrLf & _
                              "                                             (SELECT ISNULL(SUM(ISNULL(A.GoodRecQty,0)) + SUM(ISNULL(A.DefectRecQty,0)),0) " & vbCrLf & _
                              "                                                 From dbo.ReceivePASI_Detail A  " & vbCrLf & _
                              "                                                 WHERE(A.AffiliateID =  DOM.AffiliateID) " & vbCrLf & _
                              "                                                 AND A.SupplierID = DOM.SupplierID " & vbCrLf & _
                              "                                                 AND A.KanbanNo = KM.KanbanNo " & vbCrLf & _
                              "                                                 AND A.PONo = POM.PONo )), " & vbCrLf & _
                              "  				 H_SUPPLIER = POM.SupplierID, " & vbCrLf & _
                              "  				 H_RECEIVEDATE = CONVERT(CHAR(10), CONVERT(DATETIME,ISNULL(RM.ReceiveDate,'')),120),  " & vbCrLf & _
                              "  				 H_SURATJALAN = ISNULL(DOM.SuratJalanNo,'') " & vbCrLf

            ls_SQL = ls_SQL + "  				 --,H_PARTCODE = pod.PartNo " & vbCrLf & _
                              "  				  " & vbCrLf & _
                              "  		 FROM    dbo.PO_Master POM   " & vbCrLf & _
                              "  				 INNER JOIN dbo.PO_Detail POD ON POD.PONo = POM.PONo    " & vbCrLf & _
                              "  												  AND POD.SupplierID = POM.SupplierID                                              " & vbCrLf & _
                              "  												  AND POD.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                              "  												  AND POM.DeliveryByPasiCls = 1    				INNER JOIN dbo.DOSupplier_Detail DOD ON DOD.PONo = POD.PONo    " & vbCrLf & _
                              "  													 AND DOD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "  													 AND DOD.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "  													 AND DOD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  				INNER JOIN dbo.DOSupplier_Master DOM ON DOM.SuratJalanNo =  DOD.SuratJalanNo    " & vbCrLf

            ls_SQL = ls_SQL + "  													 AND DOM.SupplierID = DOD.SupplierID    " & vbCrLf & _
                              "  													 AND DOM.AffiliateID = DOD.AffiliateID    " & vbCrLf & _
                              "  				LEFT JOIN dbo.ReceivePASI_Detail RD ON RD.SuratJalanNo = DOM.SuratJalanNo    " & vbCrLf & _
                              "    													AND RD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "    													AND RD.PartNo = POD.PartNo    											  " & vbCrLf & _
                              "    													AND RD.PONo = POD.PONo " & vbCrLf & _
                              "                                                     AND RD.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                              "                                                     AND RD.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
                              "                 LEFT JOIN dbo.ReceivePASI_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo    " & vbCrLf & _
                              "    													AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                              "                                                     --AND isnull(RM.HT_CLS,0) = 0  " & vbCrLf & _
                              "  				LEFT JOIN dbo.Kanban_Detail KD ON KD.PONo = POM.PONo    " & vbCrLf & _
                              "  													 AND KD.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "  													 AND KD.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  													 AND KD.PartNo = POD.PartNo   " & vbCrLf

            ls_SQL = ls_SQL + "  													 AND KD.KanbanNo = DOD.KanbanNo  " & vbCrLf & _
                              "  				LEFT JOIN dbo.Kanban_Master KM ON KM.KanbanNo = KD.KanbanNo   " & vbCrLf & _
                              "  													 AND KM.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  													 AND KM.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  													 AND KM.KanbanNo = DOD.KanbanNo  				INNER JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                              "  				INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                              "  				INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "     WHERE RTRIM(DOD.SuratJalanNo) + RTRIM(DOD.PONo)+ RTRIM(DOD.KanbanNo)+ RTRIM(DOD.PartNo) " & vbCrLf & _
                              "  	NOT IN (SELECT RTRIM(SuratJalanNo) + RTRIM(PONo) + RTRIM(KanbanNo) + RTRIM(DOD.PartNo) FROM dbo.ReceivePASI_Detail) " & vbCrLf & _
                              " -- DATA YG BELUM DI RECEIVE " & vbCrLf & _
                              "  )x	 where H_POORDER <> ''  --HEADER " & vbCrLf

            If checkbox1.Checked = True Then
                ls_SQL = ls_SQL + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(H_PLANDELDATE,'')),106) = '" & Format(dt1.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If rbdeliver.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_SURATJALAN, '') <> '' " & vbCrLf
            ElseIf rbdeliver.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_SURATJALAN,'') = '' " & vbCrLf
            End If

            If rbreceiving.Value = "YES" Then
                ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) > 0 " & vbCrLf
            ElseIf rbreceiving.Value = "NO" Then
                ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) = 0 " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_SQL = ls_SQL + " AND H_SURATJALAN LIKE '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_SQL = ls_SQL + " AND ISNULL(H_RECEIVEDATE,'') between '" & Format(dtfrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtto.Value, "yyyy-MM-dd") & "'" & vbCrLf
            End If

            If cbosupplier.Text <> clsGlobal.gs_All And cbosupplier.Text <> "" Then
                ls_SQL = ls_SQL + " AND H_SUPPLIER = '" & Trim(cbosupplier.Text) & "'" & vbCrLf
            End If

            If cbopart.Text <> clsGlobal.gs_All And cbopart.Text <> "" Then
                ls_SQL = ls_SQL + "AND H_PARTCODE = '" & Trim(cbopart.Text) & "'" & vbCrLf
            End If

            If rbkanban.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_KANBANCLS, '') = '1' " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_KANBANCLS,'') = '0' " & vbCrLf
            End If

            If rbgoodreceive.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_RECEIVEDATE, '') <> '1900-01-01'  " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_RECEIVEDATE,'') = '1900-01-01' " & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_SQL = ls_SQL + " AND h_poorder LIKE '%" & Trim(txtpono.Text) & "%'"
            End If


            ls_SQL = ls_SQL + "  	UNION ALL  " & vbCrLf & _
                              "  	  " & vbCrLf & _
                              "  	SELECT DISTINCT coldetail = ''  ,  " & vbCrLf & _
                              "  			 coldetailname = '',   " & vbCrLf & _
                              "  			 colno = '' ,   " & vbCrLf & _
                              "  			 colperiod = '' ,   " & vbCrLf & _
                              "  			 colaffiliatecode = '',  			 colaffiliatename = '' ,   " & vbCrLf & _
                              "  			 colpono = '' ,   " & vbCrLf & _
                              "  			 colsuppliercode = '' ,   " & vbCrLf & _
                              "  			 colsuppliername = '' ,   " & vbCrLf & _
                              "  			 colpokanban = '' ,   " & vbCrLf & _
                              "  			 colkanbanno = ISNULL(KD.KanbanNo,''),   " & vbCrLf & _
                              "  			 colplandeldate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106) ,   " & vbCrLf & _
                              "  			 coldeldate = '' ,   " & vbCrLf

            ls_SQL = ls_SQL + "  			 colsj = '' ,   " & vbCrLf & _
                              "  			 colpartno = pod.PartNo ,    " & vbCrLf & _
                              "  			 colpartname = MP.PartName ,    			 coluom = UC.Description ,    " & vbCrLf & _
                              "  			 coldeliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(DOD.DOQty,0)))) ,    " & vbCrLf & _
                              "  			 colreceiveqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(RD.GoodRecQty,0))))  ,    " & vbCrLf & _
                              "  			 coldefect = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(RD.DefectRecQty,0))))  ,    " & vbCrLf & _
                              "  			 colremaining = Convert(char,CONVERT(Numeric(9,0),(ISNULL(DOD.DOQty,0) - (ISNULL(RD.GoodRecQty,0) + ISNULL(RD.DefectRecQty,0) )))) ,    " & vbCrLf & _
                              "  			 colreceivedate = '' ,   " & vbCrLf & _
                              "  			 colreceiveby = '',   " & vbCrLf & _
                              "  			 POOrder = POM.PONo, idxorder = 1, kanbanorder = ISNULL(KD.KanbanNo,'-') , affiliateorder = POM.AffiliateID , pod.KanbanCls, " & vbCrLf & _
                              "  			 H_PLANDELDATE =  CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106), " & vbCrLf

            ls_SQL = ls_SQL + "  				 H_ALREADYDELIVER = '', " & vbCrLf & _
                              "  				 H_REMAINING = CONVERT(CHAR,(ISNULL(DOD.DOQty,0) - (ISNULL(RD.GoodRecQty,0) + ISNULL(RD.DefectRecQty,0) ))), " & vbCrLf & _
                              "  				 H_SUPPLIER = POM.SupplierID, " & vbCrLf & _
                              "  				 H_RECEIVEDATE = CONVERT(CHAR(10), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),120),  " & vbCrLf & _
                              "  				 H_SURATJALAN = ISNULL(DOM.SuratJalanNo,'') " & vbCrLf & _
                              "  				 --,H_PARTCODE = pod.PartNo  " & vbCrLf & _
                              "  	 FROM    dbo.PO_Master POM   " & vbCrLf & _
                              "  			 INNER JOIN dbo.PO_Detail POD ON POD.PONo = POM.PONo    " & vbCrLf & _
                              "  											  AND POD.SupplierID = POM.SupplierID                                              " & vbCrLf & _
                              "  											  AND POD.AffiliateID = POM.AffiliateID    											  AND POM.DeliveryByPasiCls = 1  " & vbCrLf & _
                              "  			LEFT JOIN dbo.Kanban_Detail KD ON KD.PONo = POM.PONo    " & vbCrLf

            ls_SQL = ls_SQL + "  												 AND KD.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "  												 AND KD.AffiliateID = POM.AffiliateID  " & vbCrLf & _
                              "  												 AND KD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  			LEFT JOIN dbo.Kanban_Master KM ON KM.KanbanNo = KD.KanbanNo   " & vbCrLf & _
                              "  												 AND KM.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  												 AND KM.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  			INNER JOIN dbo.DOSupplier_Detail DOD ON DOD.PONo = POD.PONo    " & vbCrLf & _
                              "  												 AND DOD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "  												 AND DOD.AffiliateID = POD.AffiliateID   												 AND DOD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  												 AND DOD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "  			INNER JOIN dbo.DOSupplier_Master DOM ON DOM.SuratJalanNo = DOD.SuratJalanNo    " & vbCrLf

            ls_SQL = ls_SQL + "  												 AND DOM.SupplierID = DOD.SupplierID    " & vbCrLf & _
                              "  												 AND DOM.AffiliateID = DOD.AffiliateID    " & vbCrLf & _
                              "  			LEFT JOIN dbo.ReceivePASI_Detail RD ON RD.SuratJalanNo = DOM.SuratJalanNo    " & vbCrLf & _
                              "    												AND RD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "    												AND RD.PartNo = DOD.PartNo    											  " & vbCrLf & _
                              "    												AND RD.PONo = POD.PONo    " & vbCrLf & _
                              "    												AND RD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "                                                 AND RD.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
                              "    			LEFT JOIN dbo.ReceivePASI_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              "                                                 AND RM.SupplierID = RD.SupplierID  	         " & vbCrLf & _
                              "                                                 --AND isnull(RM.HT_CLS,0) = 0  " & vbCrLf & _
                              "  			INNER JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                              "  			INNER JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
                              "  			INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf

            ls_SQL = ls_SQL + "  			INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  	--WHERE ISNULL(POD.KanbanCls,'0') <> '0'  " & vbCrLf & _
                              "    " & vbCrLf
            '                  "  	UNION ALL	  " & vbCrLf & _
            '                  "  	  " & vbCrLf & _
            '                  "  		SELECT DISTINCT coldetail = ''  ,  " & vbCrLf & _
            '                  "  			  coldetailname = '',   			 colno = '' ,   " & vbCrLf & _
            '                  "  			 colperiod = '' ,   " & vbCrLf & _
            '                  "  			 colaffiliatecode = '',   " & vbCrLf & _
            '                  "  			 colaffiliatename = '' ,   " & vbCrLf & _
            '                  "  			 colpono = '' ,   " & vbCrLf

            'ls_SQL = ls_SQL + "  			 colsuppliercode = '' ,   " & vbCrLf & _
            '                  "  			 colsuppliername = '' ,   " & vbCrLf & _
            '                  "  			 colpokanban = '' ,   " & vbCrLf & _
            '                  "  			 colkanbanno = '',  " & vbCrLf & _
            '                  "  			 colplandeldate = '' ,   " & vbCrLf & _
            '                  "  			 coldeldate = '' ,   			 colsj = '' ,   " & vbCrLf & _
            '                  "  			 colpartno = pod.PartNo ,    " & vbCrLf & _
            '                  "  			 colpartname = MP.PartName ,    " & vbCrLf & _
            '                  "  			 coluom = UC.Description ,    " & vbCrLf & _
            '                  "  			 coldeliveryqty = CONVERT(Char,CONVERT(Numeric(9,0),(ISNULL(DOD.DOQty,0)))) ,    " & vbCrLf & _
            '                  "  			 colreceiveqty = CONVERT(CHAR,CONVERT(NUMERIC(9,0),ISNULL(RD.GoodRecQty,0)))  ,    " & vbCrLf

            'ls_SQL = ls_SQL + "  			 coldefect = Convert(Char,CONVERT(Numeric(9,0),ISNULL(RD.DefectRecQty,0)))  ,    " & vbCrLf & _
            '                  "  			 colremaining = CONVERT(CHAR,CONVERT(Numeric(9,0),(ISNULL(DOD.DOQty,0) - (ISNULL(RD.GoodRecQty,0) + ISNULL(RD.DefectRecQty,0) )))) ,    " & vbCrLf & _
            '                  "  			 colreceivedate = '' ,   " & vbCrLf & _
            '                  "  			 colreceiveby = '',   " & vbCrLf & _
            '                  "  			 POOrder = POM.PONo, idxorder = 1, kanbanorder = '-' , affiliateorder = POM.AffiliateID , pod.KanbanCls, " & vbCrLf & _
            '                  "  			 H_PLANDELDATE =  '', " & vbCrLf & _
            '                  "  				 H_ALREADYDELIVER = '', " & vbCrLf & _
            '                  "  				 H_REMAINING = CONVERT(CHAR,(ISNULL(DOD.DOQty,0) - (ISNULL(RD.GoodRecQty,0) + ISNULL(RD.DefectRecQty,0) ))), " & vbCrLf & _
            '                  "  				 H_SUPPLIER = POM.SupplierID, " & vbCrLf & _
            '                  "  				 H_RECEIVEDATE = CONVERT(CHAR(10), CONVERT(DATETIME,ISNULL(RM.ReceiveDate,'')),120),  " & vbCrLf & _
            '                  "  				 H_SURATJALAN = ISNULL(DOM.SuratJalanNo,'') " & vbCrLf

            'ls_SQL = ls_SQL + "  				 --,H_PARTCODE = pod.PartNo  " & vbCrLf & _
            '                  "  	 FROM    dbo.PO_Master POM   			 INNER JOIN dbo.PO_Detail POD ON POD.PONo = POM.PONo    " & vbCrLf & _
            '                  "  											  AND POD.SupplierID = POM.SupplierID                                              " & vbCrLf & _
            '                  "  											  AND POD.AffiliateID = POM.AffiliateID    " & vbCrLf & _
            '                  "  											  AND POM.DeliveryByPasiCls = 1  " & vbCrLf & _
            '                  "  			INNER JOIN dbo.DOSupplier_Detail DOD ON DOD.PONo = POD.PONo    " & vbCrLf & _
            '                  "  												 AND DOD.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "  												 AND DOD.AffiliateID = POD.AffiliateID   " & vbCrLf & _
            '                  "  												 AND DOD.PartNo = POD.PartNo  " & vbCrLf & _
            '                  "  			INNER JOIN dbo.DOSupplier_Master DOM ON DOM.SuratJalanNo = DOD.SuratJalanNo    " & vbCrLf & _
            '                  "  												 AND DOM.SupplierID = DOD.SupplierID    " & vbCrLf

            'ls_SQL = ls_SQL + "  												 AND DOM.AffiliateID = DOD.AffiliateID    			LEFT JOIN dbo.ReceivePASI_Detail RD ON RD.SuratJalanNo = DOM.SuratJalanNo    " & vbCrLf & _
            '                  "    												AND RD.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "    												AND RD.PartNo = DOD.PartNo    											  " & vbCrLf & _
            '                  "    												AND RD.PONo = POD.PONo    " & vbCrLf & _
            '                  "    			LEFT JOIN dbo.ReceivePASI_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo    " & vbCrLf & _
            '                  "    												AND RM.SupplierID = RD.SupplierID  	         " & vbCrLf & _
            '                  "                                                 AND isnull(RM.HT_CLS,0) = 0  " & vbCrLf & _
            '                  "  			INNER JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "  			INNER JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
            '                  "  			INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
            '                  "  			INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
            '                  "  	--WHERE ISNULL(POD.KanbanCls,'0') = '0'   " & vbCrLf

            ls_SQL = ls_SQL + "  	)x  where H_POORDER <> ''  " & vbCrLf & _
                              "  "

            If checkbox1.Checked = True Then
                ls_SQL = ls_SQL + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(H_PLANDELDATE,'')),106) = '" & Format(dt1.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If rbdeliver.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_SURATJALAN, '') <> '' " & vbCrLf
            ElseIf rbdeliver.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_SURATJALAN,'') = '' " & vbCrLf
            End If

            If rbreceiving.Value = "YES" Then
                ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) > 0 " & vbCrLf
            ElseIf rbreceiving.Value = "NO" Then
                ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) = 0 " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_SQL = ls_SQL + " AND H_SURATJALAN LIKE '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_SQL = ls_SQL + " AND ISNULL(H_RECEIVEDATE,'') between '" & Format(dtfrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtto.Value, "yyyy-MM-dd") & "'" & vbCrLf
            End If

            If cbosupplier.Text <> clsGlobal.gs_All And cbosupplier.Text <> "" Then
                ls_SQL = ls_SQL + " AND H_SUPPLIER = '" & Trim(cbosupplier.Text) & "'" & vbCrLf
            End If

            If cbopart.Text <> clsGlobal.gs_All And cbopart.Text <> "" Then
                ls_SQL = ls_SQL + "AND H_PARTCODE = '" & Trim(cbopart.Text) & "'" & vbCrLf
            End If

            If rbkanban.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_KANBANCLS, '') = '1' " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_KANBANCLS,'') = '0' " & vbCrLf
            End If

            If rbgoodreceive.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_RECEIVEDATE, '') <> '1900-01-01'  " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_RECEIVEDATE,'') = '1900-01-01' " & vbCrLf
            End If


            If txtpono.Text <> "" Then
                ls_SQL = ls_SQL + " AND h_poorder LIKE '%" & Trim(txtpono.Text) & "%'"
            End If
            ls_SQL = ls_SQL + " ORDER BY h_poorder, h_KanbanCls,h_kanbanorder, H_SURATJALAN, h_idxorder "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)

                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "  SELECT * FROM (  " & vbCrLf & _
                     "  	SELECT distinct coldetail, coldetailname, CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY h_poorder, H_SURATJALAN,H_SJ,H_PLANDELDATE,h_idxorder)) AS colno,  " & vbCrLf & _
                     "  		colperiod, colaffiliatecode, colaffiliatename ,  " & vbCrLf & _
                     "  		colpono, colsuppliercode,colsuppliername, colpokanban, colkanbanno,  " & vbCrLf & _
                     "  		colplandeldate, coldeldate, colsj,colpartno,colpartname,  " & vbCrLf & _
                     "  		coluom,coldeliveryqty,colreceiveqty,coldefect,colremaining,colreceivedate= (CASE WHEN colreceivedate = '01 Jan 1900' THEN '' ELSE colreceivedate END) ,  " & vbCrLf & _
                     "  		colreceiveby,H_POORDER, H_IDXORDER ,H_KANBANORDER ,H_AFFILIATEORDER, H_KANBANCLS, " & vbCrLf & _
                     "  		H_PLANDELDATE,H_ALREADYDELIVER,H_REMAINING,H_SUPPLIER,H_RECEIVEDATE,H_SURATJALAN, H_SJ " & vbCrLf & _
                     "  	FROM (  " & vbCrLf & _
                     "  	SELECT distinct coldetail, coldetailname = CASE WHEN ISNULL(H_SJ,'')='' THEN 'RECEIVE' ELSE 'DETAIL' END, colno = '',  " & vbCrLf & _
                     "  		colperiod, colaffiliatecode, colaffiliatename ,  " & vbCrLf & _
                     "  		colpono, colsuppliercode,colsuppliername, colpokanban, colkanbanno,  " & vbCrLf & _
                     "  		colplandeldate, coldeldate, colsj,colpartno,colpartname,  " & vbCrLf & _
                     "  		coluom,coldeliveryqty,colreceiveqty,coldefect,colremaining,colreceivedate= (CASE WHEN colreceivedate = '01 Jan 1900' THEN '' ELSE colreceivedate END) ,  " & vbCrLf & _
                     "  		colreceiveby,H_POORDER, H_IDXORDER ,H_KANBANORDER ,H_AFFILIATEORDER, H_KANBANCLS, " & vbCrLf & _
                     "  		H_PLANDELDATE,H_ALREADYDELIVER,H_REMAINING,H_SUPPLIER,H_RECEIVEDATE,H_SURATJALAN, H_SJ " & vbCrLf & _
                     "  	FROM (  " & vbCrLf & _
                     "  		SELECT DISTINCT    " & vbCrLf & _
                     "   				 coldetail = 'ReceivingEntry.aspx?prm='+CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(RM.entryDate,'')),106)   " & vbCrLf

            ls_SQL = ls_SQL + "                                                  + '|' +Rtrim(POM.SupplierID) + '|' +Rtrim(MS.SupplierName)   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(REPLACE(ISNULL(DOM.SuratJalanNo,''),'&','DAN'))   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(REPLACE(ISNULL(RD.SuratJalanNo,''),'&','DAN'))   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),106))   " & vbCrLf & _
                              "                                                  + '|' +''   --+ '|' +Rtrim(CASE WHEN POD.KanbanCls = '0' THEN '-' ELSE ISNULL(KD.KanbanNo,'') END)  " & vbCrLf & _
                              "                                                  + '|' +Rtrim(POM.AffiliateID) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOD.PONO) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(RM.DriverName) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(RM.DriverContact) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(RM.NoPol) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(RM.JenisArmada) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(RM.TotalBox) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(KM.KanbanDate) , " & vbCrLf & _
                              "   				 coldetailname = CASE WHEN ISNULL(RD.SuratjalanNo,'')='' THEN 'RECEIVE' ELSE 'DETAIL' END,    " & vbCrLf & _
                              "  				 colno = '' ,   " & vbCrLf & _
                              "  				 colperiod = Right(Convert(char(11),Convert(datetime,POM.Period),106),8),   " & vbCrLf & _
                              "  				 colaffiliatecode = POM.AffiliateID ,    " & vbCrLf & _
                              "  				 colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
                              "  				 colpono = POM.PONo ,    				 colsuppliercode = POM.SupplierID ,    " & vbCrLf

            ls_SQL = ls_SQL + "  				 colsuppliername = MS.SupplierName ,    " & vbCrLf & _
                              "  				 colpokanban = CASE WHEN ISNULL(POD.KanbanCls,'0')='0' THEN 'NO' ELSE 'YES' END ,    " & vbCrLf & _
                              "  				 colkanbanno = '', --ISNULL(KD.KanbanNo,'') ,    " & vbCrLf & _
                              "  				 colplandeldate = '', --CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106) ,    " & vbCrLf & _
                              "  				 coldeldate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),106) ,    " & vbCrLf & _
                              "  				 colsj = ISNULL(DOM.SuratJalanNo,'') ,    " & vbCrLf & _
                              "  				 colpartno = '' , " & vbCrLf & _
                              "  				 colpartname = '' ,   " & vbCrLf & _
                              "  				 coluom = '' ,   " & vbCrLf & _
                              "  				 coldeliveryqty = '' ,   				 colreceiveqty = '' ,   " & vbCrLf & _
                              "                   coldefect = '' ,  " & vbCrLf

            ls_SQL = ls_SQL + "  				 colremaining = '' ,   " & vbCrLf & _
                              "  				 colreceivedate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(RM.entryDate,'')),106) ,   " & vbCrLf & _
                              "  				 colreceiveby = ISNULL(RM.EntryUser,''),  " & vbCrLf & _
                              "  				 pom.PONo H_POORDER, H_IDXORDER = 0, H_KANBANORDER = '' ,   " & vbCrLf & _
                              "  				 H_AFFILIATEORDER = POM.AffiliateID , H_KANBANCLS = pod.KanbanCls,  " & vbCrLf & _
                              "  				 H_PLANDELDATE =  CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106), " & vbCrLf & _
                              "  				 H_ALREADYDELIVER = '', " & vbCrLf & _
                              "  				 H_REMAINING = CASE WHEN (SELECT SUM(ISNULL(A.DOQty,0)) " & vbCrLf & _
                              "                                                 From dbo.DOSupplier_Detail A " & vbCrLf & _
                              "                                                 WHERE(A.AffiliateID = DOM.AffiliateID) " & vbCrLf & _
                              "                                                 AND A.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
                              "                                                 AND A.SupplierID = DOM.SupplierID " & vbCrLf & _
                              "                                                 AND A.KanbanNo = KM.KanbanNo " & vbCrLf & _
                              "                                                 AND A.PONo = POM.PONo) - " & vbCrLf & _
                              "                                             (SELECT ISNULL(SUM(ISNULL(A.GoodRecQty,0)) + SUM(ISNULL(A.DefectRecQty,0)),0) " & vbCrLf & _
                              "                                                 From dbo.ReceivePASI_Detail A  " & vbCrLf & _
                              "                                                 WHERE(A.AffiliateID =  DOM.AffiliateID) " & vbCrLf & _
                              "                                                 AND A.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
                              "                                                 AND A.SupplierID = DOM.SupplierID " & vbCrLf & _
                              "                                                 AND A.KanbanNo = KM.KanbanNo " & vbCrLf & _
                              "                                                 AND A.PONo = POM.PONo ) > 0 THEN '1' ELSE '0' END, " & vbCrLf & _
                              "  				 H_SUPPLIER = POM.SupplierID, " & vbCrLf & _
                              "  				 H_RECEIVEDATE = CONVERT(CHAR(10), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),120),  " & vbCrLf & _
                              "  				 H_SURATJALAN = ISNULL(DOM.SuratJalanNo,'') , H_SJ = ISNULL(RM.SURATJALANNO,'') " & vbCrLf

            ls_SQL = ls_SQL + "  				 --,H_PARTCODE = pod.PartNo " & vbCrLf & _
                              "  				  " & vbCrLf & _
                              "  		 FROM    dbo.PO_Master POM   " & vbCrLf & _
                              "  				 INNER JOIN dbo.PO_Detail POD ON POD.PONo = POM.PONo    " & vbCrLf & _
                              "  												  AND POD.SupplierID = POM.SupplierID                                              " & vbCrLf & _
                              "  												  AND POD.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                              "  												  AND POM.DeliveryByPasiCls = 1    				INNER JOIN dbo.DOSupplier_Detail DOD ON DOD.PONo = POD.PONo    " & vbCrLf & _
                              "  													 AND DOD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "  													 AND DOD.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "  													 AND DOD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  				INNER JOIN dbo.DOSupplier_Master DOM ON DOM.SuratJalanNo =  DOD.SuratJalanNo    " & vbCrLf

            ls_SQL = ls_SQL + "  													 AND DOM.SupplierID = DOD.SupplierID    " & vbCrLf & _
                              "  													 AND DOM.AffiliateID = DOD.AffiliateID    " & vbCrLf & _
                              "  				LEFT JOIN dbo.ReceivePASI_Detail RD ON RD.SuratJalanNo = DOM.SuratJalanNo    " & vbCrLf & _
                              "    													AND RD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "    													AND RD.PartNo = POD.PartNo    											  " & vbCrLf & _
                              "    													AND RD.PONo = POD.PONo " & vbCrLf & _
                              "                                                     AND RD.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                              "                                                     AND RD.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
                              "                 LEFT JOIN dbo.ReceivePASI_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo    " & vbCrLf & _
                              "    													AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                              "                                                     --AND isnull(RM.HT_CLS,0) = 0  " & vbCrLf & _
                              "  				LEFT JOIN dbo.Kanban_Detail KD ON KD.PONo = POM.PONo    " & vbCrLf & _
                              "  													 AND KD.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "  													 AND KD.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  													 AND KD.PartNo = POD.PartNo   " & vbCrLf

            ls_SQL = ls_SQL + "  													 AND KD.KanbanNo = DOD.KanbanNo  " & vbCrLf & _
                              "  				LEFT JOIN dbo.Kanban_Master KM ON KM.KanbanNo = KD.KanbanNo   " & vbCrLf & _
                              "  													 AND KM.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  													 AND KM.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  													 AND KM.KanbanNo = DOD.KanbanNo  				INNER JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                              "  				INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                              "  				INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "     WHERE ISNULL(RM.SuratJalanNo,'')<> '' " & vbCrLf & _
                              "  		 --DATA DO YG SUDAH DI RECEIVE " & vbCrLf
            ls_SQL = ls_SQL + " UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "  		SELECT DISTINCT    " & vbCrLf & _
                              "   				 coldetail = 'ReceivingEntry.aspx?prm='+CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(RM.entryDate,'')),106) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(POM.SupplierID)                                                   + '|' +Rtrim(MS.SupplierName)   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(REPLACE(ISNULL(DOM.SuratJalanNo,''),'&','DAN'))   " & vbCrLf & _
                              "                                                  + '|' +''   " & vbCrLf & _
                              "                                                  + '|' +Rtrim(CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),106))   " & vbCrLf & _
                              "                                                  + '|' +''   --+ '|' +Rtrim(CASE WHEN POD.KanbanCls = '0' THEN '-' ELSE ISNULL(KD.KanbanNo,'') END)  " & vbCrLf & _
                              "                                                  + '|' +Rtrim(POM.AffiliateID) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOD.PONo) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOM.DriverName) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOM.DriverContact) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOM.NoPol) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOM.JenisArmada) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(DOM.TotalBox) " & vbCrLf & _
                              "                                                  + '|' +Rtrim(KM.KanbanDate) , " & vbCrLf & _
                              "   				 coldetailname = CASE WHEN ISNULL(RD.SuratJalanNo,'')='' THEN 'RECEIVE' ELSE 'DETAIL' END,    " & vbCrLf & _
                              "  				 colno = '' ,   " & vbCrLf & _
                              "  				 colperiod = Right(Convert(char(11),Convert(datetime,POM.Period),106),8) ,   " & vbCrLf & _
                              "  				 colaffiliatecode = POM.AffiliateID ,    " & vbCrLf & _
                              "  				 colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
                              "  				 colpono = POM.PONo , colsuppliercode = POM.SupplierID ,    " & vbCrLf

            ls_SQL = ls_SQL + "  				 colsuppliername = MS.SupplierName ,    " & vbCrLf & _
                              "  				 colpokanban = CASE WHEN ISNULL(POD.KanbanCls,'0')='0' THEN 'NO' ELSE 'YES' END ,    " & vbCrLf & _
                              "  				 colkanbanno = '', --ISNULL(KD.KanbanNo,'') ,    " & vbCrLf & _
                              "  				 colplandeldate = '', --CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106) ,    " & vbCrLf & _
                              "  				 coldeldate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),106) ,    " & vbCrLf & _
                              "  				 colsj = ISNULL(DOM.SuratJalanNo,'') ,    " & vbCrLf & _
                              "  				 colpartno = '' , " & vbCrLf & _
                              "  				 colpartname = '' ,   " & vbCrLf & _
                              "  				 coluom = '' ,   " & vbCrLf & _
                              "  				 coldeliveryqty = '' ,   				 colreceiveqty = '' ,   " & vbCrLf & _
                              "                   coldefect = '' ,  " & vbCrLf

            ls_SQL = ls_SQL + "  				 colremaining = '' ,   " & vbCrLf & _
                              "  				 colreceivedate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(RM.entryDate,'')),106) ,   " & vbCrLf & _
                              "  				 colreceiveby = ISNULL(RM.EntryUser,''),  " & vbCrLf & _
                              "  				 pom.PONo H_POORDER, H_IDXORDER = 0, H_KANBANORDER = '' ,   " & vbCrLf & _
                              "  				 H_AFFILIATEORDER = POM.AffiliateID , H_KANBANCLS = pod.KanbanCls,  " & vbCrLf & _
                              "  				 H_PLANDELDATE =  CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106), " & vbCrLf & _
                              "  				 H_ALREADYDELIVER = '', " & vbCrLf & _
                              "  				 H_REMAINING = '1', " & vbCrLf & _
                              "  				 H_SUPPLIER = POM.SupplierID, " & vbCrLf & _
                              "  				 H_RECEIVEDATE = CONVERT(CHAR(10), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),120),  " & vbCrLf & _
                              "  				 H_SURATJALAN = ISNULL(DOM.SuratJalanNo,''),H_SJ = ISNULL(RM.SURATJALANNO,'')  " & vbCrLf

            ls_SQL = ls_SQL + "  				 --,H_PARTCODE = pod.PartNo " & vbCrLf & _
                              "  				  " & vbCrLf & _
                              "  		 FROM    dbo.PO_Master POM   " & vbCrLf & _
                              "  				 INNER JOIN dbo.PO_Detail POD ON POD.PONo = POM.PONo    " & vbCrLf & _
                              "  												  AND POD.SupplierID = POM.SupplierID                                              " & vbCrLf & _
                              "  												  AND POD.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                              "  												  AND POM.DeliveryByPasiCls = 1    				INNER JOIN dbo.DOSupplier_Detail DOD ON DOD.PONo = POD.PONo    " & vbCrLf & _
                              "  													 AND DOD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "  													 AND DOD.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "  													 AND DOD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  				INNER JOIN dbo.DOSupplier_Master DOM ON DOM.SuratJalanNo =  DOD.SuratJalanNo    " & vbCrLf

            ls_SQL = ls_SQL + "  													 AND DOM.SupplierID = DOD.SupplierID    " & vbCrLf & _
                              "  													 AND DOM.AffiliateID = DOD.AffiliateID    " & vbCrLf & _
                              "  				LEFT JOIN dbo.ReceivePASI_Detail RD ON RD.SuratJalanNo = DOM.SuratJalanNo    " & vbCrLf & _
                              "    													AND RD.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "    													AND RD.PartNo = POD.PartNo    											  " & vbCrLf & _
                              "    													AND RD.PONo = POD.PONo " & vbCrLf & _
                              "                                                     AND RD.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                              "                                                     AND RD.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
                              "                 LEFT JOIN dbo.ReceivePASI_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo    " & vbCrLf & _
                              "    													AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                              "                                                     --AND isnull(RM.HT_CLS,0) = 0  " & vbCrLf & _
                              "  				LEFT JOIN dbo.Kanban_Detail KD ON KD.PONo = POM.PONo    " & vbCrLf & _
                              "  													 AND KD.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "  													 AND KD.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  													 AND KD.PartNo = POD.PartNo   " & vbCrLf

            ls_SQL = ls_SQL + "  													 AND KD.KanbanNo = DOD.KanbanNo  " & vbCrLf & _
                              "  				LEFT JOIN dbo.Kanban_Master KM ON KM.KanbanNo = KD.KanbanNo   " & vbCrLf & _
                              "  													 AND KM.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  													 AND KM.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "  													 AND KM.KanbanNo = DOD.KanbanNo  				INNER JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                              "  				INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
                              "  				INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf
            ls_SQL = ls_SQL + "     WHERE --RTRIM(DOD.SuratJalanNo) + RTRIM(DOD.PONo)+ RTRIM(DOD.KanbanNo)+ RTRIM(DOD.PartNo) " & vbCrLf & _
                              "  	--NOT IN (SELECT RTRIM(SuratJalanNo) + RTRIM(PONo) + RTRIM(KanbanNo) + RTRIM(DOD.PartNo) FROM dbo.ReceivePASI_Detail) " & vbCrLf & _
                              "     --AND " & vbCrLf & _
                              "     ISNULL(RM.SuratJalanNo,'') =  '' " & vbCrLf & _
                              " -- DATA YG BELUM DI RECEIVE " & vbCrLf
            ls_SQL = ls_SQL + "  )x	)xx  where H_POORDER <> ''  --HEADER " & vbCrLf

            If checkbox1.Checked = True Then
                ls_SQL = ls_SQL + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(H_PLANDELDATE,'')),106) = '" & Format(dt1.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If rbdeliver.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_SURATJALAN, '') <> '' " & vbCrLf
            ElseIf rbdeliver.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_SURATJALAN,'') = '' " & vbCrLf
            End If

            If rbreceiving.Value = "YES" Then
                ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) > 0 " & vbCrLf
            ElseIf rbreceiving.Value = "NO" Then
                ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) = 0 " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_SQL = ls_SQL + " AND H_SURATJALAN LIKE '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_SQL = ls_SQL + " AND ISNULL(H_RECEIVEDATE,'') between '" & Format(dtfrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtto.Value, "yyyy-MM-dd") & "'" & vbCrLf
            End If

            If cbosupplier.Text <> clsGlobal.gs_All And cbosupplier.Text <> "" Then
                ls_SQL = ls_SQL + " AND H_SUPPLIER = '" & Trim(cbosupplier.Text) & "'" & vbCrLf
            End If

            If cbopart.Text <> clsGlobal.gs_All And cbopart.Text <> "" Then
                ls_SQL = ls_SQL + "AND H_PARTCODE = '" & Trim(cbopart.Text) & "'" & vbCrLf
            End If

            If rbkanban.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_KANBANCLS, '') = '1' " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_KANBANCLS,'') = '0' " & vbCrLf
            End If

            If rbgoodreceive.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_RECEIVEDATE, '') <> '1900-01-01'  " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_RECEIVEDATE,'') = '1900-01-01' " & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_SQL = ls_SQL + " AND h_poorder LIKE '%" & Trim(txtpono.Text) & "%'"
            End If


            'ls_SQL = ls_SQL + "  	UNION ALL  " & vbCrLf & _
            '                  "  	  " & vbCrLf & _
            '                  "  	SELECT DISTINCT coldetail = ''  ,  " & vbCrLf & _
            '                  "  			 coldetailname = '',   " & vbCrLf & _
            '                  "  			 colno = '' ,   " & vbCrLf & _
            '                  "  			 colperiod = '' ,   " & vbCrLf & _
            '                  "  			 colaffiliatecode = '',  			 colaffiliatename = '' ,   " & vbCrLf & _
            '                  "  			 colpono = '' ,   " & vbCrLf & _
            '                  "  			 colsuppliercode = '' ,   " & vbCrLf & _
            '                  "  			 colsuppliername = '' ,   " & vbCrLf & _
            '                  "  			 colpokanban = '' ,   " & vbCrLf & _
            '                  "  			 colkanbanno = ISNULL(KD.KanbanNo,''),   " & vbCrLf & _
            '                  "  			 colplandeldate = CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106) ,   " & vbCrLf & _
            '                  "  			 coldeldate = '' ,   " & vbCrLf

            'ls_SQL = ls_SQL + "  			 colsj = '' ,   " & vbCrLf & _
            '                  "  			 colpartno = pod.PartNo ,    " & vbCrLf & _
            '                  "  			 colpartname = MP.PartName ,    			 coluom = UC.Description ,    " & vbCrLf & _
            '                  "  			 coldeliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(DOD.DOQty,0)))) ,    " & vbCrLf & _
            '                  "  			 colreceiveqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(RD.GoodRecQty,0))))  ,    " & vbCrLf & _
            '                  "  			 coldefect = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(RD.DefectRecQty,0))))  ,    " & vbCrLf & _
            '                  "  			 colremaining = Convert(char,CONVERT(Numeric(9,0),(ISNULL(DOD.DOQty,0) - (ISNULL(RD.GoodRecQty,0) + ISNULL(RD.DefectRecQty,0) )))) ,    " & vbCrLf & _
            '                  "  			 colreceivedate = '' ,   " & vbCrLf & _
            '                  "  			 colreceiveby = '',   " & vbCrLf & _
            '                  "  			 POOrder = POM.PONo, idxorder = 1, kanbanorder = ISNULL(KD.KanbanNo,'-') , affiliateorder = POM.AffiliateID , pod.KanbanCls, " & vbCrLf & _
            '                  "  			 H_PLANDELDATE =  CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106), " & vbCrLf

            'ls_SQL = ls_SQL + "  				 H_ALREADYDELIVER = '', " & vbCrLf & _
            '                  "  				 H_REMAINING = CASE WHEN (ISNULL(DOD.DOQty,0) - (ISNULL(RD.GoodRecQty,0) + ISNULL(RD.DefectRecQty,0) )) > 0 THEN '1' ELSE '0' END , " & vbCrLf & _
            '                  "  				 H_SUPPLIER = POM.SupplierID, " & vbCrLf & _
            '                  "  				 H_RECEIVEDATE = CONVERT(CHAR(10), CONVERT(DATETIME,ISNULL(DOM.DeliveryDate,'')),120),  " & vbCrLf & _
            '                  "  				 H_SURATJALAN = ISNULL(DOM.SuratJalanNo,''), H_SJ = ISNULL(RM.SURATJALANNO,'')  " & vbCrLf & _
            '                  "  				 --,H_PARTCODE = pod.PartNo  " & vbCrLf & _
            '                  "  	 FROM    dbo.PO_Master POM   " & vbCrLf & _
            '                  "  			 INNER JOIN dbo.PO_Detail POD ON POD.PONo = POM.PONo    " & vbCrLf & _
            '                  "  											  AND POD.SupplierID = POM.SupplierID                                              " & vbCrLf & _
            '                  "  											  AND POD.AffiliateID = POM.AffiliateID    											  AND POM.DeliveryByPasiCls = 1  " & vbCrLf & _
            '                  "  			LEFT JOIN dbo.Kanban_Detail KD ON KD.PONo = POM.PONo    " & vbCrLf

            'ls_SQL = ls_SQL + "  												 AND KD.SupplierID = POM.SupplierID    " & vbCrLf & _
            '                  "  												 AND KD.AffiliateID = POM.AffiliateID  " & vbCrLf & _
            '                  "  												 AND KD.PartNo = POD.PartNo  " & vbCrLf & _
            '                  "  			LEFT JOIN dbo.Kanban_Master KM ON KM.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "  												 AND KM.SupplierID = POM.SupplierID   " & vbCrLf & _
            '                  "  												 AND KM.AffiliateID = POM.AffiliateID   " & vbCrLf & _
            '                  "  			INNER JOIN dbo.DOSupplier_Detail DOD ON DOD.PONo = POD.PONo    " & vbCrLf & _
            '                  "  												 AND DOD.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "  												 AND DOD.AffiliateID = POD.AffiliateID   												 AND DOD.PartNo = POD.PartNo  " & vbCrLf & _
            '                  "  												 AND DOD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
            '                  "  			INNER JOIN dbo.DOSupplier_Master DOM ON DOM.SuratJalanNo = DOD.SuratJalanNo    " & vbCrLf

            'ls_SQL = ls_SQL + "  												 AND DOM.SupplierID = DOD.SupplierID    " & vbCrLf & _
            '                  "  												 AND DOM.AffiliateID = DOD.AffiliateID    " & vbCrLf & _
            '                  "  			LEFT JOIN dbo.ReceivePASI_Detail RD ON RD.SuratJalanNo = DOM.SuratJalanNo    " & vbCrLf & _
            '                  "    												AND RD.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "    												AND RD.PartNo = DOD.PartNo    											  " & vbCrLf & _
            '                  "    												AND RD.PONo = POD.PONo    " & vbCrLf & _
            '                  "    												AND RD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
            '                  "                                                 AND RD.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
            '                  "    			LEFT JOIN dbo.ReceivePASI_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
            '                  "                                                 AND RM.SupplierID = RD.SupplierID  	         " & vbCrLf & _
            '                  "                                                 --AND isnull(RM.HT_CLS,0) = 0  " & vbCrLf & _
            '                  "  			INNER JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "  			LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
            '                  "  			INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf

            'ls_SQL = ls_SQL + "  			INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
            '                  "  	--WHERE ISNULL(POD.KanbanCls,'0') <> '0'  " & vbCrLf & _
            '                  "    " & vbCrLf

            ls_SQL = ls_SQL + "  	)x  where H_POORDER <> ''  " & vbCrLf & _
                              "  "

            If checkbox1.Checked = True Then
                ls_SQL = ls_SQL + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(H_PLANDELDATE,'')),106) = '" & Format(dt1.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If rbdeliver.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_SURATJALAN, '') <> '' " & vbCrLf
            ElseIf rbdeliver.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_SURATJALAN,'') = '' " & vbCrLf
            End If

            If rbreceiving.Value = "YES" Then
                ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) > 0 " & vbCrLf
            ElseIf rbreceiving.Value = "NO" Then
                ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) = 0 " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_SQL = ls_SQL + " AND H_SURATJALAN LIKE '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_SQL = ls_SQL + " AND ISNULL(H_RECEIVEDATE,'') between '" & Format(dtfrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtto.Value, "yyyy-MM-dd") & "'" & vbCrLf
            End If

            If cbosupplier.Text <> clsGlobal.gs_All And cbosupplier.Text <> "" Then
                ls_SQL = ls_SQL + " AND H_SUPPLIER = '" & Trim(cbosupplier.Text) & "'" & vbCrLf
            End If

            If cbopart.Text <> clsGlobal.gs_All And cbopart.Text <> "" Then
                ls_SQL = ls_SQL + "AND H_PARTCODE = '" & Trim(cbopart.Text) & "'" & vbCrLf
            End If

            If rbkanban.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_KANBANCLS, '') = '1' " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_KANBANCLS,'') = '0' " & vbCrLf
            End If

            If rbgoodreceive.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(H_RECEIVEDATE, '') <> '1900-01-01'  " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(H_RECEIVEDATE,'') = '1900-01-01' " & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_SQL = ls_SQL + " AND h_poorder LIKE '%" & Trim(txtpono.Text) & "%'"
            End If
            ls_SQL = ls_SQL + " ORDER BY h_poorder, H_SURATJALAN, H_SJ, H_PLANDELDATE, h_idxorder "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)

                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub
#End Region

#Region "FORM EVENT"

    Private Sub grid_BatchUpdate(sender As Object, e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate

    End Sub
    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Dim pPlan As Date = Split(e.Parameters, "|")(1)
            Dim pSupplierDeliver As String = Split(e.Parameters, "|")(2)
            Dim pRemaining As String = Split(e.Parameters, "|")(3)
            Dim psj As String = Split(e.Parameters, "|")(4)
            Dim pDateFrom As Date = Split(e.Parameters, "|")(5)
            Dim pDateTo As Date = Split(e.Parameters, "|")(6)
            Dim pSupplier As String = Split(e.Parameters, "|")(7)
            Dim pPart As String = Split(e.Parameters, "|")(8)
            Dim pPoNo As String = Split(e.Parameters, "|")(9)
            Dim pKanban As String = Split(e.Parameters, "|")(10)

            Select Case pAction
                Case "gridload"
                    'Call up_GridLoad(pPlan, pSupplierDeliver, pRemaining, psj, pDateFrom, pDateTo, pSupplier, pPart, pPoNo, pKanban)
                    Call up_GridLoad()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "kosong"

            End Select

EndProcedure:
            Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

#End Region


#Region "MergeGrid"


    Private Sub MergeCells(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer, ByVal cell As TableCell)
        Dim isNextTheSame As Boolean = IsNextRowHasSameData(column, visibleIndex)
        If isNextTheSame Then
            If (Not mergedCells.ContainsKey(column)) Then
                mergedCells(column) = cell
            End If
        End If
        If IsPrevRowHasSameData(column, visibleIndex) Then
            CType(cell.Parent, TableRow).Cells.Remove(cell)
            If mergedCells.ContainsKey(column) Then
                Dim mergedCell As TableCell = mergedCells(column)
                If (Not cellRowSpans.ContainsKey(mergedCell)) Then
                    cellRowSpans(mergedCell) = 1
                End If
                cellRowSpans(mergedCell) = cellRowSpans(mergedCell) + 1
            End If
        End If
        If (Not isNextTheSame) Then
            mergedCells.Remove(column)
        End If
    End Sub
    Private Function IsNextRowHasSameData(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer) As Boolean
        'is it the last visible row
        If visibleIndex >= grid.VisibleRowCount - 1 Then
            Return False
        End If

        Return IsSameData(column.FieldName, visibleIndex, visibleIndex + 1)
    End Function
    Private Function IsPrevRowHasSameData(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer) As Boolean
        Dim grid As ASPxGridView = column.Grid
        'is it the first visible row
        If visibleIndex <= Me.grid.VisibleStartIndex Then
            Return False
        End If

        Return IsSameData(column.FieldName, visibleIndex, visibleIndex - 1)
    End Function
    Private Function IsSameData(ByVal fieldName As String, ByVal visibleIndex1 As Integer, ByVal visibleIndex2 As Integer) As Boolean
        ' is it a group row?
        If grid.GetRowLevel(visibleIndex2) <> grid.GroupCount Then
            Return False
        End If

        Return Object.Equals(grid.GetRowValues(visibleIndex1, fieldName), grid.GetRowValues(visibleIndex2, fieldName))
    End Function

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        'add the attribute that will be used to find which column the cell belongs to
        e.Cell.Attributes.Add("colaffiliatecode", e.DataColumn.VisibleIndex.ToString())

        If cellRowSpans.ContainsKey(e.Cell) Then
            e.Cell.RowSpan = cellRowSpans(e.Cell)
        End If

    End Sub

    Private Sub grid_HtmlRowCreated(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowCreated
        'If grid.GetRowLevel(e.VisibleIndex) <> grid.GroupCount Then
        '    Return
        'End If

        'For i As Integer = e.Row.Cells.Count - 1 To 0 Step -1
        '    Dim dataCell As DevExpress.Web.ASPxGridView.Rendering.GridViewTableDataCell = TryCast(e.Row.Cells(i), DevExpress.Web.ASPxGridView.Rendering.GridViewTableDataCell)
        '    If dataCell IsNot Nothing Then
        '        If dataCell.DataColumn.Name <> "colpartno" Then
        '            MergeCells(dataCell.DataColumn, e.VisibleIndex, dataCell)
        '        End If
        '    End If
        'Next i

    End Sub
#End Region


    'Private Sub GridView1_CellMerge(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.CellMergeEventArgs) Handles GridView1.CellMerge
    '    'edit merge disini
    '    Select Case e.Column.FieldName 'cek nama kolom 
    '        Case "TitleOfCourtesy"
    '            'logika, jika ountry DAN city DAN TitleOfCourtesy maka baru kita merge
    '            'jika tidak maka jangan merge
    '            If GridView1.GetRowCellValue(e.RowHandle1, "Country") = GridView1.GetRowCellValue(e.RowHandle2, "Country") And _
    '                GridView1.GetRowCellValue(e.RowHandle1, "City") = GridView1.GetRowCellValue(e.RowHandle2, "City") And _
    '                GridView1.GetRowCellValue(e.RowHandle1, "TitleOfCourtesy") = GridView1.GetRowCellValue(e.RowHandle2, "TitleOfCourtesy") Then
    '                e.Merge = True
    '            Else
    '                e.Merge = False
    '            End If
    '            e.Handled = True 'untuk men-apply setting merge
    '    End Select

    'End Sub

    Private Sub grid_HtmlRowPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowPrepared
        If e.RowType <> GridViewRowType.Data Then Return
        If e.GetValue("colpartno").ToString = "" Then e.Row.BackColor = Drawing.Color.LightGray
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
End Class