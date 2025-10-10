Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO

Public Class ReceivingEntry
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = ""

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "E02"

    'Parsing parameter from Supplier/PASI Delivery Confirmation
    Dim pm_ReceivedDate As String = "", _
        pm_SupplierCode As String = "", pm_SupplierName As String = "", _
        pm_SupplierSJNo As String = "", pm_SupplierPlanDeliveryDate As String = "", pm_SupplierDeliveryDate As String = "", _
        pm_PASISJNo As String = "", pm_PASIDeliveryDate As String = "", _
        pm_DeliveryLocationCode As String = "", pm_DeliveryLocationName As String = "", _
        pm_DriverName As String = "", pm_DriverContact As String = "", pm_NoPol As String = "", pm_JenisArmada As String = "", pm_TotalBox As String = ""
    Dim pm_PONo As String = "", pm_KanbanNo As String = ""
    Dim pm_DeliveryBypasi As String

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim dtHeader As DataTable
    Dim dtDetail As DataTable
#End Region

#Region "Procedures"
    Private Sub up_Initialize()
        Dim script As String = _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(lblInfo, lblInfo.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        'Combo Performance
        With cboPerformanceCls
            ls_SQL = "SELECT PerformanceCls, Description FROM dbo.MS_PerformanceCls" & vbCrLf
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PerformanceCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 140

                .TextField = "PerformanceCls"
                .DataBind()
            End Using
        End With
    End Sub

    Private Sub UpdateExcel(ByVal pIsNewData As Boolean, _
                        Optional ByVal pAffCode As String = "", _
                        Optional ByVal pSuratJalan As String = "", _
                        Optional ByVal pSuppCode As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " UPDATE dbo.ReceiveAffiliate_Master " & vbCrLf & _
                          " SET ExcelCls='1'" & vbCrLf & _
                          " WHERE SuratJalanNo='" & pSuratJalan & "'  " & vbCrLf & _
                          " AND AffiliateID='" & pAffCode & "' " 

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""

            ls_SQL = "  SELECT ColNo = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PONo, KanbanCls, KanbanNo)),  " & vbCrLf & _
                  "  	   Supplier, PONo, POKanban = CASE WHEN ISNULL(KanbanCls,'0') = '1' THEN 'YES' ELSE 'NO' END,   " & vbCrLf & _
                  "  	   KanbanNo, PartNo, PartName, UOM, QtyBox,  " & vbCrLf & _
                  "  	   SupplierDeliveryQty = SUM(SupplierDeliveryQty), PASIGoodReceivingQty = SUM(PASIGoodReceivingQty),  " & vbCrLf & _
                  "  	   PASIDefectQty = SUM(PASIDefectQty), PASIDeliveryQty = SUM(PASIDeliveryQty), GoodReceivingQty = SUM(GoodReceivingQty),  " & vbCrLf & _
                  "         DefectReceivingQty = SUM(DefectReceivingQty), RemainingReceivingQty = SUM(RemainingReceivingQty),  " & vbCrLf & _
                  "         ReceivingQtyBox = SUM(ReceivingQtyBox), UnitCls, DeliveryByPASICls, IsSaved,sjpasi  " & vbCrLf & _
                  "    FROM (  " & vbCrLf & _
                  "  		  SELECT DISTINCT PONo = POD.PONo, Supplier = isnull(DPD.SupplierID,''),  " & vbCrLf & _
                  "  				 KanbanCls = POD.KanbanCls,  " & vbCrLf & _
                  "  				 KanbanNo = KD.KanbanNo,  "

            ls_SQL = ls_SQL + "  				 PartNo = POD.PartNo,  " & vbCrLf & _
                              "  				 PartName = MP.PartName,  " & vbCrLf & _
                              "  				 UOM = MU.Description,  " & vbCrLf & _
                              "  				 QtyBox = ISNULL(POD.POQtyBox,MPM.QtyBox),  " & vbCrLf & _
                              "  				 SupplierDeliveryQty = CONVERT(NUMERIC(18,0),ISNULL(DSD.DOQty,'0')),  " & vbCrLf & _
                              "  				 PASIGoodReceivingQty = CONVERT(NUMERIC(18,0),ISNULL(RPD.GoodRecQty,'0')),  " & vbCrLf & _
                              "  				 PASIDefectQty = CONVERT(NUMERIC(18,0),ISNULL(RPD.DefectRecQty,'0')),  " & vbCrLf & _
                              "  				 PASIDeliveryQty = CONVERT(NUMERIC(18,0),ISNULL(DPD.DOQty,'0')),  " & vbCrLf & _
                              "  				 GoodReceivingQty = CONVERT(NUMERIC(18,0), CASE WHEN POM.DeliveryByPASICls = '1' THEN RTRIM(CONVERT(NUMERIC(18,0),COALESCE(RAD.RecQty,DPD.DOQty,'0')))  " & vbCrLf & _
                              "                                           ELSE RTRIM(CONVERT(NUMERIC(18,0),COALESCE(RAD.RecQty,DSD.DOQty,'0'))) END),  " & vbCrLf & _
                              "  				 DefectReceivingQty = CONVERT(NUMERIC(18,0),ISNULL(RAD.DefectQty,'0')),  "

            ls_SQL = ls_SQL + "  				 RemainingReceivingQty = CONVERT(NUMERIC(18,0),CASE WHEN ISNULL(POM.DeliveryByPASICls,'0') = '1' THEN RTRIM((ISNULL(DPD.DOQty,0) - (ISNULL(RAD.RecQty,0) + ISNULL(RAD.DefectQty,0))))  " & vbCrLf & _
                              "  											  WHEN ISNULL(POM.DeliveryByPASICls,'0') = '0' THEN RTRIM((ISNULL(DSD.DOQty,0) - (ISNULL(RAD.RecQty,0) + ISNULL(RAD.DefectQty,0)))) END),  " & vbCrLf & _
                              "  				 ReceivingQtyBox = CASE WHEN ISNULL(POD.POQtyBox,MPM.QtyBox) IS NULL THEN 0 ELSE " & vbCrLf & _
                              "                                    CEILING(((CASE WHEN POM.DeliveryByPASICls = '1' THEN COALESCE(RAD.RecQty,DPD.DOQty,0) " & vbCrLf & _
                              "                                                    ELSE COALESCE(RAD.RecQty,DSD.DOQty,0) END) + ISNULL(RAD.DefectQty,0)) / ISNULL(POD.POQtyBox,MPM.QtyBox))" & vbCrLf & _
                              "                                    END,  " & vbCrLf & _
                              "                   UnitCls = KD.UnitCls,  " & vbCrLf & _
                              "  				 POM.DeliveryByPASICls,  " & vbCrLf & _
                              "                   IsSaved = (CASE WHEN RAD.RecQty IS NULL AND RAD.DefectQty IS NULL THEN 'NO' ELSE 'YES' END)  " & vbCrLf & _
                              "                   ,sjpasi = DPD.suratjalanno, sjsupplier = DSM.SuratJalanNo  " & vbCrLf & _
                              "  			FROM PO_DETAIL POD   " & vbCrLf & _
                              "  				 LEFT JOIN PO_Master POM ON POM.AffiliateID =POD.AffiliateID  " & vbCrLf & _
                              "  					AND POM.SupplierID =POD.SupplierID  "

            ls_SQL = ls_SQL + "  					AND POM.PONO =POD.PONO  " & vbCrLf & _
                              "  				 LEFT JOIN Kanban_Detail KD ON KD.AffiliateID =POD.AffiliateID  " & vbCrLf & _
                              "  					AND KD.SupplierID =POD.SupplierID  " & vbCrLf & _
                              "  					AND KD.PONO =POD.PONO  " & vbCrLf & _
                              "  					AND KD.PartNo =POD.PartNo  " & vbCrLf & _
                              "  				 LEFT JOIN Kanban_Master KM ON KD.AffiliateID =KM.AffiliateID  " & vbCrLf & _
                              "  					AND KD.SupplierID =KM.SupplierID  " & vbCrLf & _
                              "  					AND KD.KanbanNo =KM.KanbanNo  " & vbCrLf & _
                              "                      AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                              "  				 LEFT JOIN DOSupplier_Detail DSD ON KD.AffiliateID =DSD.AffiliateID  " & vbCrLf & _
                              "  					AND KD.SupplierID =DSD.SupplierID  "

            ls_SQL = ls_SQL + "  					AND KD.PONO =DSD.PONO  " & vbCrLf & _
                              "  					AND KD.PartNo =DSD.PartNo  " & vbCrLf & _
                              "  					AND KD.KanbanNo =DSD.KanbanNo  " & vbCrLf & _
                              "  				 LEFT JOIN DOSupplier_Master DSM ON DSM.AffiliateID =DSD.AffiliateID  " & vbCrLf & _
                              "  					AND DSM.SupplierID =DSD.SupplierID  " & vbCrLf & _
                              "  					AND DSM.SuratJalanNo =DSD.SuratJalanNo  " & vbCrLf & _
                              "  				 LEFT JOIN DOPASI_Detail DPD ON KD.AffiliateID =DPD.AffiliateID  " & vbCrLf & _
                              "  					AND KD.SupplierID =DPD.SupplierID  " & vbCrLf & _
                              "  					AND KD.PONO =DPD.PONO  " & vbCrLf & _
                              "  					AND KD.PartNo =DPD.PartNo  " & vbCrLf & _
                              "  					AND KD.KanbanNo =DPD.KanbanNo  "

            ls_SQL = ls_SQL + "                      AND DPD.SuratjalanNOSupplier = DSM.SuratJalanNo  " & vbCrLf & _
                              "  				 LEFT JOIN DOPASI_Master DPM ON DPM.AffiliateID =DPD.AffiliateID  " & vbCrLf & _
                              "  					--AND DPM.SupplierID =DPD.SupplierID  " & vbCrLf & _
                              "  					AND DPM.SuratJalanNo =DPD.SuratJalanNo  " & vbCrLf & _
                              "  				 LEFT JOIN ReceivePASI_Detail RPD ON RPD.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                              "  					AND RPD.SupplierID = KD.SupplierID  " & vbCrLf & _
                              "  					AND RPD.PONo = KD.PONo  " & vbCrLf & _
                              "  					AND RPD.PartNo = KD.PartNo  " & vbCrLf & _
                              "  					AND RPD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "                      AND RPD.SuratJalanNo = DSM.SuratJalanNo  " & vbCrLf & _
                              "  				 LEFT JOIN ReceiveAffiliate_Detail RAD ON RAD.AffiliateID = KD.AffiliateID  "

            ls_SQL = ls_SQL + "  					AND RAD.SupplierID = KD.SupplierID  " & vbCrLf & _
                              "  					AND RAD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "  					AND RAD.PONo = KD.PONo  " & vbCrLf & _
                              "  					AND RAD.PartNo = KD.PartNo  " & vbCrLf & _
                              "                      AND RAD.SuratJalanNo = DPM.SuratJalanNo  " & vbCrLf & _
                              "  				 LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
                              "  					--AND RAM.SupplierID = RAD.SupplierID  " & vbCrLf & _
                              "  					AND RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
                              "  				 LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                              "  				 LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo AND MPM.AffiliateID = POD.AffiliateID AND MPM.SupplierID = POD.SupplierID  " & vbCrLf & _
                              "  				 LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls  " & vbCrLf & _
                              " 		   WHERE POD.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf

            If Session("DeliveryByPasi") = "1" Then
                ls_SQL = ls_SQL + "              AND DPD.SuratJalanNo = '" & pm_PASISJNo & "'" & vbCrLf
                'ls_SQL = ls_SQL + "              AND DSM.SuratJalanNo = '" & pm_SupplierSJNo & "'" & vbCrLf
            Else
                ls_SQL = ls_SQL + "              AND DSM.SuratJalanNo = '" & pm_SupplierSJNo & "'" & vbCrLf
            End If

            ls_SQL = ls_SQL + "        ) RecEntry "
            ls_SQL = ls_SQL + " GROUP BY  supplier, PONo, KanbanCls,KanbanNo, PartNo, PartName, UOM, QtyBox,UnitCls, DeliveryByPASICls, IsSaved,sjpasi "
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Sub up_IsiMaster(ByVal pSJ As String)
        Dim ls_SQL As String = ""
        Dim ls_HT As String = ""

        'pSJ = "'" & (Replace(pSJ, "'", "")) & "'"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()


            ls_SQL = " select Distinct ReceiveDate = Convert(char(11), convert(datetime, getdate())), " & vbCrLf & _
                  " SupplierID = DOM.SupplierID, " & vbCrLf & _
                  " SupplierName = MS.SupplierName, " & vbCrLf & _
                  " SupplierSJNo = DOD.SuratJalanNoSupplier, " & vbCrLf & _
                  " SupplierPlanDeliveryDate = Convert(char(11), convert(datetime, DSM.DeliveryDate)), " & vbCrLf & _
                  " SupplierDeliveryDate = Convert(char(11), convert(datetime, DSM.DeliveryDate)), " & vbCrLf & _
                  " PASISJNo = DOM.SuratJalanNo, " & vbCrLf & _
                  " PASIDeliveryDate = Convert(char(11), convert(datetime, DOM.DeliveryDate)), " & vbCrLf & _
                  " DriverName = DOM.DriverName, " & vbCrLf & _
                  " DriverContact = DOM.DriverContact, " & vbCrLf & _
                  " Nopol = DOM.Nopol, "

            ls_SQL = ls_SQL + " jenisArmada = DOM.JenisArmada, " & vbCrLf & _
                              " TotalBox = isnull(DOM.TotalBox,0), " & vbCrLf & _
                              " PerformanceCls = '', " & vbCrLf & _
                              " DeliveryLocation = KM.DeliveryLocationCode, " & vbCrLf & _
                              " DeliveryName = MD.DeliveryLocationName, DOM.InvoiceNo " & vbCrLf & _
                              " From DOPasi_Master DOM LEFT JOIN DOPasi_Detail DOD " & vbCrLf & _
                              " ON DOM.SuratJalanNO = DOD.SuratJalanno " & vbCrLf & _
                              " AND DOM.AffiliateID = DOD.AffiliateID " & vbCrLf & _
                              " LEFT JOIN DOSupplier_Detail DSD ON DSD.SuratJalanNo = DOD.SuratJalanNoSupplier " & vbCrLf & _
                              " And DSD.SupplierID = DOD.SupplierID " & vbCrLf & _
                              " AND DSD.AffiliateID = DOD.AffiliateID "

            ls_SQL = ls_SQL + " AND DSD.PartNo = DOD.PartNo " & vbCrLf & _
                              " AND DSD.PONo = DOD.PONo " & vbCrLf & _
                              " AND DSD.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                              " LEFT JOIN DOSupplier_Master DSM ON DSM.SuratJalanno = DSD.SuratJalanNo " & vbCrLf & _
                              " AND DSM.AffiliateID = DSD.AffiliateID " & vbCrLf & _
                              " AND DSM.SupplierID = DSD.SupplierID " & vbCrLf & _
                              " Left Join MS_Affiliate MA ON MA.AffiliateID = DOM.AffiliateID " & vbCrLf & _
                              " Left Join MS_Supplier MS ON MS.SupplierID = DOM.SupplierID " & vbCrLf & _
                              " LEFT JOIN Kanban_Master KM ON KM.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                              " AND KM.AffiliateID = DOD.AffiliateID " & vbCrLf & _
                              " AND KM.SupplierID = DOD.SupplierID "

            ls_SQL = ls_SQL + " LEFT JOIN MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
                              " Where DOM.SuratJalanNo = '" & Trim(pSJ) & "' "



            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Try
                    With ds.Tables(0)
                        txtRecDate.Text = Trim(.Rows(0).Item("ReceiveDate"))
                        'txtSupplierCode.Text = Trim(.Rows(0).Item("SupplierID"))
                        'txtSupplierName.Text = Trim(.Rows(0).Item("SupplierName"))
                        txtSupplierSJNo.Text = Trim(.Rows(0).Item("InvoiceNo"))
                        txtSupplierPlanDeliveryDate.Text = Trim(.Rows(0).Item("SupplierPlanDeliveryDate"))
                        txtSupplierDeliveryDate.Text = Trim(.Rows(0).Item("SupplierDeliveryDate"))
                        txtPASISJNo.Text = Trim(.Rows(0).Item("PASISJNo"))
                        txtPASIDeliveryDate.Text = Trim(.Rows(0).Item("PASIDeliveryDate"))
                        txtDeliveryLocationCode.Text = Trim(.Rows(0).Item("ReceiveDate"))
                        txtDeliveryLocationName.Text = Trim(.Rows(0).Item("ReceiveDate"))
                        txtDriverName.Text = Trim(.Rows(0).Item("DriverName"))
                        txtDriverContact.Text = Trim(.Rows(0).Item("DriverContact"))
                        txtNoPol.Text = Trim(.Rows(0).Item("Nopol"))
                        txtJenisArmada.Text = Trim(.Rows(0).Item("jenisArmada"))
                        txtTotalBox.Text = CInt(Trim(Trim(.Rows(0).Item("TotalBox"))))
                    End With
                Catch ex As Exception

                End Try
            End If
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "SELECT TOP 0 " & vbCrLf & _
                     " 		 ColNo = 0, PONo = '', POKanban = '', KanbanNo = '', PartNo = '', PartName = '', UOM = '', QtyBox = '', " & vbCrLf & _
                     "       SupplierDeliveryQty = '', PASIGoodReceivingQty = '', PASIDefectQty = '', PASIDeliveryQty = '', GoodReceivingQty = '', " & vbCrLf & _
                     "       DefectReceivingQty = '', RemainingReceivingQty = '', ReceivingQtyBox = 0, UnitCls = ''" & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Function uf_SetPerformanceCls(ByVal pSJNo As String) As String
        Dim ls_PerfCls As String = "", ls_Desc As String = ""
        ls_SQL = "SELECT MP.PerformanceCls, MP.Description " & vbCrLf & _
                 "  FROM dbo.ReceiveAffiliate_Master RAM " & vbCrLf & _
                 "       LEFT JOIN dbo.MS_PerformanceCls MP ON MP.PerformanceCls = RAM.PerformanceCls " & vbCrLf & _
                 " WHERE RAM.SuratJalanNo = '" & pSJNo & "'" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                ls_PerfCls = ds.Tables(0).Rows(0).Item("PerformanceCls")
                ls_Desc = ds.Tables(0).Rows(0).Item("Description")
            End If

            'Set
            uf_SetPerformanceCls = "cboPerformanceCls.SetText('" & ls_PerfCls & "'); txtPerformanceDesc.SetText('" & ls_Desc & "');"
        End Using
    End Function

    Private Sub up_SaveData()
        Try
            Dim sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            Dim sqlCmd As New SqlCommand
            Dim iLoop As Integer = 0

            Using scope As New TransactionScope
                sqlConn.Open()

                Dim ls_SuratJalanNo As String = "", ls_SupplierID As String = "", ls_ReceiveDate As String = "", ls_ReceiveBy As String = "", _
                    ls_JenisArmada As String = "", ls_DriverName As String = "", ls_DriverContact As String = "", ls_NoPol As String = "", _
                    ls_TotalBox As String = "0", ls_PerformanceCls As String = ""

                Dim ls_PONo As String = "", ls_POKanbanCls As String = "", ls_KanbanNo As String = "", ls_PartNo As String = "", _
                    ls_UnitCls As String = "", ls_RecQty As String = "0", ls_DefQty As String = "0"


                'INSERT MASTER
                If txtPASISJNo.Text <> "" Then ls_SuratJalanNo = txtPASISJNo.Text Else ls_SuratJalanNo = txtSupplierSJNo.Text
                'ls_SupplierID = txtSupplierCode.Text
                ls_ReceiveDate = Format(CDate(txtRecDate.Text), "yyyy-MM-dd") & " " & Right(Trim(txtRecDate.Text), 8)
                ls_ReceiveBy = Session("UserID")
                ls_JenisArmada = txtJenisArmada.Text
                ls_DriverName = txtDriverName.Text
                ls_DriverContact = txtDriverContact.Text
                ls_NoPol = txtNoPol.Text
                ls_PerformanceCls = cboPerformanceCls.Text

                If txtTotalBox.Text = "" Then ls_TotalBox = "0" Else ls_TotalBox = txtTotalBox.Text

                'ls_SQL = uf_SaveAffiliateMaster(ls_SuratJalanNo, ls_SupplierID, ls_ReceiveDate, ls_ReceiveBy, ls_JenisArmada, ls_DriverName, ls_DriverContact, ls_NoPol, ls_TotalBox, ls_PerformanceCls)
                'sqlCmd = New SqlCommand(ls_SQL, sqlConn)
                'sqlCmd.ExecuteNonQuery()
                'sqlCmd.Dispose()


                'INSERT DETAIL
                With grid
                    For iLoop = 0 To .VisibleRowCount - 1

                        If .GetRowValues(iLoop, "QtyBox") = 0 Then
                            Session("E02Msg") = "Qty/Box not found in Part Mapping Master, please check again with PASI!"
                            Exit Sub
                        End If

                        ls_PONo = .GetRowValues(iLoop, "PONo")
                        If .GetRowValues(iLoop, "POKanban") = "YES" Then ls_POKanbanCls = "1" Else ls_POKanbanCls = "0"
                        ls_KanbanNo = .GetRowValues(iLoop, "KanbanNo")
                        ls_PartNo = .GetRowValues(iLoop, "PartNo")
                        ls_UnitCls = .GetRowValues(iLoop, "UnitCls")
                        ls_RecQty = .GetRowValues(iLoop, "GoodReceivingQty")
                        ls_DefQty = .GetRowValues(iLoop, "DefectReceivingQty")
                        ls_SupplierID = .GetRowValues(iLoop, "Supplier")

                        'INSERT MASTER
                        ls_SQL = uf_SaveAffiliateMaster(ls_SuratJalanNo, ls_SupplierID, ls_ReceiveDate, ls_ReceiveBy, ls_JenisArmada, ls_DriverName, ls_DriverContact, ls_NoPol, ls_TotalBox, ls_PerformanceCls)
                        sqlCmd = New SqlCommand(ls_SQL, sqlConn)
                        sqlCmd.ExecuteNonQuery()
                        sqlCmd.Dispose()

                        'INSERT DETAIL
                        ls_SQL = uf_SaveAffiliateDetail(ls_SuratJalanNo, ls_SupplierID, ls_PONo, ls_POKanbanCls, ls_KanbanNo, ls_PartNo, ls_UnitCls, ls_RecQty, ls_DefQty)
                        sqlCmd = New SqlCommand(ls_SQL, sqlConn)
                        sqlCmd.ExecuteNonQuery()
                        sqlCmd.Dispose()
                    Next iLoop
                End With

                Call clsMsg.DisplayMessage(lblInfo, "1002", clsMessage.MsgType.InformationMessage)
                Session("E02Msg") = lblInfo.Text
                scope.Complete()
            End Using

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("E02Msg") = lblInfo.Text
        End Try
    End Sub

    Private Sub up_Delete()
        Try
            Dim sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            Dim sqlCmd As New SqlCommand
            Dim iLoop As Integer = 0

            Using scope As New TransactionScope
                sqlConn.Open()

                Dim ls_SuratJalanNo As String = "", ls_SupplierID As String = ""
                Dim ls_PONo As String = "", ls_PartNo As String = "", ls_kanbanno As String = ""

                If txtPASISJNo.Text <> "" Then ls_SuratJalanNo = txtPASISJNo.Text Else ls_SuratJalanNo = txtSupplierSJNo.Text
                'ls_SupplierID = txtSupplierCode.Text

                'DELETE DETAIL
                ls_SQL = uf_DeleteAffiliateDetail(ls_SuratJalanNo)
                sqlCmd = New SqlCommand(ls_SQL, sqlConn)
                sqlCmd.ExecuteNonQuery()
                sqlCmd.Dispose()

                ls_SQL = uf_DeleteAffiliateDetailSeq(ls_SuratJalanNo)
                sqlCmd = New SqlCommand(ls_SQL, sqlConn)
                sqlCmd.ExecuteNonQuery()
                sqlCmd.Dispose()

                'DELETE MASTER
                ls_SQL = uf_DeleteAffiliateMaster(ls_SuratJalanNo)
                sqlCmd = New SqlCommand(ls_SQL, sqlConn)
                sqlCmd.ExecuteNonQuery()
                sqlCmd.Dispose()


                scope.Complete()
            End Using

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("E02Msg") = lblInfo.Text
        End Try
    End Sub
#End Region

#Region "Functions"
    Private Function uf_SaveAffiliateMaster(ByVal pSuratJalanNo As String, ByVal pSupplierID As String, ByVal pReceiveDate As String, _
                                            ByVal pReceiveBy As String, ByVal pJenisArmada As String, ByVal pDriverName As String, _
                                            ByVal pDriverContact As String, ByVal pNoPol As String, ByVal pTotalBox As String, _
                                            ByVal pPerformanceCls As String) As String

        ls_SQL = ""
        ls_SQL = ls_SQL + " IF NOT EXISTS (SELECT SuratJalanNo FROM ReceiveAffiliate_Master WHERE SuratJalanNo = '" & pSuratJalanNo & "' AND AffiliateID = '" & Session("AffiliateID") & "' AND SupplierID = '" & pSupplierID & "' ) " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " 	INSERT INTO ReceiveAffiliate_Master (SuratJalanNo,AffiliateID,SupplierID,ReceiveDate,ReceiveBy,JenisArmada,DriverName,DriverContact,NoPol,TotalBox,EntryDate,EntryUser,UpdateDate,UpdateUser,HT_Cls,PerformanceCls) " & vbCrLf & _
                          " 	VALUES('" & pSuratJalanNo & "','" & Session("AffiliateID") & "','" & pSupplierID & "','" & pReceiveDate & "','" & pReceiveBy & "','" & pJenisArmada & "','" & pDriverName & "','" & pDriverContact & "','" & pNoPol & "'," & pTotalBox & ",GETDATE(),'" & Session("UserID") & "',NULL,NULL,'0','" & pPerformanceCls & "') " & vbCrLf & _
                          " END " & vbCrLf & _
                          " ELSE " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " 	UPDATE ReceiveAffiliate_Master " & vbCrLf & _
                          " 	   SET ReceiveDate = '" & pReceiveDate & "', " & vbCrLf & _
                          " 		   ReceiveBy = '" & pReceiveBy & "', " & vbCrLf & _
                          " 		   JenisArmada = '" & pJenisArmada & "', " & vbCrLf
        ls_SQL = ls_SQL + " 		   DriverName = '" & pDriverName & "', " & vbCrLf & _
                          " 		   DriverContact = '" & pDriverContact & "', " & vbCrLf & _
                          " 		   NoPol = '" & pNoPol & "', " & vbCrLf & _
                          " 		   TotalBox = " & pTotalBox & ", " & vbCrLf & _
                          " 		   PerformanceCls = '" & pPerformanceCls & "', " & vbCrLf & _
                          " 		   UpdateDate = GETDATE(), " & vbCrLf & _
                          " 		   UpdateUser = '" & Session("UserID") & "' " & vbCrLf & _
                          " 	 WHERE SuratJalanNo = '" & pSuratJalanNo & "' AND AffiliateID = '" & Session("AffiliateID") & "' AND SupplierID = '" & pSupplierID & "' " & vbCrLf & _
                          " END " & vbCrLf & _
                          "  "
        Return ls_SQL
    End Function

    Private Function uf_SaveAffiliateDetail(ByVal pSuratJalanNo As String, ByVal pSupplierID As String, ByVal pPONo As String, ByVal pPOKanbanCls As String, ByVal pKanbanNo As String, _
                                            ByVal pPartNo As String, ByVal pUnitCls As String, ByVal pRecQty As String, ByVal pDefectQty As String) As String

        ls_SQL = ""
        ls_SQL = ls_SQL + " IF NOT EXISTS (SELECT SuratJalanNo FROM ReceiveAffiliate_Detail WHERE SuratJalanNo = '" & pSuratJalanNo & "' AND AffiliateID = '" & Session("AffiliateID") & "' AND SupplierID = '" & pSupplierID & "' AND PONo = '" & pPONo & "' AND PartNo = '" & pPartNo & "' AND KanbanNo = '" & pKanbanNo & "') " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " 	INSERT INTO ReceiveAffiliate_Detail (SuratJalanNo,SupplierID,AffiliateID,PONo,POKanbanCls,KanbanNo,PartNo,UnitCls,RecQty,DefectQty) " & vbCrLf & _
                          " 	VALUES('" & pSuratJalanNo & "','" & pSupplierID & "','" & Session("AffiliateID") & "','" & pPONo & "','" & pPOKanbanCls & "','" & pKanbanNo & "','" & pPartNo & "','" & pUnitCls & "'," & pRecQty & "," & pDefectQty & ") " & vbCrLf & _
                          " END " & vbCrLf & _
                          " ELSE " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " 	UPDATE ReceiveAffiliate_Detail " & vbCrLf & _
                          " 	   SET POKanbanCls = '" & pPOKanbanCls & "', " & vbCrLf & _
                          " 		   UnitCls = '" & pUnitCls & "', " & vbCrLf
        ls_SQL = ls_SQL + " 		   RecQty = " & pRecQty & ", " & vbCrLf & _
                          " 		   DefectQty = " & pDefectQty & " " & vbCrLf & _
                          " 	 WHERE SuratJalanNo = '" & pSuratJalanNo & "' AND AffiliateID = '" & Session("AffiliateID") & "'  " & vbCrLf & _
                          " 	   AND SupplierID = '" & pSupplierID & "' AND PONo = '" & pPONo & "' AND PartNo = '" & pPartNo & "' " & vbCrLf & _
                          "        AND KanbanNo = '" & pKanbanNo & "' " & vbCrLf & _
                          " END " & vbCrLf & _
                          "  "

        Return ls_SQL
    End Function

    Private Function uf_DeleteAffiliateMaster(ByVal pSuratJalanNo As String) As String
        ls_SQL = "DELETE FROM ReceiveAffiliate_Master WHERE SuratJalanNo = '" & pSuratJalanNo & "' AND AffiliateID = '" & Session("AffiliateID") & "'"

        Return ls_SQL
    End Function

    Private Function uf_DeleteAffiliateDetail(ByVal pSuratJalanNo As String) As String
        ls_SQL = "DELETE FROM ReceiveAffiliate_Detail " & vbCrLf & _
                 " WHERE SuratJalanNo = '" & pSuratJalanNo & "' " & vbCrLf & _
                 "   AND AffiliateID = '" & Session("AffiliateID") & "' " 

        Return ls_SQL
    End Function

    Private Function uf_DeleteAffiliateDetailSeq(ByVal pSuratJalanNo As String) As String
        ls_SQL = "DELETE FROM ReceiveAffiliateSeq_Detail " & vbCrLf & _
                 " WHERE SuratJalanNo = '" & pSuratJalanNo & "' " & vbCrLf & _
                 "   AND AffiliateID = '" & Session("AffiliateID") & "' "

        Return ls_SQL
    End Function

    Private Function uf_CalculateBox() As String
        Dim retValue As String = "0"
        Dim iRow As Integer = 0

        With grid
            For iRow = 0 To .VisibleRowCount - 1
                retValue = CDbl(retValue) + CDbl(.GetRowValues(iRow, "ReceivingQtyBox"))
            Next iRow
        End With

        Return retValue
    End Function

    Private Sub ExcelBC40()
        Call HeaderExcelBC40()
        Call DetailExcelBC40()
        FileName = "Template BC4.0.xlsx"
        FilePath = Server.MapPath("~\Template\" & FileName)
        Call epplusExportHeaderExcel(FilePath, "", dtHeader, "A:17", "")

        'Call epplusExportExcel(FilePath, "KEDUA", dtDetail, "A:7", "")
    End Sub

    Private Sub HeaderExcelBC40()

        Dim ds As New DataSet
        Dim ls_sql As String = ""

        ls_sql = " SELECT DISTINCT " & vbCrLf & _
                  " IzinTPB = (Select Rtrim(IzinTPB) from MS_Affiliate where AffiliateID = PLM.AffiliateID), " & vbCrLf & _
                  " BCPerson = (Select Rtrim(BCPerson) from MS_Affiliate where AffiliateID = PLM.AffiliateID), " & vbCrLf & _
                  " KantorPabean = (Select Rtrim(KantorPabean) from MS_Affiliate where AffiliateID = PLM.AffiliateID), " & vbCrLf & _
                  " NPWP = (Select Rtrim(NPWP) from MS_Affiliate where AffiliateID = PLM.AffiliateID), " & vbCrLf & _
                  " Buyer = (Select Rtrim(AffiliateName) from MS_Affiliate where AffiliateID = PLM.AffiliateID),  " & vbCrLf & _
                  " AlamatBuyer =(Select Rtrim(Address) from MS_Affiliate where AffiliateID = PLM.AffiliateID),   " & vbCrLf & _
                  " City =(Select Rtrim(City) from MS_Affiliate where AffiliateID = PLM.AffiliateID),   " & vbCrLf & _
                  " NPWPPengirim = (Select Rtrim(NPWP) from MS_Affiliate where AffiliateID = 'PASI'), " & vbCrLf & _
                  " Pengirim = (Select Rtrim(AffiliateName) from MS_Affiliate where AffiliateID = 'PASI'),  " & vbCrLf & _
                  " AlamatPengirim =(Select Rtrim(Address) from MS_Affiliate where AffiliateID = 'PASI'),   " & vbCrLf & _
                  " ShipCls = Rtrim(isnull(POM.ShipCls,'')), " & vbCrLf & _
                  " NoPol = Rtrim(isnull(PLM.NoPol,'')),   " & vbCrLf & _
                  " InvoiceNo = Rtrim(coalesce(PLM.InvoiceNo,'-')),   " & vbCrLf & _
                  " Invdate = Coalesce(DPM.DeliveryDate, DSM.DeliveryDate),  " & vbCrLf & _
                  " PONo = (SELECT (STUFF((SELECT distinct ', ' + RTrim(PLPASI_Detail.PONo) FROM PLPASI_Detail WHERE SuratJalanNo='" & Trim(txtPASISJNo.Text) & "' AND AffiliateID='" & Session("AffiliateID") & "' FOR XML PATH('')), 1, 2, ''))), " & vbCrLf & _
                  " PODate = POM.EntryDate, " & vbCrLf

        ls_sql = ls_sql + " Currency = isnull(MC.Description,''), " & vbCrLf & _
                          " JumlahHarga = (SELECT SUM(isnull(DPD.Price,0)*DOQty) FROM PLPASI_Detail PLPD " & vbCrLf & _
                          " LEFT JOIN MS_Price MPr1 ON MPr1.PartNo = PLPD.PartNo  " & vbCrLf & _
                          " AND PLPD.AffiliateID = MPr1.AffiliateID WHERE PLPD.SuratJalanNo='" & Trim(txtPASISJNo.Text) & "' AND PLPD.AffiliateID='" & Session("AffiliateID") & "' AND (DPM.DeliveryDate between Mpr1.StartDate and Mpr1.EndDate)), " & vbCrLf & _
                          " JumlahKemasan =  (SELECT SUM(CONVERT(NUMERIC,ISNULL(CartonQty,0))) FROM PLPASI_Detail WHERE SuratJalanNo='" & Trim(txtPASISJNo.Text) & "' AND AffiliateID='" & Session("AffiliateID") & "'), " & vbCrLf & _
                          " JumlahQty =  (SELECT SUM(CONVERT(NUMERIC,ISNULL(DOQty,0))) FROM PLPASI_Detail WHERE SuratJalanNo='" & Trim(txtPASISJNo.Text) & "' AND AffiliateID='" & Session("AffiliateID") & "'), " & vbCrLf & _
                          " BeratBersih = (SELECT SUM(ISNULL(CartonQty,0) * (b.NetWeight/1000)) FROM PLPASI_Detail a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID WHERE SuratJalanNo='" & Trim(txtPASISJNo.Text) & "' AND a.AffiliateID='" & Session("AffiliateID") & "'), " & vbCrLf & _
                          " BeratKotor = (SELECT SUM(ISNULL(CartonQty,0) * (b.GrossWeight/1000)) FROM PLPASI_Detail a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID WHERE SuratJalanNo='" & Trim(txtPASISJNo.Text) & "' AND a.AffiliateID='" & Session("AffiliateID") & "') " & vbCrLf & _
                          " FROM PLPASI_Detail PLD " & vbCrLf & _
                          " LEFT JOIN PLPASI_Master PLM ON PLM.SuratJalanNo = PLD.SuratJalanNo AND PLM.AffiliateID = PLD.AffiliateID  " & vbCrLf & _
                          " LEFT JOIN PO_Master POM ON POM.PONo = PLD.PONo AND POM.AffiliateID = PLD.AffiliateID AND POM.SupplierID = PLD.SupplierID " & vbCrLf & _
                          " LEFT JOIN DOPasi_Detail DPD  ON DPD.SuratJalanNo = PLD.SuratJalanNo   " & vbCrLf & _
                          "   	AND DPD.SupplierID = PLD.SupplierID   " & vbCrLf

        ls_sql = ls_sql + "   	AND DPD.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
                          "   	AND DPD.PONo = PLD.PONo   " & vbCrLf & _
                          " LEFT JOIN DOPASI_Master DPM  ON DPM.SuratJalanNo = DPD.SuratJalanNo  	  " & vbCrLf & _
                          "   	--AND DPD.SupplierID = DPM.SupplierID   " & vbCrLf & _
                          "   	AND DPD.AffiliateID = DPM.AffiliateID     " & vbCrLf & _
                          " LEFT JOIN DOSupplier_Detail DSD ON DSD.SuratJalanNo = PLD.SuratJalanNo   " & vbCrLf & _
                          "   	AND DSD.SupplierID = PLD.SupplierID   " & vbCrLf & _
                          "   	AND DSD.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
                          "   	AND DSD.PONo = PLD.PONo   " & vbCrLf & _
                          " LEFT JOIN DOSUPPLIER_Master DSM ON DSM.SuratJalanNo = DSD.SuratJalanNo  	  " & vbCrLf & _
                          "   	AND DSD.SupplierID = DSM.SupplierID   " & vbCrLf

        ls_sql = ls_sql + "   	AND DSD.AffiliateID = DSM.AffiliateID   " & vbCrLf & _
                          " LEFT JOIN MS_Parts MP ON MP.PartNo = PLD.PartNo  " & vbCrLf & _
                          " LEFT JOIN MS_Price MPr ON MPr.PartNo = PLD.PartNo AND PLD.AffiliateID = MPr.AffiliateID " & vbCrLf & _
                          " 	AND MPR.PartNo = PLD.PartNo and COALESCE(DPM.DeliveryDate,DSM.DeliveryDate) between MPR.StartDate and MPR.EndDate " & vbCrLf & _
                          " LEFT JOIN MS_CurrCls MC ON MC.CurrCls = MPr.CurrCls " & vbCrLf

        ls_sql = ls_sql + " WHERE PLM.SuratJalanNo='" & Trim(txtPASISJNo.Text) & "' AND PLM.AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
                          " GROUP BY POM.EntryDate, PLD.PONo, PLM.AffiliateID,PLM.ViaDelivery,PLM.InvoiceNo,POM.ShipCls,DPM.DeliveryDate,DSM.DeliveryDate,MC.DESCRIPTION,PLM.NoPol " & vbCrLf


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using
        dtHeader = ds.Tables(0)
    End Sub

    Private Sub DetailExcelBC40()

        Dim ds As New DataSet
        Dim ls_sql As String = ""
        ls_sql = "Select Row = ROW_NUMBER() OVER (ORDER BY PartNo, PartName, Currency, KanbanNo) ,PartNo,PartName,Currency, Qty = SUM(QTY),Harga = SUM(Harga), KanbanNo from (" & vbCrLf
        ls_sql = ls_sql + " SELECT DISTINCT " & vbCrLf & _
                  " Row =  ROW_NUMBER() OVER (ORDER BY PLD.PartNo), " & vbCrLf & _
                  " PartNo = Rtrim(MP.PartNo), PartName = Rtrim(MP.PartGroupName), " & vbCrLf & _
                  " Qty =  CONVERT(NUMERIC,ISNULL(PLD.DOQty,0)), " & vbCrLf & _
                  " Currency = isnull(MC.Description,''), " & vbCrLf & _
                  " Harga = isnull(DPD.Price,0)*PLD.DOQty " & vbCrLf & _
                  " ,PLD.KanbanNo " & vbCrLf & _
                  " FROM PLPASI_Detail PLD " & vbCrLf & _
                  " LEFT JOIN PLPASI_Master PLM ON PLM.SuratJalanNo = PLD.SuratJalanNo  " & vbCrLf & _
                  " 	AND PLM.AffiliateID = PLD.AffiliateID  " & vbCrLf & _
                  " LEFT JOIN MS_Parts MP ON MP.PartNo = PLD.PartNo  " & vbCrLf

        ls_sql = ls_sql + " LEFT JOIN MS_Price MPr ON MPr.PartNo = PLD.PartNo  " & vbCrLf & _
                          " 	AND PLD.AffiliateID = MPr.AffiliateID " & vbCrLf & _
                          "     and PLM.DeliveryDate between MPR.StartDate and MPR.EndDate " & vbCrLf & _
                          " LEFT JOIN MS_CurrCls MC ON MC.CurrCls = MPr.CurrCls " & vbCrLf & _
                          " LEFT JOIN PO_Master POM ON POM.PONo = PLD.PONo  " & vbCrLf & _
                          " 	AND POM.AffiliateID = PLD.AffiliateID AND POM.SupplierID = PLD.SupplierID " & vbCrLf & _
                          " LEFT JOIN DOPasi_Detail DPD  ON DPD.SuratJalanNo = PLD.SuratJalanNo   " & vbCrLf & _
                          "   	AND DPD.SupplierID = PLD.SupplierID   " & vbCrLf & _
                          "   	AND DPD.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
                          "   	AND DPD.PONo = PLD.PONo   " & vbCrLf & _
                          " LEFT JOIN DOPASI_Master DPM  ON DPM.SuratJalanNo = DPD.SuratJalanNo  	  " & vbCrLf & _
                          "   	AND DPD.SupplierID = DPM.SupplierID   " & vbCrLf

        ls_sql = ls_sql + "   	AND DPD.AffiliateID = DPM.AffiliateID     " & vbCrLf & _
                          " LEFT JOIN DOSupplier_Detail DSD ON DSD.SuratJalanNo = PLD.SuratJalanNo   " & vbCrLf & _
                          "   	AND DSD.SupplierID = PLD.SupplierID   " & vbCrLf & _
                          "   	AND DSD.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
                          "   	AND DSD.PONo = PLD.PONo   " & vbCrLf & _
                          " LEFT JOIN DOSUPPLIER_Master DSM ON DSM.SuratJalanNo = DSD.SuratJalanNo  	  " & vbCrLf & _
                          "   	AND DSD.SupplierID = DSM.SupplierID   " & vbCrLf & _
                          "   	AND DSD.AffiliateID = DSM.AffiliateID   " & vbCrLf

        ls_sql = ls_sql + " WHERE PLM.SuratJalanNo='" & Trim(txtPASISJNo.Text) & "' AND PLM.AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
                          " GROUP BY PLD.PartNo,MP.PartNo,MP.PartName,PLD.DOQty,MC.DESCRIPTION,DPD.Price,MP.PartGroupName,PLD.kanbanno" & vbCrLf
        ls_sql = ls_sql + " )x Group by PartNo, PartName, Currency, KanbanNo"


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using
        dtDetail = ds.Tables(0)
    End Sub

    Private Sub epplusExportHeaderExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "BC 4.0 " & Trim(txtPASISJNo.Text) & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Receiving\" & tempFile)

            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet
            Dim wsCount As ExcelWorksheet

            ws = exl.Workbook.Worksheets("form (2)")
            Dim irow As Integer = 0
            Dim irowxl As Integer = 0

            With ws
                If pData.Rows.Count > 0 Then
                    .Cells("C5").Value = ": " & pData.Rows(irow)("KantorPabean") 'KantorPabean
                    .Cells("C11").Value = ": " & pData.Rows(irow)("NPWP")
                    .Cells("C12").Value = ": " & pData.Rows(irow)("Buyer")
                    .Cells("C13:E14").Merge = True
                    .Cells("C13:E14").Style.WrapText = True
                    .Cells("C13:E14").Value = ": " & pData.Rows(irow)("AlamatBuyer")
                    .Cells("I12").Value = "" & pData.Rows(irow)("NPWPPengirim")
                    .Cells("I13").Value = ": " & pData.Rows(irow)("Pengirim")
                    .Cells("I14:J16").Merge = True
                    .Cells("I14:J16").Style.WrapText = True
                    .Cells("I14:J16").Value = ": " & pData.Rows(irow)("AlamatPengirim")

                    .Cells("C18").Value = ": " & pData.Rows(irow)("InvoiceNo")
                    .Cells("E18").Value = Format(pData.Rows(irow)("Invdate"), "dd-MMM-yyyy")

                    .Cells("C21").Value = ": " & pData.Rows(irow)("PONo")
                    .Cells("E21").Value = Format(pData.Rows(irow)("PODate"), "dd-MMM-yyyy")

                    .Cells("B23").Value = "Jenis Sarana Pengangkut darat : " & pData.Rows(irow)("ShipCls")
                    .Cells("I23").Value = ": " & pData.Rows(irow)("NoPol")

                    .Cells("C25").Value = pData.Rows(irow)("JumlahHarga")
                    If .Cells("C25").Value <> "0" Then
                        .Cells("C25").Style.Numberformat.Format = "#,###"
                    End If

                    .Cells("H28").Value = pData.Rows(irow)("JumlahKemasan")

                    .Cells("H31").Value = Format(pData.Rows(irow)("BeratKotor"), "###,##0.0#") & " Kg"
                    .Cells("J31").Value = Format(pData.Rows(irow)("BeratBersih"), "###,##0.0#") & " Kg"

                    .Cells("I54").Value = pData.Rows(irow)("City")

                    .Cells("J54").FormulaR1C1 = "=R[-36]C[-5]"
                    .Cells("H60").Value = "(.........." & pData.Rows(irow)("BCPerson") & "..........)"

                End If
            End With

            ws = exl.Workbook.Worksheets("Lembar Lanjutan (2)")
            With ws
                If pData.Rows.Count > 0 Then
                    .Cells("C8").Value = ": " & pData.Rows(irow)("KantorPabean")
                    .Cells("F62").Value = pData.Rows(irow)("City")
                    .Cells("G62").FormulaR1C1 = "='form (2)'!R[-8]C[3]"
                    .Cells("F65").Value = "(.........." & pData.Rows(irow)("BCPerson") & "..........)"
                End If
            End With

            ws = exl.Workbook.Worksheets("Lembar Lanjutan (3)")
            With ws
                If pData.Rows.Count > 0 Then
                    .Cells("C8").Value = ": " & pData.Rows(irow)("KantorPabean")
                    .Cells("F62").Value = pData.Rows(irow)("City")
                    .Cells("G62").FormulaR1C1 = "='form (2)'!R[-8]C[3]"
                    .Cells("F65").Value = "(.........." & pData.Rows(irow)("BCPerson") & "..........)"
                End If
            End With

            For irow = 0 To dtDetail.Rows.Count - 1
                If irow <= 12 Then
                    'Sheet pertama muat 13 item
                    If irow = 0 Then irowxl = 38
                    wsCount = exl.Workbook.Worksheets("form (2)")
                    For icol = 1 To dtDetail.Columns.Count
                        wsCount.Cells("A" & irowxl).Value = dtDetail.Rows(irow)("Row")
                        wsCount.Cells("B" & irowxl).Value = dtDetail.Rows(irow)("PartName")
                        wsCount.Cells("D" & irowxl).Value = dtDetail.Rows(irow)("PartNo")
                        wsCount.Cells("H" & irowxl).Value = Format(dtDetail.Rows(irow)("Qty"), "#,##0") & " PCS"
                        wsCount.Cells("J" & irowxl).Value = dtDetail.Rows(irow)("Harga")
                        If wsCount.Cells("J" & irowxl).Value <> "0" Then
                            wsCount.Cells("J" & irowxl).Style.Numberformat.Format = "_([$Rp-421]* #,##0_);_([$Rp-421]* (#,##0);_([$Rp-421]* ""-""_);_(@_)"
                        End If
                    Next
                ElseIf irow > 12 And irow <= 54 Then
                    'Sheet kedua muat 42 item
                    If irow = 13 Then irowxl = 17
                    wsCount = exl.Workbook.Worksheets("Lembar Lanjutan (2)")
                    For icol = 1 To dtDetail.Columns.Count
                        wsCount.Cells("A" & irowxl).Value = dtDetail.Rows(irow)("Row")
                        wsCount.Cells("B" & irowxl).Value = dtDetail.Rows(irow)("PartName")
                        wsCount.Cells("C" & irowxl).Value = dtDetail.Rows(irow)("PartNo")
                        wsCount.Cells("F" & irowxl).Value = Format(dtDetail.Rows(irow)("Qty"), "#,##0") & " PCS"
                        wsCount.Cells("H" & irowxl).Value = dtDetail.Rows(irow)("Harga")
                        If wsCount.Cells("H" & irowxl).Value <> "0" Then
                            wsCount.Cells("H" & irowxl).Style.Numberformat.Format = "_([$Rp-421]* #,##0_);_([$Rp-421]* (#,##0);_([$Rp-421]* ""-""_);_(@_)"
                        End If
                    Next
                ElseIf irow > 54 And irow <= 96 Then
                    'Sheet kedua muat 42 item
                    If irow = 55 Then irowxl = 17
                    wsCount = exl.Workbook.Worksheets("Lembar Lanjutan (3)")
                    For icol = 1 To dtDetail.Columns.Count
                        wsCount.Cells("A" & irowxl).Value = dtDetail.Rows(irow)("Row")
                        wsCount.Cells("B" & irowxl).Value = dtDetail.Rows(irow)("PartName")
                        wsCount.Cells("C" & irowxl).Value = dtDetail.Rows(irow)("PartNo")
                        wsCount.Cells("F" & irowxl).Value = Format(dtDetail.Rows(irow)("Qty"), "#,##0") & " PCS"
                        wsCount.Cells("H" & irowxl).Value = dtDetail.Rows(irow)("Harga")
                        If wsCount.Cells("H" & irowxl).Value <> "0" Then
                            wsCount.Cells("H" & irowxl).Style.Numberformat.Format = "_([$Rp-421]* #,##0_);_([$Rp-421]* (#,##0);_([$Rp-421]* ""-""_);_(@_)"
                        End If
                    Next
                End If

                irowxl = irowxl + 1
            Next

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Receiving\" & tempFile & "")
            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub
#End Region

#Region "Form Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim ls_QueryString As String = ""

            If Not IsNothing(Request.QueryString("prm")) Then
                ls_QueryString = Request.QueryString("prm").ToString()
                Session("E02ParamPageLoad") = Request.QueryString("prm").ToString()

                pm_ReceivedDate = Split(ls_QueryString, "|")(0)
                pm_SupplierCode = Split(ls_QueryString, "|")(1)
                pm_SupplierName = Split(ls_QueryString, "|")(2)
                pm_SupplierSJNo = Split(ls_QueryString, "|")(3)
                pm_SupplierPlanDeliveryDate = Split(ls_QueryString, "|")(4)
                pm_SupplierDeliveryDate = Split(ls_QueryString, "|")(5)
                pm_PASISJNo = Split(ls_QueryString, "|")(6)
                pm_PASIDeliveryDate = Split(ls_QueryString, "|")(7)
                pm_DeliveryLocationCode = Split(ls_QueryString, "|")(8)
                pm_DeliveryLocationName = Split(ls_QueryString, "|")(9)
                pm_DriverName = Split(ls_QueryString, "|")(10)
                pm_DriverContact = Split(ls_QueryString, "|")(11)
                pm_NoPol = Split(ls_QueryString, "|")(12)
                pm_JenisArmada = Split(ls_QueryString, "|")(13)
                'pm_TotalBox = Split(ls_QueryString, "|")(14)
                pm_PONo = Split(ls_QueryString, "|")(15)
                pm_KanbanNo = Split(ls_QueryString, "|")(16)
                pm_DeliveryBypasi = Split(ls_QueryString, "|")(17)
                Session("E02KanbanNo") = pm_KanbanNo
                Session("DeliveryBypasi") = pm_DeliveryBypasi

                If (Not IsPostBack) AndAlso (Not IsCallback) Then
                    Call up_Initialize()
                    Call up_FillCombo()

                    txtRecDate.Text = pm_ReceivedDate
                    'txtSupplierCode.Text = pm_SupplierCode
                    'txtSupplierName.Text = pm_SupplierName
                    txtSupplierSJNo.Text = pm_SupplierSJNo
                    txtSupplierPlanDeliveryDate.Text = pm_SupplierPlanDeliveryDate
                    txtSupplierDeliveryDate.Text = pm_SupplierDeliveryDate
                    txtPASISJNo.Text = pm_PASISJNo
                    txtPASIDeliveryDate.Text = pm_PASIDeliveryDate
                    txtDeliveryLocationCode.Text = pm_DeliveryLocationCode
                    txtDeliveryLocationName.Text = pm_DeliveryLocationName
                    txtDriverName.Text = pm_DriverName
                    txtDriverContact.Text = pm_DriverContact
                    txtNoPol.Text = pm_NoPol
                    txtJenisArmada.Text = pm_JenisArmada
                    If Trim(pm_TotalBox) <> "" Then
                        txtTotalBox.Text = CInt(pm_TotalBox)
                    End If


                    Call up_GridLoad()

                    Dim ls_SJNo As String = ""
                    ls_SJNo = IIf(pm_PASISJNo = "", pm_SupplierSJNo, pm_PASISJNo)

                    'ScriptManager.RegisterStartupScript(cboPerformanceCls, cboPerformanceCls.GetType(), "SetPerformCls", uf_SetPerformanceCls(ls_SJNo), True)
                    ScriptManager.RegisterStartupScript(txtTotalBox, txtTotalBox.GetType(), "SetTotalBox", "txtTotalBox.SetText('" & uf_CalculateBox() & "'); " & uf_SetPerformanceCls(ls_SJNo), True)
                End If
            ElseIf Not IsNothing(Request.QueryString("id2")) Then
                ls_QueryString = Request.QueryString("id2").ToString()
                Session("DeliveryBypasi") = "1"

                Call up_IsiMaster(clsNotification.DecryptURL(Request.QueryString("id2")))
                pm_PASISJNo = clsNotification.DecryptURL(Request.QueryString("id2"))
                Call up_GridLoad()
            End If
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("E02Msg")
        Session.Remove("E02KanbanNo")
        Session.Remove("E02ParamPageLoad")

        Session("fromE02") = "true"
        Response.Redirect("~/Receiving/SuppPASIDeliveryConf.aspx")
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Try
            Dim ls_SuratJalanNo As String = "", ls_AffilateID As String = Session("AffiliateID"), ls_SupplierID As String = "", _
                ls_PONo As String = "", ls_POKanbanCls As String = "", ls_KanbanNo As String = "", ls_PartNo As String = "", _
                ls_UnitCls As String = "", ls_RecQty As String = "0", ls_DefQty As String = "0", ls_DeliveryByPASICls As String = "", _
                ls_Delivery As String = ""
            Dim iLoop As Integer = 0

            If txtPASISJNo.Text <> "" Then ls_SuratJalanNo = txtPASISJNo.Text Else ls_SuratJalanNo = txtSupplierSJNo.Text
            'ls_SupplierID = txtSupplierCode.Text

            Using scope As New TransactionScope

                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                    sqlConn.Open()

                    For iLoop = 0 To e.UpdateValues.Count - 1
                        If e.UpdateValues(iLoop).NewValues("QtyBox") = 0 Then
                            Session("E02Msg") = "Qty/Box not found in Part Mapping Master, please check again with PASI!"
                            Exit Sub
                        End If
                        ls_PONo = e.UpdateValues(iLoop).NewValues("PONo").ToString()
                        If e.UpdateValues(iLoop).NewValues("POKanban").ToString() = "YES" Then ls_POKanbanCls = "1" Else ls_POKanbanCls = "0"
                        ls_KanbanNo = e.UpdateValues(iLoop).NewValues("KanbanNo").ToString()
                        ls_PartNo = e.UpdateValues(iLoop).NewValues("PartNo").ToString()
                        ls_UnitCls = e.UpdateValues(iLoop).NewValues("UnitCls").ToString()
                        ls_RecQty = e.UpdateValues(iLoop).NewValues("GoodReceivingQty").ToString()
                        ls_DefQty = e.UpdateValues(iLoop).NewValues("DefectReceivingQty").ToString()
                        ls_DeliveryByPASICls = e.UpdateValues(iLoop).NewValues("DeliveryByPASICls").ToString()
                        ls_SupplierID = e.UpdateValues(iLoop).NewValues("Supplier").ToString()

                        If ls_DeliveryByPASICls = "1" Then
                            ls_Delivery = e.UpdateValues(iLoop).NewValues("PASIDeliveryQty").ToString().Trim()
                            Call clsMsg.DisplayMessage(lblInfo, "7001", clsMessage.MsgType.ErrorMessage)
                            lblInfo.Text = Replace(lblInfo.Text, "%%", "PASI Delivery Qty")
                        Else
                            ls_Delivery = e.UpdateValues(iLoop).NewValues("SupplierDeliveryQty").ToString().Trim()
                            Call clsMsg.DisplayMessage(lblInfo, "7001", clsMessage.MsgType.ErrorMessage)
                            lblInfo.Text = Replace(lblInfo.Text, "%%", "Supplier Delivery Qty")
                        End If

                        If CDbl(ls_Delivery) < (CDbl(ls_RecQty) + CDbl(ls_DefQty)) Then
                            Session("E02Msg") = lblInfo.Text
                            Exit Sub
                        End If

                        ls_SQL = uf_SaveAffiliateDetail(ls_SuratJalanNo, ls_SupplierID, ls_PONo, ls_POKanbanCls, ls_KanbanNo, ls_PartNo, ls_UnitCls, ls_RecQty, ls_DefQty)
                        Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
                        sqlCmd.ExecuteNonQuery()
                        sqlCmd.Dispose()
                    Next iLoop

                End Using
                scope.Complete()
            End Using

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("E02Msg") = lblInfo.Text
        End Try

        grid.SettingsPager.PageSize = grid.VisibleRowCount + 1
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        Session("E02Msg") = lblInfo.Text
                    Else
                        grid.PageIndex = 0

                        grid.JSProperties("cpTotalBox") = uf_CalculateBox()
                    End If

                Case "save"
                    If IsNothing(Session("E02Msg")) Then
                        Call up_GridLoad()
                        Call up_SaveData()
                        Call up_GridLoad()
                        grid.JSProperties("cpTotalBox") = uf_CalculateBox()

                    End If

                Case "delete"
                    Call up_Delete()
                    grid.JSProperties("cpTotalBox") = uf_CalculateBox()
                    Call clsMsg.DisplayMessage(lblInfo, "1003", clsMessage.MsgType.InformationMessage)
                    Session("E02Msg") = lblInfo.Text

                Case "sendtosupplier"
                    Call UpdateExcel(True, Session("AffiliateID"), IIf(txtPASISJNo.Text.Trim = "", txtSupplierSJNo.Text.Trim, txtPASISJNo.Text.Trim), "")
                    Call clsMsg.DisplayMessage(lblInfo, "1010", clsMessage.MsgType.InformationMessage)
                    Session("E02Msg") = lblInfo.Text

                Case "clear"
                    Call up_GridLoadWhenEventChange()

                Case "bc40"
                    Call ExcelBC40()
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("E02Msg") = lblInfo.Text
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
        If (Not IsNothing(Session("E02Msg"))) Then grid.JSProperties("cpMessage") = Session("E02Msg") : Session.Remove("E02Msg")
    End Sub

    Private Sub grid_CustomColumnDisplayText(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewColumnDisplayTextEventArgs) Handles grid.CustomColumnDisplayText
        If (e.Column.FieldName = "QtyBox" Or e.Column.FieldName = "SupplierDeliveryQty" Or _
            e.Column.FieldName = "PASIGoodReceivingQty" Or e.Column.FieldName = "PASIDefectQty" Or _
            e.Column.FieldName = "PASIDeliveryQty" Or e.Column.FieldName = "GoodReceivingQty" Or _
            e.Column.FieldName = "DefectReceivingQty" Or e.Column.FieldName = "RemainingReceivingQty" Or _
            e.Column.FieldName = "ReceivingQtyBox") Then

            Dim ls_Value As String = e.GetFieldValue(e.Column.FieldName)
            If IsNothing(ls_Value) Then ls_Value = "0"
            e.DisplayText = FormatNumber(ls_Value.Trim, 0, TriState.True)
            If ls_Value = "" Then e.DisplayText = "0"
        End If
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If Not (e.DataColumn.FieldName = "GoodReceivingQty" Or _
            e.DataColumn.FieldName = "DefectReceivingQty") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        Else
            e.Cell.BackColor = Color.White
        End If

        If e.DataColumn.FieldName = "RemainingReceivingQty" Then
            If e.GetValue("DeliveryByPASICls") = "1" Then
                'DELIVERY BY PASI (SUPPLIER to PASI to AFFILIATE)
                If CDbl(e.GetValue("PASIDeliveryQty")) > (CDbl(e.GetValue("GoodReceivingQty")) + CDbl(e.GetValue("DefectReceivingQty"))) Then
                    e.Cell.BackColor = Color.Fuchsia
                End If

            ElseIf e.GetValue("DeliveryByPASICls") = "0" Then
                'DIRECT DELIVERY (SUPPLIER to AFFILIATE)
                If CDbl(e.GetValue("SupplierDeliveryQty")) > (CDbl(e.GetValue("GoodReceivingQty")) + CDbl(e.GetValue("DefectReceivingQty"))) Then
                    e.Cell.BackColor = Color.Fuchsia
                End If
            End If
        End If

        If e.DataColumn.FieldName = "GoodReceivingQty" Or e.DataColumn.FieldName = "DefectReceivingQty" Then
            If e.GetValue("IsSaved") = "YES" Then
                e.Cell.BackColor = Color.White
            Else
                e.Cell.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub

    Private Sub btnPrintGR_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrintGR.Click
        If txtPASISJNo.Text <> "" Then
            Session("E02SupplierSJNo") = txtPASISJNo.Text
        Else
            Session("E02SupplierSJNo") = txtSupplierSJNo.Text
        End If

        Dim ls_sql As String = ""
        Session.Remove("REPORT")
        Session.Remove("Query")

        ls_sql = " SELECT  ColNo = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY PONo, KanbanNo )), * " & vbCrLf & _
                  " FROM " & vbCrLf & _
                  " ( " & vbCrLf & _
                  " 	SELECT DISTINCT " & vbCrLf & _
                  " 		ReceiveDate = Format(RAM.ReceiveDate,'yyyy-MM-dd') ,  " & vbCrLf & _
                  " 		RAM.JenisArmada ,  " & vbCrLf & _
                  " 		RAM.NoPol ,  " & vbCrLf & _
                  " 		DeliveryTo = MDP.DeliveryLocationName ,  " & vbCrLf & _
                  " 		PASISJNo = RAM.SuratJalanNo,  " & vbCrLf & _
                  " 		PASIDeliveryDate = DPM.DeliveryDate ,  " & vbCrLf & _
                  " 		RAM.DriverName ,  "

        ls_sql = ls_sql + " 		TotalBox = (SELECT SUM(ABC.RecQty / ISNULL(POD2.POQtyBox,DEF.QtyBox)) FROM ReceiveAffiliate_Detail ABC  " & vbCrLf & _
                          " 					LEFT JOIN PO_Detail POD2 ON ABC.PONo = POD2.PONo And ABC.AffiliateID = POD2.AffiliateID And ABC.SupplierID = POD2.SupplierID And ABC.PartNo = POD2.PartNo " & vbCrLf & _
                          " 					LEFT JOIN MS_PartMapping DEF ON DEF.AffiliateID = ABC.AffiliateID and DEF.SupplierID = ABC.SupplierID and DEF.PartNo = ABC.PartNo " & vbCrLf & _
                          " 					WHERE ABC.SuratJalanNo = '" & Session("E02SupplierSJNo") & "' and ABC.AffiliateID = '" & Session("AffiliateID") & "'),  " & vbCrLf & _
                          " 		RAD.PONo,  " & vbCrLf & _
                          " 		POKanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '1' THEN 'YES' ELSE 'NO' END ,  " & vbCrLf & _
                          " 		RAD.KanbanNo ,  " & vbCrLf & _
                          " 		RAD.PartNo ,  " & vbCrLf & _
                          " 		MP.PartName ,  " & vbCrLf & _
                          " 		UOM = MU.Description,  " & vbCrLf & _
                          " 		QtyBox = ISNULL(POD.POQtyBox,MPM.QtyBox),  " & vbCrLf & _
                          " 		PerformanceCls = ISNULL(RTRIM(MPC.Description), '-'), "

        ls_sql = ls_sql + " 		PASIDeliveryQty = DPD.DOQty, " & vbCrLf & _
                          " 		AffiliateRecQty = RAD.RecQty, " & vbCrLf & _
                          " 		AffiliateDefQty = RAD.DefectQty, " & vbCrLf & _
                          " 		AffiliateRemQty = DPD.DOQty - (RAD.RecQty  + RAD.DefectQty), " & vbCrLf & _
                          " 		AffiliateRecBox = RAD.RecQty / ISNULL(POD.POQtyBox,MPM.QtyBox) " & vbCrLf & _
                          " 	FROM ReceiveAffiliate_Master RAM " & vbCrLf & _
                          " 	LEFT JOIN ReceiveAffiliate_Detail RAD ON RAM.SuratJalanNo = RAD.SuratJalanNo and RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          "     LEFT JOIN PO_Detail POD ON RAD.PONo = POD.PONo And RAD.AffiliateID = POD.AffiliateID And RAD.SupplierID = POD.SupplierID And RAD.PartNo = POD.PartNo " & vbCrLf & _
                          " 	LEFT JOIN DOPASI_Master DPM ON DPM.AffiliateID = RAM.AffiliateID and DPM.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 	LEFT JOIN DOPASI_Detail DPD ON DPD.AffiliateID = RAD.AffiliateID and DPD.SupplierID = RAD.SupplierID and DPD.PartNo = RAD.PartNo AND RAD.SuratJalanNo = DPD.SuratJalanNo And RAD.KanbanNo = DPD.KanbanNo" & vbCrLf & _
                          " 	LEFT JOIN MS_DeliveryPlace MDP ON MDP.AffiliateID = RAM.AffiliateID  " & vbCrLf & _
                          " 	LEFT JOIN MS_Parts MP ON MP.PartNo = RAD.PartNo "

        ls_sql = ls_sql + " 	LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                          " 	LEFT JOIN MS_PartMapping MPM ON MPM.AffiliateID = RAD.AffiliateID and MPM.SupplierID = RAD.SupplierID AND MPM.PartNo = RAD.PartNo " & vbCrLf & _
                          " 	LEFT JOIN MS_PerformanceCls MPC ON MPC.PerformanceCls = RAM.PerformanceCls  " & vbCrLf & _
                          " 	WHERE RAM.AffiliateID = '" & Session("AffiliateID") & "' and RAM.SuratJalanNo = '" & Session("E02SupplierSJNo") & "' " & vbCrLf & _
                          " )XYZ " & vbCrLf & _
                          "  "

        Session("REPORT") = "GR"
        Session("Query") = ls_sql
        Response.Redirect("~/Receiving/GoodReceivingReportCR.aspx")
    End Sub
#End Region

End Class