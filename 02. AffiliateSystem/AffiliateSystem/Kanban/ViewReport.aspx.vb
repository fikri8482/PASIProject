Imports System.Data.SqlClient

Public Class ViewReport
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim paramDT1 As Date
    Dim paramDT2 As Date
    Dim paramSupplier As String

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
#End Region


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ds As New DataSet
        Dim ls_sql As String = ""
        Dim ls_sparator As String = "|"

        If Session("KCR-ReportCode") = "KanbanCard" Then
            Dim Report As New KanbanCard2

            'ls_sql = " SELECT  Rtrim(KM.KanbanNo) AS kanbanNo ,  " & vbCrLf & _
            '      "  Rtrim(KM.SupplierID) AS SupplierID ,  " & vbCrLf & _
            '      "  Rtrim(MSS.SupplierName) AS SupplierName ,  KD.PartNo AS PartNo ,  " & vbCrLf & _
            '      "  Rtrim(MSP.PartName) AS PartName ,  " & vbCrLf & _
            '      "  KD.KanbanQty Qty ,  " & vbCrLf & _
            '      "  Rtrim(KM.AffiliateID) AS Cust ,  " & vbCrLf & _
            '      "  DeliveryDate = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(KM.Kanbandate,'')), 112) ,  " & vbCrLf & _
            '      "  CONVERT(CHAR(5), KM.KanbanTime) AS TIME ,  " & vbCrLf & _
            '      "  Rtrim(ISNULL(MDP.DeliveryLocationName,'')) Location ,  " & vbCrLf & _
            '      "  Rtrim(KD.PONo) PONo ,  " & vbCrLf & _
            '      "  Barcode = RTRIM(Barcode) ,  "

            'ls_sql = ls_sql + "  QtyBox = Ceiling(MSP.QtyBox),startno = convert(numeric,seqno), total = Ceiling(KanbanQty/MSP.QtyBox)  " & vbCrLf & _
            '                  "  FROM    dbo.Kanban_Master KM  " & vbCrLf & _
            '                  "  LEFT JOIN dbo.Kanban_Detail KD ON KM.AffiliateID = KD.AffiliateID   and kanbanqty <> 0 " & vbCrLf & _
            '                  "  AND KM.KanbanNo = KD.KanbanNo  " & vbCrLf & _
            '                  "  AND KM.SupplierID = KD.SupplierID  " & vbCrLf & _
            '                  "   LEFT JOIN dbo.MS_DeliveryPlace MDP ON MDP.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
            '                  "  LEFT JOIN dbo.MS_Supplier MSS ON MSS.SupplierID = KM.SupplierID  " & vbCrLf & _
            '                  "  LEFT JOIN dbo.MS_Parts MSP ON MSP.PartNo = KD.PartNo  " & vbCrLf & _
            '                  "  LEFT JOIN Kanban_Barcode KB ON KB.PONO = KD.PONO " & vbCrLf & _
            '                  "                              AND KB.KanbanNo = KD.Kanbanno   " & vbCrLf & _
            '                  "                              AND KB.AffiliateID = KD.AffiliateID  "

            'ls_sql = ls_sql + "                              AND KB.SupplierID = KD.SupplierID  " & vbCrLf & _
            '                  "                              AND KB.DeliveryLocationCode = KD.DeliveryLocationCode " & vbCrLf & _
            '                  "                              AND KB.PartNo = KD.PartNo " & vbCrLf & _
            '                  "  WHERE   KD.AffiliateID = '" & Session("AffiliateID") & "'  " & vbCrLf

            'If Session("KCR-KanbanDate") <> "" And Session("KCR-Form") = "KanbanCreate" Then
            '    ls_sql = ls_sql + " AND convert(char(11), convert(datetime, KanbanDate),106) = '" & Session("KCR-KanbanDate") & "'  "
            'ElseIf Session("KCR-KanbanDate") <> "" Then
            '    ls_sql = ls_sql + " AND convert(char(11), convert(datetime, KanbanDate),106) IN (" & Session("KCR-KanbanDate") & ")  "
            'End If

            'If Session("KCR-SupplierCode") <> "" And Session("KCR-Form") = "KanbanCreate" Then
            '    ls_sql = ls_sql + " AND KD.SupplierID  = '" & Session("KCR-SupplierCode") & "'  " & vbCrLf
            'ElseIf Session("KCR-SupplierCode") <> "" Then
            '    ls_sql = ls_sql + " AND KD.SupplierID IN (" & Session("KCR-SupplierCode") & ")  " & vbCrLf
            'End If

            'If Session("KCR-DeliveryLocation") <> "" Then
            '    ls_sql = ls_sql + " AND KD.DeliveryLocationCode = '" & Session("KCR-DeliveryLocation") & "'" & vbCrLf
            'End If

            'ls_sql = ls_sql + " order by kanbanno, partno, startno "

            ls_sql = "  SELECT * FROM (SELECT DISTINCT ETACust = Rtrim(CONVERT(CHAR(5), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) , " & vbCrLf & _
                  "         ETACustYear = '/' " & vbCrLf & _
                  "         + CONVERT(CHAR(4), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 120) , " & vbCrLf & _
                  "         ETACustTime = CONVERT(CHAR(5), KM.KanbanTime) , " & vbCrLf & _
                  "         ETAPasi = CONVERT(CHAR(5), CONVERT(DATETIME, ISNULL(MEP.ETDPASI, '')), 103) , " & vbCrLf & _
                  "         ETAPasiYear = '/' " & vbCrLf & _
                  "         + CONVERT(CHAR(4), CONVERT(DATETIME, ISNULL(MEP.ETDPasi, '')), 120) , " & vbCrLf & _
                  "         ETAPasiTime = '12:00' , " & vbCrLf & _
                  "         KanbanNo = RTRIM(KM.KanbanNo) , " & vbCrLf & _
                  "         SeqStart = Rtrim(CONVERT(NUMERIC, isnull(seqnoStart,0))) , " & vbCrLf & _
                  "         SeqEnd = Rtrim(CONVERT(NUMERIC, isnull(seqnoEnd,0))) , " & vbCrLf

            ls_sql = ls_sql + "         PartNo1 = LEFT(Rtrim(KD.PartNo),2) , " & vbCrLf &
                              "         PartNo2 = SUBSTRING(Rtrim(KD.PartNo),3,9) , " & vbCrLf &
                              "         PartNo3 = SUBSTRING(Rtrim(KD.PartNo),10,10) , " & vbCrLf &
                              "         PartName = Rtrim(MP.PartName) , " & vbCrLf &
                              "         PartCMCode = Rtrim(ISNULL(MP.PartCarMaker, '')) , " & vbCrLf &
                              "         PartCMName = Rtrim(ISNULL(MP.PartGroupName, '')) , " & vbCrLf &
                              "         Qty = Replace(Rtrim(ML.QtyBox),'.00','') , " & vbCrLf &
                              "         BoxNo = Rtrim(isnull(KB.BoxNo,'')) , " & vbCrLf &
                              "         Cust = RTRIM(KM.AffiliateID) , " & vbCrLf &
                              "         AFFCode = Rtrim(ISNULL(MA.AffiliateCode, '')) , " & vbCrLf &
                              "         Location = Rtrim(ISNULL(ML.LocationID, '')) , " & vbCrLf &
                              "         SupplierID = RTRIM(KM.SupplierID) + '#1' , " & vbCrLf &
                              "         SupplierCode = Rtrim(ISNULL(MS.SupplierCode, '')) , " & vbCrLf

            ls_sql = ls_sql + "         Barcode = RTRIM(KB.barcode2) " & vbCrLf & _
                              "  FROM   dbo.Kanban_Master KM " & vbCrLf & _
                              "         LEFT JOIN dbo.Kanban_Detail KD ON KM.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                                           AND kanbanqty <> 0 " & vbCrLf & _
                              "                                           AND KM.KanbanNo = KD.KanbanNo " & vbCrLf & _
                              "                                           AND KM.SupplierID = KD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = KM.SupplierID " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo " & vbCrLf & _
                              "         INNER JOIN Kanban_Barcode KB ON KB.PONO = KD.PONO " & vbCrLf & _
                              "                                        AND KB.KanbanNo = KD.Kanbanno " & vbCrLf & _
                              "                                        AND KB.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                                        AND KB.DeliveryLocationCode = KD.DeliveryLocationCode " & vbCrLf

            ls_sql = ls_sql + "                                        AND KB.SupplierID = KD.SupplierID " & vbCrLf & _
                              "                                        AND KB.PartNo = KD.PartNo " & vbCrLf & _
                              "         LEFT JOIN MS_PartMapping ML ON KD.AffiliateID = ML.affiliateID AND ML.SupplierID = KD.SupplierID AND KD.PartNo = ML.PartNo" & vbCrLf & _
                              "         --LEFT JOIN ms_dock MD ON MD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = KM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN ms_Etd_pasi MEP ON MEP.affiliateID = KM.AffiliateID " & vbCrLf & _
                              "                                      AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(KM.Kanbandate,'')), 112) = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MEP.ETAAffiliate,'')), 112) " & vbCrLf & _
                              "  WHERE   KD.AffiliateID = '" & Session("AffiliateID") & "'  " & vbCrLf

            If Session("KCR-KanbanDate") <> "" And Session("KCR-Form") = "KanbanCreate" Then
                ls_sql = ls_sql + " AND convert(char(11), convert(datetime, KanbanDate),106) = '" & Session("KCR-KanbanDate") & "'  "
            ElseIf Session("KCR-KanbanDate") <> "" Then
                ls_sql = ls_sql + " AND convert(char(11), convert(datetime, KanbanDate),106) IN (" & Session("KCR-KanbanDate") & ")  "
            End If

            If Session("KCR-SupplierCode") <> "" And Session("KCR-Form") = "KanbanCreate" Then
                ls_sql = ls_sql + " AND KD.SupplierID  = '" & Session("KCR-SupplierCode") & "'  " & vbCrLf
            ElseIf Session("KCR-SupplierCode") <> "" Then
                ls_sql = ls_sql + " AND KD.SupplierID IN (" & Session("KCR-SupplierCode") & ")  " & vbCrLf
            End If

            If Session("KCR-DeliveryLocation") <> "" And Session("KCR-Form") = "KanbanCreate" Then
                ls_sql = ls_sql + " AND KD.DeliveryLocationCode  = '" & Session("KCR-DeliveryLocation") & "'  " & vbCrLf
            ElseIf Session("KCR-SupplierCode") <> "" Then
                ls_sql = ls_sql + " AND KD.DeliveryLocationCode IN (" & Session("KCR-DeliveryLocation") & ")  " & vbCrLf
            End If

            'If Session("KCR-DeliveryLocation") <> "" Then
            '    ls_sql = ls_sql + " AND KD.DeliveryLocationCode IN (" & Session("KCR-DeliveryLocation") & ")" & vbCrLf
            'End If

            If Session("KCR-kanbanno") <> "" Then
                ls_sql = ls_sql + " AND KD.kanbanno in (" & Session("KCR-kanbanno") & ")" & vbCrLf
            End If

            'ls_sql = ls_sql + " order by kanbanno, KD.partno, seqNostart "
            ls_sql = ls_sql + " ) XYZ order by kanbanno, BoxNo "

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
            ASPxDocumentViewer1.DataBind()
        ElseIf Session("KCR-ReportCode") = "KanbanCycle" Then
            Dim report As New RptKanbanCycle
            ls_sql = "Select *,NO =ROW_NUMBER() OVER(PARTITION BY KanbanNo ORDER BY kanbanno,partno DESC)  from (" & vbCrLf
            ls_sql = ls_sql + " SELECT  DISTINCT KM.KanbanNo , " & vbCrLf & _
                             "         SupplierCode = KM.SupplierID, " & vbCrLf & _
                              "         MSS.SupplierName, " & vbCrLf & _
                              "         KanbanDate = convert(char(11), convert(datetime, KM.EntryDate),106) , " & vbCrLf & _
                              "         KanbanTime = convert(char(5),convert(datetime,KM.EntryDate),114), " & vbCrLf & _
                              "         KD.PartNo, " & vbCrLf & _
                              "         PONo, " & vbCrLf & _
                              "         PartName = Rtrim(MSP.PartName), " & vbCrLf & _
                              "         MOQ, " & vbCrLf & _
                              "         Unit = Description, " & vbCrLf & _
                              "         KanbanQty, " & vbCrLf & _
                              "         QtyBox, " & vbCrLf & _
                              "         ETA = convert(char(11), convert(datetime, ES.ETDSupplier),106) , " & vbCrLf & _
                              "         ETAT = '-' " & vbCrLf & _
                              "         --,NO =ROW_NUMBER() OVER(PARTITION BY KD.KanbanNo ORDER BY KD.kanbanno,KD.partno DESC) " & vbCrLf & _
                              "         , AFF = isnull(MA.AffiliateName,''), AFFADD = Rtrim(isnull(MA.Address,'')), AFFADD2 = Rtrim(isnull(MA.City,'')), " & vbCrLf & _
                              "          AFFADD3 = Rtrim(isnull(MA.Phone1,'')),AFFADD4 = Rtrim(isnull(MA.Fax,'')) " & vbCrLf

            ls_sql = ls_sql + " FROM    dbo.Kanban_Master KM " & vbCrLf & _
                              "         INNER JOIN dbo.Kanban_Detail KD ON KM.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                                            AND KM.KanbanNo = KD.KanbanNo AND KM.DeliveryLocationCode = KD.DeliveryLocationCode" & vbCrLf & _
                              "                                            AND KM.SupplierID = KD.SupplierID and kanbanqty <> 0" & vbCrLf & _
                              "         LEFT JOIN MS_ETD_PASI MEP ON KM.AffiliateID = MEP.AffiliateID AND CONVERT(CHAR(11),isnull(MEP.ETAAFFILIATE,''),106) = CONVERT(CHAR(11),isnull(KM.KanbanDate,''),106) " & vbCrLf & _
                              "         LEFT JOIN MS_ETD_Supplier_Pasi ES ON MEP.ETDPASI = ES.ETAPASI " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Supplier MSS ON MSS.SupplierID = KD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = KM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Parts MSP ON MSP.PartNo = KD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID  " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_UnitCls MUC ON MUC.UnitCls = MSP.UnitCls " & vbCrLf

            If Session("KCR-Form") = "KanbanCreate" Then
                ls_sql = ls_sql + " WHERE convert(char(11), convert(datetime, KanbanDate),106) = '" & Session("KCR-KanbanDate") & "' " & vbCrLf & _
                                  " AND KD.SupplierID = '" & Session("KCR-SupplierCode") & "' and KanbanQty <> 0"
            Else
                ls_sql = ls_sql + " WHERE convert(char(11), convert(datetime, KanbanDate),106) IN (" & Session("KCR-KanbanDate") & ") " & vbCrLf & _
                                  " AND KD.SupplierID IN (" & Session("KCR-SupplierCode") & ") and kanbanQty <> 0"
            End If

            If Session("KCR-DeliveryLocation") <> "" And Session("KCR-Form") = "KanbanCreate" Then
                ls_sql = ls_sql + " AND KD.DeliveryLocationCode  = '" & Session("KCR-DeliveryLocation") & "'  " & vbCrLf
            ElseIf Session("KCR-SupplierCode") <> "" Then
                ls_sql = ls_sql + " AND KD.DeliveryLocationCode IN (" & Session("KCR-DeliveryLocation") & ")  " & vbCrLf
            End If


            If Session("KCR-Form") = "KanbanList" Then
                If (Session("KCR-kanbanno") <> "" Or Session("KCR-kanbanno") <> clsGlobal.gs_All) Then
                    ls_sql = ls_sql + " AND KD.kanbanno IN (" & Session("KCR-kanbanno") & ")" & vbCrLf
                End If
            End If
            If Session("KCR-Form") = "KanbanCreate" Then
                If (Session("KCR-kanbanno") <> "" Or Session("KCR-kanbanno") <> clsGlobal.gs_All) Then
                    ls_sql = ls_sql + " AND KD.kanbanno = '" & Session("KCR-kanbanno") & "'" & vbCrLf
                End If
            End If
            ls_sql = ls_sql + ")x"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = report
            ASPxDocumentViewer1.DataBind()
        End If

    End Sub

    Protected Sub btnsubmenu_Click(sender As Object, e As EventArgs) Handles btnsubmenu.Click
        If Session("KCR-Form") = "KanbanCreate" Then
            Response.Redirect("~/kanban/KanbanCreate.aspx?prm=" + Format(Session("tmp_kanbandate"), "dd MMM yyyy") + "|" + Session("tmp_suppID") + "|" + Session("tmp_suppname") + "|" + Session("tmp_affentrydate") + "|" + Session("tmp_affentryname") + "|" + Session("tmp_affappdate") + "|" + Session("tmp_affappname") + "|" + Session("tmp_suppappdate") + "|" + Session("tmp_suppappname") + "|" + Session("tmp_dt1") + "|" + Session("tmp_dt2") + "|" + Session("tmp_cbosupplier") + "|" + Session("tmp_cbosupplierName") + "|" + Session("tmp_cbolocation") + "|" + Session("tmp_Location") + "|" + Session("tmp_cbolocation1") + "|" + Session("tmp_Location1") + "|" + Session("tmp_Kanbanno"))
        ElseIf Session("KCR-Form") = "KanbanList" Then
            Session("KCR-Load") = "NO"
            Session("KCR-ReportCode") = ""
            Response.Redirect("~/kanban/KanbanList.aspx")
        End If
    End Sub
End Class