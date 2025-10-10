Imports System.Data.SqlClient

Public Class clsPOExportMonthly
    Public Property SupplierID As String
    Public Property AffiliateID As String
    Public Property PartNo As String
    Public Property Period As Date
    Public Property POMonthly As String
    Public Property DeliveryLocation As String

    Public Property Order1 As String
    Public Property Order2 As String
    Public Property Order3 As String
    Public Property Order4 As String
    Public Property Order5 As String

    Public Property Vendor1 As Date
    Public Property Vendor2 As Date
    Public Property Vendorr3 As Date
    Public Property Vendor4 As Date
    Public Property Vendor5 As Date

    Public Property ETDPort1 As Date
    Public Property ETDPort2 As Date
    Public Property ETDPort3 As Date
    Public Property ETDPort4 As Date
    Public Property ETDPort5 As Date

    Public Property ETAPort1 As Date
    Public Property ETAPort2 As Date
    Public Property ETAPort3 As Date
    Public Property ETAPort4 As Date
    Public Property ETAPort5 As Date

    Public Property ETAFactory1 As Date
    Public Property ETAFactory2 As Date
    Public Property ETAFactory3 As Date
    Public Property ETAFactory4 As Date
    Public Property ETAFactory5 As Date



    Public Shared Function GetTableMonthly(ByVal pAffiliateID As String, ByVal pPeriod As Date, ByVal pPOMonthly As String, ByVal pDeliveryLocation As String, ByVal pOrder1 As String, ByVal pOrder2 As String, ByVal pOrder3 As String, ByVal pOrder4 As String, ByVal pOrder5 As String, ByVal pVendor1 As String, ByVal pVendor2 As String, ByVal pVendor3 As String, ByVal pVendor4 As String, ByVal pVendor5 As String, ByVal pETDPort1 As String, ByVal pETDPort2 As String, ByVal pETDPort3 As String, ByVal pETDPort4 As String, ByVal pETDPort5 As String, ByVal pETAPort1 As String, ByVal pETAPort2 As String, ByVal pETAPort3 As String, ByVal pETAPort4 As String, ByVal pETAPort5 As String, ByVal pETAFactory1 As String, ByVal pETAFactory2 As String, ByVal pETAFactory3 As String, ByVal pETAFactory4 As String, ByVal pETAFactory5 As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        'If pAffiliateID.Trim <> "" Then
        '    pWhere = pWhere + " and e.AffiliateID like '%" & pAffiliateID & "%'"
        'End If

        'If pAffiliateName.Trim <> "" Then
        '    pWhere = pWhere + " and AffiliateName like '%" & pAffiliateName & "%'"
        'End If

        'If pCommercial.Trim <> "" Then
        '    pWhere = pWhere + " and CommercialCls like '%" & pCommercial & "%'"
        'End If

        'If pPeriod <> isd"" Then
        '    pWhere = pWhere + " and Period like '%" & pPeriod & "%'"
        'End If

        'If pPOMonthly <> "" Then
        '    pWhere = pWhere + " and EmergencyCls like '%" & pPOMonthly & "%'"
        'End If

        'If pDeliveryLocation <> "" Then
        '    pWhere = pWhere + " and ForwarderID like '%" & pDeliveryLocation & "%'"
        'End If

        'If pOrder1 <> "" Then
        '    pWhere = pWhere + " and OrderNo1 like '%" & pOrder1 & "%'"
        'End If

        'If pOrder2 <> "" Then
        '    pWhere = pWhere + " and OrderNo2 like '%" & pOrder2 & "%'"
        'End If

        'If pOrder3 <> "" Then
        '    pWhere = pWhere + " and OrderNo3 like '%" & pOrder3 & "%'"
        'End If

        'If pOrder4 <> "" Then
        '    pWhere = pWhere + " and OrderNo4 like '%" & pOrder4 & "%'"
        'End If

        'If pOrder5 <> "" Then
        '    pWhere = pWhere + " and OrderNo5 like '%" & pOrder5 & "%'"
        'End If

        'If pVendor1 <> "" Then
        '    pWhere = pWhere + " and ETDVendor1 like '%" & pVendor1 & "%'"
        'End If

        'If pVendor2 <> "" Then
        '    pWhere = pWhere + " and ETDVendor2 like '%" & pVendor2 & "%'"
        'End If

        'If pVendor3 <> "" Then
        '    pWhere = pWhere + " and ETDVendor3 like '%" & pVendor3 & "%'"
        'End If

        'If pVendor4 <> "" Then
        '    pWhere = pWhere + " and ETDVendor4 like '%" & pVendor4 & "%'"
        'End If

        'If pVendor5 <> "" Then
        '    pWhere = pWhere + " and ETDVendor5 like '%" & pVendor5 & "%'"
        'End If

        'If pETDPort1 <> "" Then
        '    pWhere = pWhere + " and ETDPort1 like '%" & pETDPort1 & "%'"
        'End If

        'If pETDPort2 <> "" Then
        '    pWhere = pWhere + " and ETDPort2 like '%" & pETDPort2 & "%'"
        'End If

        'If pETDPort3 <> "" Then
        '    pWhere = pWhere + " and ETDPort3 like '%" & pETDPort3 & "%'"
        'End If

        'If pETDPort4 <> "" Then
        '    pWhere = pWhere + " and ETDPort4 like '%" & pETDPort4 & "%'"
        'End If

        'If pETDPort5 <> "" Then
        '    pWhere = pWhere + " and ETDPort5 like '%" & pETDPort5 & "%'"
        'End If

        'If pETAPort1 <> "" Then
        '    pWhere = pWhere + " and ETAPort1 like '%" & pETAPort1 & "%'"
        'End If

        'If pETAPort2 <> "" Then
        '    pWhere = pWhere + " and ETAPort2 like '%" & pETAPort2 & "%'"
        'End If

        'If pETAPort3 <> "" Then
        '    pWhere = pWhere + " and ETAPort3 like '%" & pETAPort3 & "%'"
        'End If

        'If pETAPort4 <> "" Then
        '    pWhere = pWhere + " and ETAPort4 like '%" & pETAPort4 & "%'"
        'End If

        'If pETAPort5 <> "" Then
        '    pWhere = pWhere + " and ETAPort5 like '%" & pETAPort5 & "%'"
        'End If

        'If pETAFactory1 <> "" Then
        '    pWhere = pWhere + " and ETAFactory1 like '%" & pETAFactory1 & "%'"
        'End If

        'If pETAFactory2 <> "" Then
        '    pWhere = pWhere + " and ETAFactory2 like '%" & pETAFactory2 & "%'"
        'End If

        'If pETAFactory3 <> "" Then
        '    pWhere = pWhere + " and ETAFactory3 like '%" & pETAFactory3 & "%'"
        'End If

        'If pETAFactory4 <> "" Then
        '    pWhere = pWhere + " and ETAFactory4 like '%" & pETAFactory4 & "%'"
        'End If

        'If pETAFactory5 <> "" Then
        '    pWhere = pWhere + " and ETAFactory5 like '%" & pETAFactory5 & "%'"
        'End If

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""

                ls_sql = " select " & vbCrLf & _
                  " 	row_number() over (order by a.AffiliateID, a.SupplierID ) NoUrut,  " & vbCrLf & _
                  "  	RTRIM(a.PartNo)PartNo,  " & vbCrLf & _
                  "  	c.Description,  " & vbCrLf & _
                  "  	'0' Week1,   " & vbCrLf & _
                  "  	'0' Week2,   " & vbCrLf & _
                  "  	'0' Week3,   "

                ls_sql = ls_sql + "  	'0' Week4,   " & vbCrLf & _
                                  "  	'0' Week5,   " & vbCrLf & _
                                  "      PreviousForecast = isnull((select qty from MS_Forecast MF where MF.PartNo = d.PartNo and d.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,-1,'" & Format(pPeriod, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,-1,'" & Format(pPeriod, "yyyy-MM-dd") & "'))),1), " & vbCrLf & _
                              " 	Forecast1 = isnull((select qty from MS_Forecast MF where MF.PartNo = d.PartNo and d.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(pPeriod, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(pPeriod, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     Forecast2 = isnull((select qty from MS_Forecast MF where MF.PartNo = d.PartNo and d.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(pPeriod, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(pPeriod, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     Forecast3 = isnull((select qty from MS_Forecast MF where MF.PartNo = d.PartNo and d.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(pPeriod, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(pPeriod, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                                  "  	e.ShipCls,   " & vbCrLf & _
                                  "  	e.CommercialCls,   " & vbCrLf & _
                                  "  	e.ForwarderID,   " & vbCrLf & _
                                  "  	e.Period,   " & vbCrLf & _
                                  "  e.EmergencyCls, " & vbCrLf & _
                                  " e.OrderNo1, " & vbCrLf & _
                " e.ETDVendor1, " & vbCrLf & _
                " e.ETDPort1, " & vbCrLf & _
                " e.ETAPort1, " & vbCrLf & _
                " e.ETAFactory1, " & vbCrLf & _
                " e.OrderNo2, " & vbCrLf & _
                " e.ETDVendor2, " & vbCrLf & _
                " e.ETDPort2, " & vbCrLf & _
                " e.ETAPort2, " & vbCrLf & _
                " e.ETAFactory2, " & vbCrLf & _
                " e.OrderNo3, " & vbCrLf & _
                " e.ETDVendor3, " & vbCrLf & _
                " e.ETDPort3, " & vbCrLf & _
                " e.ETAPort3, " & vbCrLf & _
                " e.ETAFactory3, " & vbCrLf & _
                " e.OrderNo4, " & vbCrLf & _
                " e.ETDVendor4, " & vbCrLf & _
                " e.ETDPort4, " & vbCrLf & _
                " e.ETAPort4, " & vbCrLf & _
                " e.ETAFactory4, " & vbCrLf & _
                " e.OrderNo5, " & vbCrLf & _
                " e.ETDVendor5, " & vbCrLf & _
                " e.ETDPort5, " & vbCrLf & _
                " e.ETAPort5, " & vbCrLf & _
                " e.ETAFactory5, " & vbCrLf & _
                                  "  	RTRIM(a.AffiliateID)AffiliateID,  " & vbCrLf & _
                                  "  	RTRIM(a.SupplierID)SupplierID  " & vbCrLf & _
                                  "  from MS_PartMapping a  " & vbCrLf & _
                                  "  INNER join MS_Parts b on a.PartNo = b.PartNo  " & vbCrLf & _
                                  "  LEFT join MS_UnitCls c on c.UnitCls = b.UnitCls  " & vbCrLf & _
                                  "  left join PO_Detail_Export d on a.PartNo = d.PartNo AND A.AffiliateID = D.AffiliateID AND A.SupplierID = D.SupplierID  " & vbCrLf & _
                                  "  Left join PO_Master_Export e on d.PONo = e.PONo AND d.AffiliateID = e.AffiliateID AND d.SupplierID = e.SupplierID  " & vbCrLf & _
                                  "  where a.AffiliateID = a.AffiliateID AND NOT EXISTS " & vbCrLf & _
                                  "  ( " & vbCrLf & _
                                  " SELECT * FROM  PO_Detail_Export X WHERE X.PONo = e.PONo and 'A' = 'A' " & pWhere & " " & vbCrLf & _
                                  "  ) "
                ' and a.SupplierID = '" & cboSupplier.Text.Trim & "'
                ls_sql = ls_sql + "  union all " & vbCrLf & _
                                  "  select " & vbCrLf & _
                                  " 	row_number() over (order by e.AffiliateID, e.SupplierID ) NoUrut,  " & vbCrLf & _
                                  "  	RTRIM(B.PartNo)PartNo,  " & vbCrLf & _
                                  "  	d.Description,  " & vbCrLf & _
                                  "  	B.Week1,   " & vbCrLf & _
                                  "  	B.Week2,   " & vbCrLf & _
                                  "  	B.Week3,   "

                ls_sql = ls_sql + "  	B.Week4,   " & vbCrLf & _
                                  "  	B.Week5,   " & vbCrLf & _
                                  "     PreviousForecast = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,-1,'" & Format(pPeriod, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,-1,'" & Format(pPeriod, "yyyy-MM-dd") & "'))),1), " & vbCrLf & _
                              " 	Forecast1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(pPeriod, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(pPeriod, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     Forecast2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(pPeriod, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(pPeriod, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     Forecast3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(pPeriod, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(pPeriod, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                                 "  	e.ShipCls,   " & vbCrLf & _
                                  "  	e.CommercialCls,   " & vbCrLf & _
                                  "  	e.ForwarderID,   " & vbCrLf & _
                                  "  	e.Period,   " & vbCrLf & _
                                  "  e.EmergencyCls, " & vbCrLf & _
                                  " e.OrderNo1, " & vbCrLf & _
                " e.ETDVendor1, " & vbCrLf & _
                " e.ETDPort1, " & vbCrLf & _
                " e.ETAPort1, " & vbCrLf & _
                " e.ETAFactory1, " & vbCrLf & _
                " e.OrderNo2, " & vbCrLf & _
                " e.ETDVendor2, " & vbCrLf & _
                " e.ETDPort2, " & vbCrLf & _
                " e.ETAPort2, " & vbCrLf & _
                " e.ETAFactory2, " & vbCrLf & _
                " e.OrderNo3, " & vbCrLf & _
                " e.ETDVendor3, " & vbCrLf & _
                " e.ETDPort3, " & vbCrLf & _
                " e.ETAPort3, " & vbCrLf & _
                " e.ETAFactory3, " & vbCrLf & _
                " e.OrderNo4, " & vbCrLf & _
                " e.ETDVendor4, " & vbCrLf & _
                " e.ETDPort4, " & vbCrLf & _
                " e.ETAPort4, " & vbCrLf & _
                " e.ETAFactory4, " & vbCrLf & _
                " e.OrderNo5, " & vbCrLf & _
                " e.ETDVendor5, " & vbCrLf & _
                " e.ETDPort5, " & vbCrLf & _
                " e.ETAPort5, " & vbCrLf & _
                " e.ETAFactory5, " & vbCrLf & _
                                  "  	RTRIM(e.AffiliateID)AffiliateID,  " & vbCrLf & _
                                  "  	RTRIM(e.SupplierID)SupplierID  " & vbCrLf & _
                                  " from PO_Master_Export e  " & vbCrLf & _
                                  "  INNER join PO_Detail_Export b on e.PONo = b.PONo AND e.AffiliateID = B.AffiliateID AND e.SupplierID = B.SupplierID  " & vbCrLf & _
                                  "  LEFT join MS_Parts c on c.PartNo = B.PartNo  " & vbCrLf & _
                                  "  LEFT join MS_UnitCls d on d.UnitCls = c.UnitCls  " & vbCrLf & _
                                  "  where 'A' = 'A' " & pWhere & " " & vbCrLf
                '          ls_sql = " SELECT RTRIM([PONo] " & vbCrLf & _
                '" ,[AffiliateID] " & vbCrLf & _
                '" ,[SupplierID] " & vbCrLf & _
                '" ,[ForwarderID]" & vbCrLf & _
                '" ,[Period] " & vbCrLf & _
                '" ,[CommercialCls] " & vbCrLf & _
                '" ,[EmergencyCls] " & vbCrLf & _
                '" ,[ShipCls] " & vbCrLf & _
                '" ,[OrderNo1] " & vbCrLf & _
                '" ,[ETDVendor1] " & vbCrLf & _
                '" ,[ETDPort1] " & vbCrLf & _
                '" ,[ETAPort1] " & vbCrLf & _
                '" ,[ETAFactory1] " & vbCrLf & _
                '" ,[OrderNo2] " & vbCrLf & _
                '" ,[ETDVendor2] " & vbCrLf & _
                '" ,[ETDPort2] " & vbCrLf & _
                '" ,[ETAPort2] " & vbCrLf & _
                '" ,[ETAFactory2] " & vbCrLf & _
                '" ,[OrderNo3] " & vbCrLf & _
                '" ,[ETDVendor3] " & vbCrLf & _
                '" ,[ETDPort3] " & vbCrLf & _
                '" ,[ETAPort3] " & vbCrLf & _
                '" ,[ETAFactory3] " & vbCrLf & _
                '" ,[OrderNo4] " & vbCrLf & _
                '" ,[ETDVendor4] " & vbCrLf & _
                '" ,[ETDPort4] " & vbCrLf & _
                '" ,[ETAPort4] " & vbCrLf & _
                '" ,[ETAFactory4] " & vbCrLf & _
                '" ,[OrderNo5] " & vbCrLf & _
                '" ,[ETDVendor5] " & vbCrLf & _
                '" ,[ETDPort5] " & vbCrLf & _
                '" ,[ETAPort5] " & vbCrLf & _
                '" ,[ETAFactory5] " & vbCrLf & _
                '                    "   FROM [dbo].[PO_Master_Export] a " & vbCrLf & _
                '                    "  INNER join PO_Detail_Export b on e.PONo = b.PONo AND e.AffiliateID = B.AffiliateID AND e.SupplierID = B.SupplierID  " & vbCrLf & _
                '                    " where 'A' = 'A' " & pWhere & ""

                Dim Cmd As New SqlCommand(ls_sql, cn)
                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                
                Return dt
            End Using
        Catch ex As Exception
            pErr = ex.Message
            Return Nothing
        End Try
    End Function


End Class
