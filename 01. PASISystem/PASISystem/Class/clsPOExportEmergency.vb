Imports System.Data.SqlClient

Public Class clsPOExportEmergency
    Public Property SupplierID As String
    Public Property AffiliateID As String
    Public Property PartNo As String
    Public Property Period As Date
    Public Property POMonthly As String
    Public Property DeliveryLocation As String

    Public Property Order1 As String

    Public Property Vendor1 As Date


    Public Property ETDPort1 As Date


    Public Property ETAPort1 As Date

    Public Property ETAFactory1 As Date

    Public Shared Function GetTableEmergency(ByVal pAffiliateID As String, ByVal pAffiliateName As String, ByVal pPeriod As Date, ByVal pPOMonthly As String, ByVal pDeliveryLocation As String, ByVal pDeliveryLocationName As String, ByVal pOrder1 As String, ByVal pVendor1 As Date, ByVal pETDPort1 As Date, ByVal pETAPort1 As Date, ByVal pETAFactory1 As Date, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pAffiliateID.Trim <> "" Then
            pWhere = pWhere + " and AffiliateID like '%" & pAffiliateID & "%'"
        End If

        If pAffiliateName.Trim <> "" Then
            pWhere = pWhere + " and AffiliateName like '%" & pAffiliateName & "%'"
        End If

        'If pCommercial.Trim <> "" Then
        '    pWhere = pWhere + " and CommercialCls like '%" & pCommercial & "%'"
        'End If

        If pPeriod <> "" Then
            pWhere = pWhere + " and Period like '%" & pPeriod & "%'"
        End If

        If pPOMonthly <> "" Then
            pWhere = pWhere + " and EmergencyCls like '%" & pPOMonthly & "%'"
        End If

        If pDeliveryLocation <> "" Then
            pWhere = pWhere + " and ForwarderID like '%" & pDeliveryLocation & "%'"
        End If

        If pOrder1 <> "" Then
            pWhere = pWhere + " and OrderNo1 like '%" & pOrder1 & "%'"
        End If

        If pVendor1 <> "" Then
            pWhere = pWhere + " and ETDVendor1 like '%" & pVendor1 & "%'"
        End If

        If pETDPort1 <> "" Then
            pWhere = pWhere + " and ETDPort1 like '%" & pETDPort1 & "%'"
        End If

        If pETAPort1 <> "" Then
            pWhere = pWhere + " and ETAPort1 like '%" & pETAPort1 & "%'"
        End If

        If pETAFactory1 <> "" Then
            pWhere = pWhere + " and ETAFactory1 like '%" & pETAFactory1 & "%'"
        End If

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""

                ls_sql = " select " & vbCrLf & _
                  "    '0' AllowAccess, " & vbCrLf & _
                  " 	row_number() over (order by a.AffiliateID, a.SupplierID ) NoUrut,  " & vbCrLf & _
                  "  	RTRIM(a.PartNo)PartNo,  " & vbCrLf & _
                  "  	RTRIM(b.PartName)PartName,  " & vbCrLf & _
                  "  	RTRIM(b.UnitCls)UnitCls,  " & vbCrLf & _
                  "  	RTRIM(b.MOQ)MOQ,  " & vbCrLf & _
                  "  	b.QtyBox,  " & vbCrLf & _
                  "  	'0' Week1,   " & vbCrLf & _
                  "  	'0' Week2,   " & vbCrLf & _
                  "  	'0' Week3,   "

                ls_sql = ls_sql + "  	'0' Week4,   " & vbCrLf & _
                                  "  	'0' Week5,   " & vbCrLf & _
                                  "  	 TotalPOQty = (Week1+Week2+Week3+Week4+Week5),   " & vbCrLf & _
                                  "  	e.PONo,   " & vbCrLf & _
                                  "  	e.ShipCls,   " & vbCrLf & _
                                  "  	e.CommercialCls,   " & vbCrLf & _
                                  "  	e.ForwarderID,   " & vbCrLf & _
                                  "  	e.Period,   " & vbCrLf & _
                                  " d.OrderNo1, " & vbCrLf & _
                " d.ETDVendor1, " & vbCrLf & _
                " d.ETDPort1, " & vbCrLf & _
                " d.ETAPort1, " & vbCrLf & _
                " d.ETAFactory1, " & vbCrLf & _
                                  "  	RTRIM(a.AffiliateID)AffiliateID,  " & vbCrLf & _
                                  "  	RTRIM(a.SupplierID)SupplierID  " & vbCrLf & _
                                  "  from MS_PartMapping a  " & vbCrLf & _
                                  "  INNER join MS_Parts b on a.PartNo = b.PartNo  " & vbCrLf & _
                                  "  LEFT join MS_UnitCls c on c.UnitCls = b.UnitCls  " & vbCrLf & _
                                  "  left join PO_Detail_Export d on a.PartNo = d.PartNo AND A.AffiliateID = D.AffiliateID AND A.SupplierID = D.SupplierID  " & vbCrLf & _
                                  "  Left join PO_Master_Export e on d.PONo = e.PONo AND d.AffiliateID = e.AffiliateID AND d.SupplierID = e.SupplierID  " & vbCrLf & _
                                  "  where a.AffiliateID = a.AffiliateID  AND NOT EXISTS " & vbCrLf & _
                                  "  ( " & vbCrLf & _
                                  " SELECT * FROM  PO_Detail_Export X WHERE X.PONo = e.PONo and 'A' = 'A' " & pWhere & " " & vbCrLf & _
                                  "  ) "
                ' and a.SupplierID = '" & cboSupplier.Text.Trim & "'
                ls_sql = ls_sql + "  union all " & vbCrLf & _
                                  "  select '1' AllowAccess, " & vbCrLf & _
                                  " 	row_number() over (order by e.AffiliateID, e.SupplierID ) NoUrut,  " & vbCrLf & _
                                  "  	RTRIM(B.PartNo)PartNo,  " & vbCrLf & _
                                  "  	RTRIM(C.PartName)PartName,  " & vbCrLf & _
                                  "  	RTRIM(d.Description)Description,  " & vbCrLf & _
                                  "  	RTRIM(C.MOQ)MOQ,  " & vbCrLf & _
                                  "  	C.QtyBox,  " & vbCrLf & _
                                  "  	B.Week1,   " & vbCrLf & _
                                  "  	B.Week2,   " & vbCrLf & _
                                  "  	B.Week3,   "

                ls_sql = ls_sql + "  	B.Week4,   " & vbCrLf & _
                                  "  	B.Week5,   " & vbCrLf & _
                                  "  	TotalPOQty = (Week1+Week2+Week3+Week4+Week5),   " & vbCrLf & _
                                  "  	e.PONo,   " & vbCrLf & _
                                  "  	e.ShipCls,   " & vbCrLf & _
                                  "  	e.CommercialCls,   " & vbCrLf & _
                                  "  	e.ForwarderID,   " & vbCrLf & _
                                  "  	e.Period,   " & vbCrLf & _
                                  " d.OrderNo1, " & vbCrLf & _
                " d.ETDVendor1, " & vbCrLf & _
                " d.ETDPort1, " & vbCrLf & _
                " d.ETAPort1, " & vbCrLf & _
                " d.ETAFactory1, " & vbCrLf & _
                                  "  	RTRIM(e.AffiliateID)AffiliateID,  " & vbCrLf & _
                                  "  	RTRIM(e.SupplierID)SupplierID  " & vbCrLf & _
                                  " from PO_Master_Export e  " & vbCrLf & _
                                  "  INNER join PO_Detail_Export b on e.PONo = b.PONo AND e.AffiliateID = B.AffiliateID AND e.SupplierID = B.SupplierID  " & vbCrLf & _
                                  "  LEFT join MS_Parts c on c.PartNo = B.PartNo  " & vbCrLf & _
                                  "  LEFT join MS_UnitCls d on d.UnitCls = c.UnitCls  " & vbCrLf & _
                                  "  where 'A' = 'A' " & pWhere & " " & vbCrLf & _
                                  "  "

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
