Imports System.Data.SqlClient

Public Class clsPO
    Dim m_PONo As String
    Dim m_Quota As String
    Dim m_Supplier As String

    Public Property PONo As String
        Get
            Return m_PONo
        End Get
        Set(ByVal value As String)
            m_PONo = value
        End Set
    End Property

    Public Property Quota As String
        Get
            Return m_Quota
        End Get
        Set(ByVal value As String)
            m_Quota = value
        End Set
    End Property

    Public Property Supplier As String
        Get
            Return m_Supplier
        End Get
        Set(ByVal value As String)
            m_Supplier = value
        End Set
    End Property

    Public Function POKanban(ByVal pPONO As String, ByVal pAffiliateID As String, ByVal pSupplierID As String) As String
        Dim retVal As String = "NO"
        Dim clsGlobal As New clsGlobal

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim ls_SQL As String = "select top 1 KanbanCls from PO_Detail where PONo = '" & pPONO & "' and AffiliateID = '" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'"
            Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
            Dim sqlRdr As SqlDataReader = sqlCmd.ExecuteReader()

            If sqlRdr.Read() Then
                If sqlRdr("KanbanCls") = "1" Then
                    retVal = "YES"
                ElseIf sqlRdr("KanbanCls") = "0" Then
                    retVal = "NO"
                End If
            Else
                retVal = "NO"
            End If

            sqlConn.Close()
        End Using

        Return retVal
    End Function

    Public Function Check_CreateKanban(ByVal pPONO As String, ByVal pAffiliateID As String, ByVal pSupplierID As String) As Boolean
        Dim retVal As String = False
        Dim clsGlobal As New clsGlobal

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim ls_SQL As String = "select top 1 KanbanNo from Kanban_Detail where AffiliateID = '" & pAffiliateID & "' and PONo = '" & pPONO & "' and SupplierID = '" & pSupplierID & "'"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                retVal = True
            Else
                retVal = False
            End If

            sqlConn.Close()
        End Using

        Return retVal
    End Function

    Public Shared Function GetTable(ByVal pPONo As String, ByVal pAffiliateID As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        pErr = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " select a.Period, a.PONo, a.AffiliateID, case when b.KanbanCls = 0 then 'NO' else 'YES' end POKanban, a.ShipCls, " & vbCrLf & _
                          " b.PartNo, c.PartName, d.Description UOM, c.MOQ, c.Maker, c.Project, b.SupplierID, b.POQty, " & vbCrLf & _
                          " Forecast1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(MF.Period) = Year(DATEADD(MONTH,1,a.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,a.Period))),0), " & vbCrLf & _
                          " Forecast2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(MF.Period) = Year(DATEADD(MONTH,2,a.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,a.Period))),0), " & vbCrLf & _
                          " Forecast3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(MF.Period) = Year(DATEADD(MONTH,3,a.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,a.Period))),0), " & vbCrLf & _
                          " DeliveryD1, DeliveryD2, DeliveryD3, DeliveryD4, DeliveryD5, DeliveryD6, DeliveryD7, DeliveryD8, DeliveryD9, DeliveryD10, " & vbCrLf & _
                          " DeliveryD11, DeliveryD12, DeliveryD13, DeliveryD14, DeliveryD15, DeliveryD16, DeliveryD17, DeliveryD18, DeliveryD19, DeliveryD20, " & vbCrLf & _
                          " DeliveryD21, DeliveryD22, DeliveryD23, DeliveryD24, DeliveryD25, DeliveryD26, DeliveryD27, DeliveryD28, DeliveryD29, DeliveryD30, DeliveryD31 " & vbCrLf & _
                          " from PO_Master a " & vbCrLf & _
                          " inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                          " left join MS_Parts c on c.PartNo = b.PartNo "

                ls_sql = ls_sql + " left join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                                  " where a.PONo = '" & pPONo & "' and a.AffiliateID = '" & pAffiliateID & "'"

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
