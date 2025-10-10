Imports System.Data.SqlClient

Public Class clsMaster    
    Public Property SupplierID As String
    Public Property AffiliateID As String
    Public Property AffiliateCode As String
    Public Property PartNo As String
    Public Property ETDAffiliate As String
    Public Property ETDPASI As String

    Public Property ETAPASI As String
    Public Property ETDSupplier As String

    Public Property ETAAffiliate As String

    Public Property AffiliateName As String
    Public Property SupplierName As String
    Public Property Address As String
    Public Property City As String
    Public Property PostalCode As String
    Public Property Phone1 As String
    Public Property Phone2 As String
    Public Property Fax As String
    Public Property NPWP As String
    Public Property PODeliveryBy As String
    Public Property FolderOES As String

    Public Property SupplierCode As String
    Public Property SupplierType As String

    Public Property StartDate As String
    Public Property EndDate As String
    Public Property EffectiveDate As String
    Public Property CurrCls As String
    Public Property Price As Double
    Public Property PriceCategory As String

    Public Property Quota As Integer

    Public Property PartName As String
    Public Property CarMakerCode As String
    Public Property CarMakerName As String
    Public Property PartGroupName As String
    Public Property HSCode As String
    Public Property FGCls As String
    Public Property UnitCls As String
    Public Property KanbanCls As String
    Public Property Maker As String
    Public Property Project As String
    Public Property PackingCls As String
    Public Property MOQ As Integer
    Public Property QtyBox As Integer
    Public Property BoxPallet As Integer
    Public Property NetWeight As Double
    Public Property GrossWeight As Double
    Public Property ItmWidth As Double
    Public Property ItmLength As Double
    Public Property ItmHeight As Double

    Public Property MontlyCapacity As Double
    Public Property DailyCapacity As Double

    Public Property DeliveryLocationCode As String
    Public Property DeliveryLocationName As String
    Public Property DefaultCls As String

    Public Property KantorPabean As String
    Public Property IzinTPB As String
    Public Property BCPerson As String

    Public Property ConsigneeCode As String
    Public Property ConsigneeName As String
    Public Property ConsigneeAddress As String
    Public Property BuyerCode As String
    Public Property BuyerName As String
    Public Property BuyerAddress As String
    Public Property Location As String

    Public Property ETDVendor As String
    Public Property ETDPort As String
    Public Property ETAForwarder As String
    Public Property ETAPort As String
    Public Property ETAFactory As String
    Public Property CutOfDate As String
    Public Property Period As Date

    Public Property Week As String
    Public Property DestinationPort As String
    Public Property OverseasCls As String
    Public Property LabelCode As String


    Public Shared Function GetTableAffiliate(ByVal pAffiliateID As String, ByVal pAffiliateName As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pAffiliateID.Trim <> "" Then
            pWhere = pWhere + " and AffiliateID like '%" & pAffiliateID & "%'"
        End If

        If pAffiliateName.Trim <> "" Then
            pWhere = pWhere + " and AffiliateName like '%" & pAffiliateName & "%'"
        End If

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM([AffiliateID]) " & vbCrLf &
                          "       ,TRIM([AffiliateCode]) " & vbCrLf &
                          "       ,RTRIM([ConsigneeCode]) " & vbCrLf &
                          "       ,RTRIM([BuyerCode]) " & vbCrLf &
                          "       ,[AffiliateName] " & vbCrLf &
                          "       ,[Address] " & vbCrLf &
                          "       ,[ConsigneeName] " & vbCrLf &
                          "       ,[ConsigneeAddress] " & vbCrLf &
                          "       ,[BuyerName] " & vbCrLf &
                          "       ,[BuyerAddress] " & vbCrLf &
                          "       ,[DestinationPort] " & vbCrLf &
                          "       ,[City] " & vbCrLf &
                          "       ,[PostalCode] " & vbCrLf &
                          "       ,[Phone1] " & vbCrLf &
                          "       ,[Phone2] " & vbCrLf &
                          "       ,[Fax] " & vbCrLf &
                          "       ,[NPWP] " & vbCrLf &
                          "       ,[KantorPabean] " & vbCrLf &
                          "       ,[IzinTPB] " & vbCrLf &
                          "       ,[BCPerson] " & vbCrLf &
                          "       ,[PODeliveryBy] " & vbCrLf &
                          "       ,CASE WHEN OverseasCls = '1' THEN 'YES' ELSE 'NO' END OverseasCls" & vbCrLf &
                          "       ,[FolderOES] " & vbCrLf &
                          "   FROM [dbo].[MS_Affiliate] " & vbCrLf &
                          " where 'A' = 'A' " & pWhere & ""

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

    Public Shared Function GetTableDeliveryLocation(Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM([AffiliateID]) " & vbCrLf & _
                          "       ,[DeliveryLocationCode] " & vbCrLf & _
                          "       ,[DeliveryLocationName] " & vbCrLf & _
                          "       ,[Address] " & vbCrLf & _
                          "       ,[City] " & vbCrLf & _
                          "       ,[PostalCode] " & vbCrLf & _
                          "       ,[Phone1] " & vbCrLf & _
                          "       ,[Phone2] " & vbCrLf & _
                          "       ,[Fax] " & vbCrLf & _
                          "       ,[NPWP] " & vbCrLf & _
                          "       ,[PODeliveryBy] " & vbCrLf & _
                          "       ,[DefaultCls] " & vbCrLf & _
                          "   FROM [dbo].[MS_DeliveryPlace] " & vbCrLf & _
                          " where 'A' = 'A' " & pWhere & ""

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

    Public Shared Function GetTableSuppCapacity(ByVal pSupplier As String, ByVal pPartNo As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pSupplier.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and SupplierID = '" & pSupplier & "'"
        End If

        If pPartNo.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and PartNo = '" & pPartNo & "'"
        End If

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM([SupplierID]) " & vbCrLf & _
                          "       ,[PartNo] " & vbCrLf & _
                          "       ,[DailyDeliveryCapacity] " & vbCrLf & _
                          "       ,[MonthlyInjectionCapacity] " & vbCrLf & _
                          "   FROM [dbo].[MS_SupplierCapacity] " & vbCrLf & _
                          " where 'A' = 'A' " & pWhere & ""

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

    Public Shared Function GetTableSupplier(ByVal pAffiliateID As String, ByVal pAffiliateName As String, ByVal pSupplierType As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pAffiliateID.Trim <> "" Then
            pWhere = pWhere + " and SupplierID like '%" & pAffiliateID & "%'"
        End If

        If pAffiliateName.Trim <> "" Then
            pWhere = pWhere + " and SupplierName like '%" & pAffiliateName & "%'"
        End If

        If pSupplierType.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and SupplierType = '" & pSupplierType & "'"
        End If

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM([SupplierID]) " & vbCrLf & _
                          "       ,[SupplierName] " & vbCrLf & _
                          "       ,[SupplierType] " & vbCrLf & _
                          "       ,[SupplierCode] " & vbCrLf & _
                          "       ,[LabelCode] " & vbCrLf & _
                          "       ,[Address] " & vbCrLf & _
                          "       ,[City] " & vbCrLf & _
                          "       ,[PostalCode] " & vbCrLf & _
                          "       ,[Phone1] " & vbCrLf & _
                          "       ,[Phone2] " & vbCrLf & _
                          "       ,[Fax] "

                ls_sql = ls_sql + "       ,[NPWP] " & vbCrLf & _
                                  "   FROM [dbo].[MS_Supplier] " & vbCrLf & _
                                  " where 'A' = 'A' " & pWhere & ""

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

    Public Shared Function GetTablePart(ByVal pPartNo As String, ByVal pPartName As String, ByVal pMaker As String, ByVal pProject As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pPartNo.Trim <> "" Then
            pWhere = pWhere + " and PartNo like '%" & pPartNo & "%'"
        End If

        If pPartName.Trim <> "" Then
            pWhere = pWhere + " and PartName like '%" & pPartName & "%'"
        End If

        If pMaker.Trim <> "" Then
            pWhere = pWhere + " and Maker = '" & pMaker & "'"
        End If

        If pProject.Trim <> "" Then
            pWhere = pWhere + " and Project = '" & pProject & "'"
        End If

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM(a.[PartNo])PartNo " & vbCrLf & _
                          "       ,a.[PartName] " & vbCrLf & _
                          "       ,RTRIM(a.[PartCarMaker]) PartCarMaker " & vbCrLf & _
                          "       ,RTRIM(a.[PartCarName]) PartCarName " & vbCrLf & _
                          "       ,RTRIM(a.[PartGroupName]) PartGroupName " & vbCrLf & _
                          "       ,RTRIM(a.[HSCode]) HSCode " & vbCrLf & _
                          "       ,RTRIM(b.[Description])UnitCls" & vbCrLf & _
                          "       ,CASE WHEN a.[KanbanCls] = '1' then 'YES' else 'NO' END KanbanCls" & vbCrLf & _
                          "       ,a.[Maker] " & vbCrLf & _
                          "       ,a.[Project] " & vbCrLf & _
                          "   FROM [MS_Parts] a LEFT JOIN MS_UnitCls b ON a.UnitCls = b.UnitCls " & vbCrLf & _
                          " where 'A' = 'A' " & pWhere & ""

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

    Public Shared Function GetTablePartMapping(ByVal pPartNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pPartNo.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and PartNo like '%" & pPartNo & "%'"
        End If

        If pAffiliate.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and AffiliateID = '" & pAffiliate & "'"
        End If

        If pSupplier.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and SupplierID = '" & pSupplier & "'"
        End If

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM([PartNo])PartNo " & vbCrLf & _
                          "       ,RTRIM([AffiliateID])AffiliateID " & vbCrLf & _
                          "       ,RTRIM([SupplierID])SupplierID " & vbCrLf & _
                          "       ,[Quota] " & vbCrLf & _
                          "       ,RTRIM([LocationID]) LocationID " & vbCrLf & _
                          "       ,RTRIM([PackingCls]) PackingCls " & vbCrLf & _
                          "       ,[MOQ] " & vbCrLf & _
                          "       ,[QtyBox] " & vbCrLf & _
                          "       ,[BoxPallet] " & vbCrLf & _
                          "       ,[NetWeight] " & vbCrLf & _
                          "       ,[GrossWeight] " & vbCrLf & _
                          "       ,[Length] " & vbCrLf & _
                          "       ,[Width] " & vbCrLf & _
                          "       ,[Height] " & vbCrLf & _
                          "   FROM [dbo].[MS_PartMapping] " & vbCrLf & _
                          " where 'A' = 'A' " & pWhere & ""

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

    Public Shared Function GetTablePriceSupplier(ByVal pPartNo As String, ByVal pAffiliate As String, ByVal pAdditional As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pPartNo.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and PartNo = '" & pPartNo & "'"
        End If

        If pAffiliate.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and AffiliateID = '" & pAffiliate & "'"
        End If
        pWhere = pWhere + pAdditional

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = "SELECT RTRIM(a.AffiliateID) SupplierID, RTRIM(PartNo) PartNo, " & vbCrLf & _
                            "StartDate, " & vbCrLf & _
                            "EndDate, " & vbCrLf & _
                            "EffectiveDate, " & vbCrLf & _
                            "RTRIM(ISNULL(c.Description, ''))CurrCls, " & vbCrLf & _
                            "Price, " & vbCrLf & _
                            "RTRIM(ISNULL(d.Description, ''))PackingType, " & vbCrLf & _
                            "RTRIM(ISNULL(e.Description, ''))PriceCategory, " & vbCrLf & _
                            "RTRIM(DeliveryLocationID)DeliveryLocationID " & vbCrLf & _
                            "FROM dbo.MS_Price a " & vbCrLf & _
                            "INNER JOIN MS_Supplier b on a.AffiliateID = b.SupplierID " & vbCrLf & _
                            "LEFT JOIN MS_CurrCls c on a.CurrCls = c.CurrCls " & vbCrLf & _
                            "LEFT JOIN MS_PackingCls d on a.PackingCls = d.PackingCls " & vbCrLf & _
                            "LEFT JOIN MS_PriceCls e on a.PriceCls = e.PriceCls " & vbCrLf & _
                            "WHERE 'A' = 'A' " & pWhere & ""

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

    Public Shared Function GetTablePricePASI(ByVal pPartNo As String, ByVal pAffiliate As String, ByVal pAdditional As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pPartNo.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and PartNo = '" & pPartNo & "'"
        End If

        If pAffiliate.Trim <> clsGlobal.gs_All Then
            pWhere = pWhere + " and a.AffiliateID = '" & pAffiliate & "'"
        End If
        pWhere = pWhere + pAdditional

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM(a.[AffiliateID]) " & vbCrLf & _
                          "       ,RTRIM([PartNo]) " & vbCrLf & _
                          "       ,RTRIM(c.Description) " & vbCrLf & _
                          "       ,RTRIM(d.Description) " & vbCrLf & _
                          "       ,CASE WHEN [CurrCls] = '01' then 'JPY' " & vbCrLf & _
                          "             WHEN [CurrCls] = '02' then 'USD' " & vbCrLf & _
                          "             WHEN [CurrCls] = '03' then 'IDR' " & vbCrLf & _
                          "             WHEN [CurrCls] = '04' then 'SGD' " & vbCrLf & _
                          "             WHEN [CurrCls] = '05' then 'EUR' END CurrCls" & vbCrLf & _
                          "       ,[Price] " & vbCrLf & _
                          "       ,[StartDate] " & vbCrLf & _
                          "       ,[EndDate] " & vbCrLf & _
                          "       ,[EffectiveDate] " & vbCrLf & _
                          "   FROM [dbo].[MS_Price] a" & vbCrLf & _
                          "   INNER JOIN MS_Affiliate b on a.AffiliateID = b.AffiliateID" & vbCrLf & _
                          "   LEFT JOIN MS_PackingCls c on a.PackingCls = c.PackingCls" & vbCrLf & _
                          "   LEFT JOIN MS_PriceCls d on a.PriceCls = d.PriceCls" & vbCrLf & _
                          "   where 'A' = 'A' " & pWhere & ""

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

    Public Shared Function GetTableETDPASI(ByVal pPeriod As Date, ByVal pAffiliateID As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM(AffiliateID), ETAAffiliate, ETDPASI " & vbCrLf & _
                                  "   FROM [dbo].[MS_ETD_PASI] " & vbCrLf & _
                                  " where YEAR(ETAAffiliate) = " & Year(pPeriod) & " and MONTH(ETAAffiliate) = " & Month(pPeriod) & " and AffiliateID = '" & pAffiliateID & "'"

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

    Public Shared Function GetTableETDViaPASI(ByVal pPeriod As Date, ByVal pAffiliateID As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM(SupplierID), ETAPASI, ETDSupplier " & vbCrLf & _
                                  "   FROM [dbo].[MS_ETD_Supplier_PASI] " & vbCrLf & _
                                  " where YEAR(ETAPASI) = " & Year(pPeriod) & " and MONTH(ETAPASI) = " & Month(pPeriod) & " and SupplierID = '" & pAffiliateID & "'"

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

    Public Shared Function GetTableETDDirect(ByVal pPeriod As Date, ByVal pAffiliateID As String, ByVal pSupplierID As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT RTRIM(SupplierID), RTRIM(AffiliateID), ETAAffiliate, ETDSupplier " & vbCrLf & _
                                  "   FROM [dbo].[MS_ETD_Supplier_Direct] " & vbCrLf & _
                                  " where YEAR(ETAAffiliate) = " & Year(pPeriod) & " and MONTH(ETAAffiliate) = " & Month(pPeriod) & " and AffiliateID = '" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'"

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

    Public Shared Function GetTableSummaryForecastPO(ByVal pPeriod As Date, ByVal pPartNo As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pPartNo <> clsGlobal.gs_All Then
            pWhere = pWhere & " and b.PartNo = '" & pPartNo & "'"
        End If


        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " declare @period date  " & vbCrLf & _
                  " set @period = '" & Format(pPeriod, "yyyy-MM") & "-01" & "'  " & vbCrLf & _
                  "  " & vbCrLf & _
                  "  select " & vbCrLf & _
                  " 	a.NoUrut, " & vbCrLf & _
                  " 	a.PartNo, " & vbCrLf & _
                  " 	a.PartName, " & vbCrLf & _
                  " 	--a.thn, " & vbCrLf & _
                  " 	--a.bln, " & vbCrLf & _
                  " 	BulanDesc, " & vbCrLf & _
                  " 	--a.DescUrut, " & vbCrLf

                ls_sql = ls_sql + " 	a.DescName, " & vbCrLf & _
                                  " 	max(isnull(b.qty1,0)) qty1, " & vbCrLf & _
                                  " 	max(isnull(b.qty2,0)) qty2, " & vbCrLf & _
                                  " 	max(isnull(b.qty3,0)) qty3, " & vbCrLf & _
                                  " 	max(isnull(b.qty4,0)) qty4, " & vbCrLf & _
                                  " 	max(isnull(b.qty5,0)) qty5 " & vbCrLf & _
                                  "  from  " & vbCrLf & _
                                  "  (  " & vbCrLf & _
                                  "  	select a.*,b.*, c.* " & vbCrLf & _
                                  "  	from  " & vbCrLf & _
                                  "  	(  "

                ls_sql = ls_sql + "  		select row_number() over (order by PartNo asc) as NoUrut, * from  " & vbCrLf & _
                                  "  		(  " & vbCrLf & _
                                  "  			select distinct a.partno, b.PartName from PO_Detail a  " & vbCrLf & _
                                  "  			inner join PO_Master c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID  " & vbCrLf & _
                                  "  			left join MS_Parts b on a.PartNo = b.PartNo  " & vbCrLf & _
                                  "  			where FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) --and a.AffiliateID = 'SUAI'  " & vbCrLf & _
                                  " 			--and a.partno='7184-8544' " & vbCrLf & _
                                  "  		) xuz  " & vbCrLf & _
                                  "  	) a  " & vbCrLf & _
                                  "  	cross join  " & vbCrLf & _
                                  "  	(		  "

                ls_sql = ls_sql + "  			select tahun thn,bulan bln,Tgl,BulanDesc   " & vbCrLf & _
                                  "  			from ms_period   " & vbCrLf & _
                                  "  			where tgl between @period and dateadd(month,2,@period)		  " & vbCrLf & _
                                  "  	) b " & vbCrLf & _
                                  " 	cross join " & vbCrLf & _
                                  " 	( " & vbCrLf & _
                                  " 		select '1'DescUrut, 'Total Forecast' DescName " & vbCrLf & _
                                  " 		union all " & vbCrLf & _
                                  " 		select '2'DescUrut, 'Total PO' DescName " & vbCrLf & _
                                  " 		union all " & vbCrLf & _
                                  " 		select '3'DescUrut, 'Total Delivery' DescName "

                ls_sql = ls_sql + " 		union all " & vbCrLf & _
                                  " 		select '4'DescUrut, 'Balance PO' DescName " & vbCrLf & _
                                  " 		union all " & vbCrLf & _
                                  " 		select '5'DescUrut, 'Diff' DescName " & vbCrLf & _
                                  " 	)c " & vbCrLf & _
                                  "  ) a  " & vbCrLf & _
                                  "  left join  " & vbCrLf & _
                                  "  ( " & vbCrLf & _
                                  "  	select '2' SeqNo, a.partno,a.thn,a.bln,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 1 then POQty end),0) qty1,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 2 then POQty end),0) qty2,  "

                ls_sql = ls_sql + "  		isnull(max(case when c.kd = 3 then POQty end),0) qty3,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 4 then POQty end),0) qty4,   " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 5 then POQty end),0) qty5 " & vbCrLf & _
                                  "  	from  " & vbCrLf & _
                                  "  	(  " & vbCrLf & _
                                  "  		select b.PartNo,year(Period) thn,month(period) bln,sum(isnull(c.POQty,b.POQty)) poqty from PO_Master a   " & vbCrLf & _
                                  "   		inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID  " & vbCrLf & _
                                  "  		left join  " & vbCrLf & _
                                  "  		(  " & vbCrLf & _
                                  "  			select a.* from PORev_Detail a  " & vbCrLf & _
                                  "  			inner join PORev_Master b on a.PONo = b.PONo and a.PORevNo = b.PORevNo and a.SeqNo = b.SeqNo  "

                ls_sql = ls_sql + "  			inner join (select MAX(SeqNo) SeqNo, PONo from PORev_Detail po group by PONo) c on a.PONo = c.PONo and a.SeqNo = c.SeqNo  " & vbCrLf & _
                                  "  			where b.FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) --and a.AffiliateID = 'SUAI' " & vbCrLf & _
                                  "  		)c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID  " & vbCrLf & _
                                  "  		left join MS_Parts d on b.PartNo = d.PartNo  " & vbCrLf & _
                                  "  		where a.FinalApproveDate is not null and Period between @period and dateadd(month,4,@period) --and a.AffiliateID = 'SUAI'  " & vbCrLf & _
                                  "   		group by b.partno,year(period),month(period)   " & vbCrLf & _
                                  "  	) a  " & vbCrLf & _
                                  "  	inner join  " & vbCrLf & _
                                  "  	(  " & vbCrLf & _
                                  "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                                  "  		where tgl between @period and dateadd(month,4,@period)		  "

                ls_sql = ls_sql + "  	) c  " & vbCrLf & _
                                  "  		on c.tahun = a.thn and c.bulan = a.bln  " & vbCrLf & _
                                  "  	group by a.partno,a.thn,a.bln  " & vbCrLf & _
                                  "   " & vbCrLf & _
                                  " 	union  all " & vbCrLf & _
                                  "  " & vbCrLf & _
                                  " 	select '1' SeqNo, a.partno,a.thn,a.bln,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 1 then Qty end),0) qty1,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 2 then Qty end),0) qty2,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 3 then Qty end),0) qty3,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 4 then Qty end),0) qty4,   "

                ls_sql = ls_sql + "  		isnull(max(case when c.kd = 5 then Qty end),0) qty5 " & vbCrLf & _
                                  " 	from " & vbCrLf & _
                                  " 	( " & vbCrLf & _
                                  " 		select partno,year(period) thn,month(period) bln,sum(Qty) qty from ms_forecast " & vbCrLf & _
                                  " 		where period between @period and dateadd(month,4,@period) " & vbCrLf & _
                                  " 		group by PartNo, Period " & vbCrLf & _
                                  " 	) a " & vbCrLf & _
                                  "  	inner join  " & vbCrLf & _
                                  "  	(  " & vbCrLf & _
                                  "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                                  "  		where tgl between @period and dateadd(month,4,@period)		  "

                ls_sql = ls_sql + "  	) c  " & vbCrLf & _
                                  "  		on c.tahun = a.thn and c.bulan = a.bln  " & vbCrLf & _
                                  " 	group by a.partno,a.thn,a.bln  " & vbCrLf & _
                                  "  " & vbCrLf & _
                                  " 	union all " & vbCrLf & _
                                  "  " & vbCrLf & _
                                  " 	select '3' SeqNo, a.partno,a.thn,a.bln,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 1 then Qty end),0) qty1,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 2 then Qty end),0) qty2,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 3 then Qty end),0) qty3,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 4 then Qty end),0) qty4,   "

                ls_sql = ls_sql + "  		isnull(max(case when c.kd = 5 then Qty end),0) qty5 " & vbCrLf & _
                                  " 	from " & vbCrLf & _
                                  " 	( " & vbCrLf & _
                                  " 		select partno,left(kanbanno,4) thn,substring(kanbanno,5,2) bln,sum(doqty) qty from dosupplier_detail " & vbCrLf & _
                                  " 		where cast(left(kanbanno,8) as date) between @period and dateadd(month,4,@period) " & vbCrLf & _
                                  " 		group by partno,left(kanbanno,4),substring(kanbanno,5,2)		 " & vbCrLf & _
                                  " 	) a " & vbCrLf & _
                                  "  	inner join  " & vbCrLf & _
                                  "  	(  " & vbCrLf & _
                                  "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                                  "  		where tgl between @period and dateadd(month,4,@period)		  "

                ls_sql = ls_sql + "  	) c  " & vbCrLf & _
                                  "  		on c.tahun = a.thn and c.bulan = a.bln  " & vbCrLf & _
                                  " 	group by a.partno,a.thn,a.bln  " & vbCrLf & _
                                  "  " & vbCrLf & _
                                  " 	union all " & vbCrLf & _
                                  "  " & vbCrLf & _
                                  " 	select '4' SeqNo, a.partno,a.thn,a.bln,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 1 then poqty end),0) - isnull(max(case when c.kd = 1 then Qty end),0) qty1,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 2 then poqty end),0) - isnull(max(case when c.kd = 2 then Qty end),0) qty2,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 3 then poqty end),0) - isnull(max(case when c.kd = 3 then Qty end),0) qty3,  " & vbCrLf & _
                                  "  		isnull(max(case when c.kd = 4 then poqty end),0) - isnull(max(case when c.kd = 4 then Qty end),0) qty4,   "

                ls_sql = ls_sql + "  		isnull(max(case when c.kd = 5 then poqty end),0) - isnull(max(case when c.kd = 5 then Qty end),0) qty5 " & vbCrLf & _
                                  " 	from " & vbCrLf & _
                                  " 	( " & vbCrLf & _
                                  " 		select partno,left(kanbanno,4) thn,substring(kanbanno,5,2) bln,sum(doqty) qty from dosupplier_detail " & vbCrLf & _
                                  " 		where cast(left(kanbanno,8) as date) between @period and dateadd(month,4,@period) " & vbCrLf & _
                                  " 		group by partno,left(kanbanno,4),substring(kanbanno,5,2)		 " & vbCrLf & _
                                  " 	) a " & vbCrLf & _
                                  " 	left join " & vbCrLf & _
                                  " 	( " & vbCrLf & _
                                  " 		select b.PartNo,year(Period) thn,month(period) bln,sum(isnull(c.POQty,b.POQty)) poqty from PO_Master a   " & vbCrLf & _
                                  "   		inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID  "

                ls_sql = ls_sql + "  		left join  " & vbCrLf & _
                                  "  		(  " & vbCrLf & _
                                  "  			select a.* from PORev_Detail a  " & vbCrLf & _
                                  "  			inner join PORev_Master b on a.PONo = b.PONo and a.PORevNo = b.PORevNo and a.SeqNo = b.SeqNo  " & vbCrLf & _
                                  "  			inner join (select MAX(SeqNo) SeqNo, PONo from PORev_Detail po group by PONo) c on a.PONo = c.PONo and a.SeqNo = c.SeqNo  " & vbCrLf & _
                                  "  			where b.FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) --and a.AffiliateID = 'SUAI' " & vbCrLf & _
                                  "  		)c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID  " & vbCrLf & _
                                  "  		left join MS_Parts d on b.PartNo = d.PartNo  " & vbCrLf & _
                                  "  		where a.FinalApproveDate is not null and Period between @period and dateadd(month,4,@period) --and a.AffiliateID = 'SUAI'  " & vbCrLf & _
                                  "   		group by b.partno,year(period),month(period)   " & vbCrLf & _
                                  " 	)b on a.bln = b.bln and a.PartNo = b.PartNo and a.thn = b.thn "

                ls_sql = ls_sql + "  	inner join  " & vbCrLf & _
                                  "  	(  " & vbCrLf & _
                                  "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                                  "  		where tgl between @period and dateadd(month,4,@period)		  " & vbCrLf & _
                                  "  	) c  " & vbCrLf & _
                                  "  		on c.tahun = a.thn and c.bulan = a.bln and b.bln = c.Bulan and b.thn= c.Tahun " & vbCrLf & _
                                  " 	group by a.partno,a.thn,a.bln  " & vbCrLf & _
                                  "  " & vbCrLf & _
                                  " 	union all " & vbCrLf & _
                                  "  " & vbCrLf & _
                                  " 	select '5' SeqNo, a.partno,a.thn,a.bln,  "

                ls_sql = ls_sql + "  		case when isnull(max(case when c.kd = 1 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 1 then poqty end),0) - isnull(max(case when c.kd = 1 then Qty end),0)) / isnull(max(case when c.kd = 1 then poqty end),0)) * 100) end  qty1, " & vbCrLf & _
                                  "  		case when isnull(max(case when c.kd = 2 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 2 then poqty end),0) - isnull(max(case when c.kd = 2 then Qty end),0)) / isnull(max(case when c.kd = 2 then poqty end),0)) * 100) end  qty2, " & vbCrLf & _
                                  "  		case when isnull(max(case when c.kd = 3 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 3 then poqty end),0) - isnull(max(case when c.kd = 3 then Qty end),0)) / isnull(max(case when c.kd = 3 then poqty end),0)) * 100) end  qty3, " & vbCrLf & _
                                  "  		case when isnull(max(case when c.kd = 4 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 4 then poqty end),0) - isnull(max(case when c.kd = 4 then Qty end),0)) / isnull(max(case when c.kd = 4 then poqty end),0)) * 100) end  qty4, " & vbCrLf & _
                                  "  		case when isnull(max(case when c.kd = 5 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 5 then poqty end),0) - isnull(max(case when c.kd = 5 then Qty end),0)) / isnull(max(case when c.kd = 5 then poqty end),0)) * 100) end  qty5 " & vbCrLf & _
                                  " 	from " & vbCrLf & _
                                  " 	( " & vbCrLf & _
                                  " 		select partno,year(period) thn,month(period) bln,sum(Qty) qty from ms_forecast " & vbCrLf & _
                                  " 		where period between @period and dateadd(month,4,@period) " & vbCrLf & _
                                  " 		group by PartNo, Period " & vbCrLf & _
                                  " 	) a "

                ls_sql = ls_sql + " 	left join " & vbCrLf & _
                                  " 	( " & vbCrLf & _
                                  " 		select b.PartNo,year(Period) thn,month(period) bln,sum(isnull(c.POQty,b.POQty)) poqty from PO_Master a   " & vbCrLf & _
                                  "   		inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID  " & vbCrLf & _
                                  "  		left join  " & vbCrLf & _
                                  "  		(  " & vbCrLf & _
                                  "  			select a.* from PORev_Detail a  " & vbCrLf & _
                                  "  			inner join PORev_Master b on a.PONo = b.PONo and a.PORevNo = b.PORevNo and a.SeqNo = b.SeqNo  " & vbCrLf & _
                                  "  			inner join (select MAX(SeqNo) SeqNo, PONo from PORev_Detail po group by PONo) c on a.PONo = c.PONo and a.SeqNo = c.SeqNo  " & vbCrLf & _
                                  "  			where b.FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) --and a.AffiliateID = 'SUAI' " & vbCrLf & _
                                  "  		)c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID  "

                ls_sql = ls_sql + "  		left join MS_Parts d on b.PartNo = d.PartNo  " & vbCrLf & _
                                  "  		where a.FinalApproveDate is not null and Period between @period and dateadd(month,4,@period) --and a.AffiliateID = 'SUAI'  " & vbCrLf & _
                                  "   		group by b.partno,year(period),month(period)   " & vbCrLf & _
                                  " 	)b on a.bln = b.bln and a.PartNo = b.PartNo and a.thn = b.thn " & vbCrLf & _
                                  "  	inner join  " & vbCrLf & _
                                  "  	(  " & vbCrLf & _
                                  "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                                  "  		where tgl between @period and dateadd(month,4,@period)		  " & vbCrLf & _
                                  "  	) c  " & vbCrLf & _
                                  "  		on c.tahun = a.thn and c.bulan = a.bln and b.bln = c.Bulan and b.thn= c.Tahun " & vbCrLf & _
                                  " 	group by a.partno,a.thn,a.bln  "

                ls_sql = ls_sql + " ) b   " & vbCrLf & _
                                  "  	on b.SeqNo = a.DescUrut and b.partno = a.partno and b.thn = a.thn  " & vbCrLf & _
                                  " 		and (b.seqno in (2,3,4,5) and b.bln = a.bln or b.seqno = 1 and b.bln >= a.bln)  " & vbCrLf & _
                                  " 	group by a.NoUrut, " & vbCrLf & _
                                  " 	a.PartNo, " & vbCrLf & _
                                  " 	a.PartName, " & vbCrLf & _
                                  " 	a.thn, " & vbCrLf & _
                                  " 	a.bln, " & vbCrLf & _
                                  " 	BulanDesc, " & vbCrLf & _
                                  " 	a.DescUrut, " & vbCrLf & _
                                  " 	a.DescName " & vbCrLf & _
                                  "  order by a.PartNo,	a.thn, a.bln, a.DescUrut " 

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

    Public Shared Function GetTableSummaryForecast(ByVal pPeriod As Date, ByVal pAffiliateID As String, ByVal pSupplierID As String, ByVal pPartNo As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        If pPartNo <> clsGlobal.gs_All Then
            pWhere = pWhere & " and b.PartNo = '" & pPartNo & "'"
        End If

        If pAffiliateID <> clsGlobal.gs_All Then
            pWhere = pWhere & " and b.AffiliateID = '" & pAffiliateID & "'"
        End If

        If pSupplierID <> clsGlobal.gs_All Then
            pWhere = pWhere & " and b.SupplierID = '" & pSupplierID & "'"
        End If

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT  " & vbCrLf & _
                                  "     row_number() over (order by b.PartNo) as NoUrut, " & vbCrLf & _
                                  " 	RTRIM(b.PartNo)PartNo,  " & vbCrLf & _
                                  " 	RTRIM(b.AffiliateID)AffiliateID, " & vbCrLf & _
                                  " 	RTRIM(b.SupplierID)SupplierID, " & vbCrLf & _
                                  " 	MOQ, " & vbCrLf & _
                                  " 	RTRIM(c.Project)Project, " & vbCrLf & _
                                  " 	RTRIM(b.PONo)PONo, " & vbCrLf & _
                                  " 	ISNULL(b.POQty,0) Bln1, " & vbCrLf & _
                                  " 	ISNULL(b.ForecastN1,0) Bln2, " & vbCrLf & _
                                  " 	ISNULL(b.ForecastN2,0) Bln3, " & vbCrLf & _
                                  " 	ISNULL(b.ForecastN3,0) Bln4 " & vbCrLf

                ls_sql = ls_sql + " FROM PO_Master a with (nolock)" & vbCrLf & _
                                  " INNER JOIN PO_Detail b with (nolock) on a.PONO = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                                  " LEFT JOIN MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                                  " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = b.PartNo AND MPM.AffiliateID = b.AffiliateID AND MPM.SupplierID = b.SupplierID  " & vbCrLf & _
                                  " WHERE YEAR(Period) = " & Year(pPeriod) & " and MONTH(Period) = " & Month(pPeriod) & " and FinalApproveDate IS NOT NULL " & pWhere & "" & vbCrLf & _
                                  " ORDER BY b.PartNo "

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

    Public Shared Function GetTableETDExport(ByVal pPeriod As Date, ByVal pAffiliateID As String, ByVal pSupplierID As String, Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " SELECT Period, AffiliateID = RTRIM(AffiliateID), SupplierID = RTRIM(SupplierID), Week, ETDVendor, ETAForwarder, ETDPort, ETAPort, ETAFactory, CutOfDate " & vbCrLf & _
                                  "   FROM [dbo].[MS_ETD_Export] " & vbCrLf & _
                                  " where YEAR(Period) = " & Year(pPeriod) & " and MONTH(Period) = " & Month(pPeriod) & " and AffiliateID = '" & pAffiliateID & "' "

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
