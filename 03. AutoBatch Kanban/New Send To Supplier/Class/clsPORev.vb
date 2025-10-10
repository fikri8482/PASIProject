Public Class clsPORev
    Shared Sub up_SendPORevDomestic(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

    End Sub

    '    Private Function bindDataHeaderPORev(ByVal pDate As Date, ByVal pAffCode As String, ByVal pPORevNo As String, ByVal pPONo As String, ByVal pSupplierID As String) As DataSet
    '        Dim ls_SQL As String = ""
    '        MdlConn.ReadConnection()
    '        ls_SQL = "   SELECT DISTINCT POM.Period PODate,PORM.Period PORevDate,PORD.AffiliateID,AffiliateName,PORM.PORevNo,PORM.PONo  " & vbCrLf & _
    '                  "   ,CASE WHEN CommercialCls = '0' THEN 'NO' ELSE 'YES' END CommercialCls  " & vbCrLf & _
    '                  "   ,PORD.SupplierID,SupplierName,ShipCls   " & vbCrLf & _
    '                  "   ,PODeliveryBy   " & vbCrLf & _
    '                  "   ,MP.KanbanCls   " & vbCrLf & _
    '                  "   ,CONVERT(DATETIME,PORM.EntryDate,120)EntryDate,ISNULL(PORM.EntryUser,'')EntryUser --1   " & vbCrLf & _
    '                  "   ,CONVERT(DATETIME,PORM.AffiliateApproveDate,120)AffiliateApproveDate,ISNULL(PORM.AffiliateApproveUser,'')AffiliateApproveUser --2   " & vbCrLf & _
    '                  "   ,CONVERT(DATETIME,PORM.PASISendAffiliateDate,120)PASISendAffiliateDate,ISNULL(PORM.PASISendAffiliateUser,'')PASISendAffiliateUser --3   " & vbCrLf & _
    '                  "   ,CONVERT(DATETIME,PORM.SupplierApproveDate,120)SupplierApproveDate,ISNULL(PORM.SupplierApproveUser,'')SupplierApproveUser --4   " & vbCrLf & _
    '                  "   ,CONVERT(DATETIME,PORM.SupplierApprovePendingDate,120)SupplierApprovePendingDate,ISNULL(PORM.SupplierApprovePendingUser,'')SupplierApprovePendingUser --5   " & vbCrLf & _
    '                  "   ,CONVERT(DATETIME,PORM.SupplierUnApproveDate,120)SupplierUnApproveDate,ISNULL(PORM.SupplierUnApproveUser,'')SupplierUnApproveUser --6   " & vbCrLf

    '        ls_SQL = ls_SQL + "   ,CONVERT(DATETIME,PORM.PASIApproveDate,120)PASIApproveDate,ISNULL(PORM.PASIApproveUser ,'')PASIApproveUser --7   " & vbCrLf & _
    '                              "   ,CONVERT(DATETIME,PORM.FinalApproveDate,120)FinalApproveDate,ISNULL(PORM.FinalApproveUser,'')FinalApproveUser --8    " & vbCrLf & _
    '                              "   ,PORD.PartNo,PartName,CASE WHEN MP.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,UnitCls,MOQ,QtyBox,CASE WHEN MP.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,UnitCls,MOQ,QtyBox   " & vbCrLf & _
    '                              "   ,PORM.CurrCls,PORD.Price,PORM.Amount,PORd.CurrCls,PORd.Amount   " & vbCrLf & _
    '                              "   ,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & pDate & "'))),0) " & vbCrLf & _
    '                              "   ,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & pDate & "'))),0) " & vbCrLf & _
    '                              "   ,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & pDate & "'))),0) " & vbCrLf & _
    '                              "   ,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10   " & vbCrLf & _
    '                              "   ,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20   " & vbCrLf & _
    '                              "   ,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31   " & vbCrLf & _
    '                              "   FROM dbo.PORev_Master PORM    " & vbCrLf

    '        ls_SQL = ls_SQL + "   LEFT JOIN dbo.PORev_Detail PORD ON PORM.PORevNo = PORD.PORevNo AND PORM.PONo = PORD.PONo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID " & vbCrLf & _
    '                              "   LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID  " & vbCrLf & _
    '                              "   LEFT JOIN dbo.PO_Detail POD ON PORM.PONo = POD.PONo AND PORM.AffiliateID = POD.AffiliateID AND PORM.SupplierID = POD.SupplierID  " & vbCrLf & _
    '                              "   LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID  " & vbCrLf & _
    '                              "   LEFT JOIN dbo.MS_Parts MP ON PORD.PartNo = MP.PartNo   " & vbCrLf & _
    '                              "   LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID   " & vbCrLf & _
    '                              "   LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID  " & vbCrLf

    '        ls_SQL = ls_SQL + " WHERE --MONTH(PORM.Period) = MONTH('" & pDate & "') AND YEAR(PORM.Period) = YEAR('" & pDate & "') AND " & vbCrLf & _
    '                          " PORM.PORevNo = '" & pPORevNo & "' AND PORM.PONo='" & pPONo & "' AND PORM.AffiliateID='" & pAffCode & "' AND PORM.SupplierID='" & pSupplierID & "'   " & vbCrLf
    '        Dim ds As New DataSet
    '        ds = uf_GetDataSet(ls_SQL)
    '        Return ds
    '    End Function

    '    Private Function BindDataExcel(ByVal pDate As Date, ByVal pAffCode As String, ByVal pPORevNo As String, ByVal pPONo As String, ByVal pSupplierID As String) As DataSet
    '        Dim ls_SQL As String = ""
    '        Dim tanggal As Date = Now
    '        MdlConn.ReadConnection()

    '        ls_SQL = "   	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo    " & vbCrLf & _
    '                  "    	  ,POKanbanCls = KanbanCls ,Description       " & vbCrLf & _
    '                  "          ,MOQ = LEFT(MOQ,LEN(MOQ)-3) , QtyBox = LEFT(QtyBox,LEN(QtyBox)-3) ,Maker      " & vbCrLf & _
    '                  "          ,POQty        " & vbCrLf & _
    '                  "          ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT       " & vbCrLf & _
    '                  "          ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5      " & vbCrLf & _
    '                  "          ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10      " & vbCrLf & _
    '                  "          ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20      " & vbCrLf & _
    '                  "          ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25      " & vbCrLf & _
    '                  "          ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31      " & vbCrLf & _
    '                  "          FROM (       			  " & vbCrLf

    '        ls_SQL = ls_SQL + "  			SELECT CONVERT(CHAR,row_number() over (order by PMU.PONo)) as NoUrut,PDU.PartNo,PDU.PartNo PartNos,PartName ,PMU.PONo        " & vbCrLf & _
    '                          "           	,CASE WHEN MPART.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,MU.DESCRIPTION     " & vbCrLf & _
    '                          "           	,MOQ =CONVERT(CHAR,MOQ),QtyBox = CONVERT(CHAR,QtyBox),ISNULL(MPART.Maker,'')Maker ,PDU.POQty     " & vbCrLf & _
    '                          "    			,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = POD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,POM.period))),0)       " & vbCrLf & _
    '                          "       		,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = POD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,POM.period))),0)       " & vbCrLf & _
    '                          "       		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = POD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,POM.period))),0)                                 " & vbCrLf & _
    '                          "         		,'BEFORE' BYWHAT    " & vbCrLf & _
    '                          "         		,PDU.DeliveryD1 ,PDU.DeliveryD2 ,PDU.DeliveryD3 ,PDU.DeliveryD4 ,PDU.DeliveryD5 ,PDU.DeliveryD6 ,PDU.DeliveryD7 ,PDU.DeliveryD8 ,PDU.DeliveryD9 ,PDU.DeliveryD10     " & vbCrLf & _
    '                          "         		,PDU.DeliveryD11 ,PDU.DeliveryD12 ,PDU.DeliveryD13 ,PDU.DeliveryD14 ,PDU.DeliveryD15 ,PDU.DeliveryD16 ,PDU.DeliveryD17 ,PDU.DeliveryD18 ,PDU.DeliveryD19 ,PDU.DeliveryD20     " & vbCrLf & _
    '                          "         		,PDU.DeliveryD21 ,PDU.DeliveryD22 ,PDU.DeliveryD23 ,PDU.DeliveryD24 ,PDU.DeliveryD25 ,PDU.DeliveryD26 ,PDU.DeliveryD27 ,PDU.DeliveryD28 ,PDU.DeliveryD29 ,PDU.DeliveryD30 ,PDU.DeliveryD31  ,row_number() over (order by PDU.PONo) as Sort         " & vbCrLf & _
    '                          "         		FROM dbo.PO_MasterUpload PMU    " & vbCrLf

    '        ls_SQL = ls_SQL + "     			LEFT JOIN dbo.PO_DetailUpload PDU ON PMU.PONo = PDU.PONo  AND PMU.AffiliateID = PDU.AffiliateID AND PMU.SupplierID = PDU.SupplierID      " & vbCrLf & _
    '                          "     			LEFT JOIN PO_Master POM ON PDU.AffiliateID = POM.AffiliateID AND PDU.PONo = POM.PONo AND PDU.SupplierID = POM.SupplierID      " & vbCrLf & _
    '                          "     			LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
    '                          "     			AND PMU.PONo = POD.PONo  AND PMU.AffiliateID = POD.AffiliateID AND PMU.SupplierID = POD.SupplierID AND POD.PartNo=PDU.PartNo " & vbCrLf & _
    '                          "  				LEFT JOIN dbo.MS_Parts MPART ON PDU.PartNo = MPART.PartNo            " & vbCrLf & _
    '                          "  				LEFT JOIN dbo.MS_Supplier MS ON PDU.SupplierID = MS.SupplierID             " & vbCrLf & _
    '                          "  				LEFT JOIN dbo.MS_Affiliate MA ON PDU.AffiliateID = MA.AffiliateID         " & vbCrLf & _
    '                          "  				LEFT JOIN dbo.MS_SupplierCapacity MSC ON PDU.PartNo = MSC.PartNo AND PDU.SupplierID=MSC.SupplierID             " & vbCrLf & _
    '                          "  				LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls                  " & vbCrLf & _
    '                          "  				LEFT JOIN dbo.MS_CurrCls MCUR1 ON PDU.CurrCls = MCUR1.CurrCls                       " & vbCrLf & _
    '                          "  			WHERE  PMU.PONo='" & pPONo.Trim & "' AND PMU.SupplierID='" & pSupplierID.Trim & "'  " & vbCrLf

    '        ls_SQL = ls_SQL + "  			GROUP BY PMU.PONo,PDU.PONo,PDU.PartNo,PartName,MPART.KanbanCls,MU.DESCRIPTION " & vbCrLf & _
    '                          "  			,MOQ,QtyBox,PDU.poqty,MPART.Maker,MonthlyProductionCapacity          " & vbCrLf & _
    '                          "         		,PDU.CurrCls,MCUR1.Description,PDU.Price,PDU.Amount ,POD.PartNo, POM.Period " & vbCrLf & _
    '                          "         		,PDU.DeliveryD1,PDU.DeliveryD2,PDU.DeliveryD3,PDU.DeliveryD4,PDU.DeliveryD5,PDU.DeliveryD6,PDU.DeliveryD7,PDU.DeliveryD8,PDU.DeliveryD9,PDU.DeliveryD10             " & vbCrLf & _
    '                          "         		,PDU.DeliveryD11,PDU.DeliveryD12,PDU.DeliveryD13,PDU.DeliveryD14,PDU.DeliveryD15,PDU.DeliveryD16,PDU.DeliveryD17,PDU.DeliveryD18,PDU.DeliveryD19,PDU.DeliveryD20        		       " & vbCrLf & _
    '                          "         		,PDU.DeliveryD21,PDU.DeliveryD22,PDU.DeliveryD23,PDU.DeliveryD24,PDU.DeliveryD25,PDU.DeliveryD26,PDU.DeliveryD27,PDU.DeliveryD28,PDU.DeliveryD29,PDU.DeliveryD30,PDU.DeliveryD31        " & vbCrLf & _
    '                          "    	)detail1  " & vbCrLf & _
    '                          "    	UNION ALL      " & vbCrLf & _
    '                          "    	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo ,POKanbanCls = KanbanCls ,Description       " & vbCrLf & _
    '                          "          ,MOQ = MOQ , QtyBox = QtyBox,Maker ,POQty        " & vbCrLf & _
    '                          "          ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT       " & vbCrLf

    '        ls_SQL = ls_SQL + "          ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5      " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10      " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20      " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25      " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31      " & vbCrLf & _
    '                          "          FROM (      " & vbCrLf & _
    '                          "          SELECT row_number() over (order by AD.PONo) as Sort ,'' as NoUrut ,'' PartNo ,AD.PartNo AS PartNos,'' PartName ,'' PONo, '' KanbanCls ,''Description ,'' MOQ,'' QtyBox ,AD.Maker              ,POQty POqty    " & vbCrLf & _
    '                          "      		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,POM.period))),0)       " & vbCrLf & _
    '                          "       		,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,POM.period))),0)       " & vbCrLf & _
    '                          "       		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,POM.period))),0)                                 " & vbCrLf & _
    '                          "       		,'AFTER' BYWHAT     " & vbCrLf

    '        ls_SQL = ls_SQL + "         	,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4, DeliveryD5      " & vbCrLf & _
    '                          "      		,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9, DeliveryD10      " & vbCrLf & _
    '                          "      		,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15      " & vbCrLf & _
    '                          "      		,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20      " & vbCrLf & _
    '                          "      		,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25      " & vbCrLf & _
    '                          "      		,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31        		 " & vbCrLf & _
    '                          "      		FROM dbo.AffiliateRev_Detail AD  " & vbCrLf & _
    '                          "      		LEFT JOIN PO_Master POM ON POM.PONo = AD.PONo AND POM.AffiliateID = AD.AffiliateID AND POM.SupplierID = AD.SupplierID   " & vbCrLf & _
    '                          "      		LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo      " & vbCrLf & _
    '                          "      		LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID             " & vbCrLf & _
    '                          "      		LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID         " & vbCrLf

    '        ls_SQL = ls_SQL + "      		LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID             " & vbCrLf & _
    '                          "      		LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls             " & vbCrLf & _
    '                          "      		LEFT JOIN dbo.MS_CurrCls MCUR1 ON AD.CurrCls = MCUR1.CurrCls                   " & vbCrLf & _
    '                          "            WHERE AD.PORevNo= '" & pPORevNo.Trim & "' AND AD.PONo='" & pPONo.Trim & "' AND AD.SupplierID='" & pSupplierID.Trim & "'     " & vbCrLf & _
    '                          "      		GROUP BY AD.PONo,AD.PartNo,PartName,AD.KanbanCls,POQty,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity,SeqNo,POM.period " & vbCrLf & _
    '                          "      		,MSC.PartNo,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5      " & vbCrLf & _
    '                          "      		,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10      " & vbCrLf & _
    '                          "      		,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15      " & vbCrLf & _
    '                          "      		,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20      " & vbCrLf & _
    '                          "      		,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25      " & vbCrLf & _
    '                          "      		,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31      " & vbCrLf

    '        ls_SQL = ls_SQL + "      	 )detail2    " & vbCrLf & _
    '                          "    	 UNION ALL    " & vbCrLf & _
    '                          "      	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo  ,POKanbanCls = KanbanCls ,DESCRIPTION  " & vbCrLf & _
    '                          "      	,MOQ = MOQ , QtyBox = QtyBox,Maker,POQty,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT       " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5      " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10      " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15   " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20      " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25      " & vbCrLf & _
    '                          "          ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31      " & vbCrLf & _
    '                          "          FROM (      " & vbCrLf

    '        ls_SQL = ls_SQL + "          SELECT row_number() over (order by AD.PONo) as Sort ,'' as NoUrut ,'' PartNo ,AD.PartNo AS PartNos,'' PartName ,'' PONo, '' KanbanCls ,''Description ,'' MOQ,'' QtyBox ,AD.Maker      " & vbCrLf & _
    '                          "  			,POQty POqty  " & vbCrLf & _
    '                          "  			,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,POM.period))),0)       " & vbCrLf & _
    '                          "       		,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,POM.period))),0)       " & vbCrLf & _
    '                          "       		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,POM.period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,POM.period))),0)                                 " & vbCrLf & _
    '                          "       		,'SUPPLIER APPROVAL' BYWHAT     " & vbCrLf & _
    '                          "     			,DeliveryD1 ,DeliveryD2 ,DeliveryD3 ,DeliveryD4 ,DeliveryD5       " & vbCrLf & _
    '                          "     			,DeliveryD6 ,DeliveryD7 ,DeliveryD8 ,DeliveryD9 ,DeliveryD10       " & vbCrLf & _
    '                          "     			,DeliveryD11 ,DeliveryD12 ,DeliveryD13 ,DeliveryD14       " & vbCrLf & _
    '                          "     			,DeliveryD15 ,DeliveryD16 ,DeliveryD17 ,DeliveryD18,DeliveryD19 ,DeliveryD20 ,DeliveryD21       " & vbCrLf & _
    '                          "     			,DeliveryD22 ,DeliveryD23 ,DeliveryD24 ,DeliveryD25 ,DeliveryD26 ,DeliveryD27 ,DeliveryD28 ,DeliveryD29       " & vbCrLf

    '        ls_SQL = ls_SQL + "     			,DeliveryD30 ,DeliveryD31       " & vbCrLf & _
    '                          "     			FROM dbo.AffiliateRev_Detail AD " & vbCrLf & _
    '                          "     			LEFT JOIN PO_Master POM ON POM.PONo = AD.PONo AND POM.AffiliateID = AD.AffiliateID AND POM.SupplierID = AD.SupplierID         		   " & vbCrLf & _
    '                          "     			LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo      " & vbCrLf & _
    '                          "     			LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID      " & vbCrLf & _
    '                          "     			LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID         " & vbCrLf & _
    '                          "     			LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID             " & vbCrLf & _
    '                          "     			LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls      " & vbCrLf & _
    '                          "     			LEFT JOIN dbo.MS_CurrCls MCUR1 ON AD.CurrCls = MCUR1.CurrCls     " & vbCrLf & _
    '                          "  			WHERE AD.PORevNo='" & pPORevNo.Trim & "' AND AD.PONo='" & pPONo.Trim & "' AND AD.SupplierID='" & pSupplierID.Trim & "'     " & vbCrLf & _
    '                          "     			GROUP BY AD.PONo,AD.PartNo,PartName,AD.KanbanCls,POQty,MU.Description,MOQ ,POM.period " & vbCrLf

    '        ls_SQL = ls_SQL + "     			,QtyBox,AD.Maker,MonthlyProductionCapacity ,POQty,MSC.PartNo            		  " & vbCrLf & _
    '                          "     			,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5      " & vbCrLf & _
    '                          "      		,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10      " & vbCrLf & _
    '                          "      		,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15      " & vbCrLf & _
    '                          "      		,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20      " & vbCrLf & _
    '                          "      		,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25      " & vbCrLf & _
    '                          "      		,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31     " & vbCrLf & _
    '                          "    	)detail3    " & vbCrLf & _
    '                          "    	ORDER BY sort, PartNo DESC    "



    '        Dim ds As New DataSet
    '        ds = uf_GetDataSet(ls_SQL)
    '        Return ds
    '    End Function

    '    Private Sub pGetExcelPORev()
    '        On Error GoTo ErrHandler
    '        Dim strFileSize As String = ""

    '        Dim ExcelBook As Excel.Workbook
    '        Dim ExcelSheet As Excel.Worksheet
    '        Dim sheetNumber As Integer = 1
    '        Dim i As Integer
    '        Const ColorYellow As Single = 65535
    '        Dim receiptCCEmail As String = ""
    '        Dim receiptEmail As String = ""
    '        Dim fromEmail As String = ""

    '        'copy file from server to local
    '        Dim fileTocopy As String
    '        Dim NewFileCopy As String
    '        Dim NewFileCopyas As String
    '        Dim pPeriod As Date

    '        Dim ls_SQL As String = ""
    '        Dim ds As New DataSet
    '        Dim dsHeader As New DataSet
    '        Dim dsDetail As New DataSet
    '        Dim dsEta As New DataSet
    '        MdlConn.ReadConnection()
    '        'ls_SQL = "SELECT a.*, b.Period FROM Affiliate_Master a " & vbCrLf & _
    '        '         "inner join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID  = b.SupplierID" & vbCrLf & _
    '        '         "WHERE a.ExcelCls='1'"
    '        ls_SQL = "SELECT a.*, b.Period FROM AffiliateRev_Master a " & vbCrLf & _
    '                 "inner join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID  = b.SupplierID" & vbCrLf & _
    '                 "WHERE ExcelCls='1'"
    '        ds = uf_GetDataSet(ls_SQL)
    '        If ds.Tables(0).Rows.Count > 0 Then
    '            Dim fi As New FileInfo(Trim(txtAttachment.Text) & "\Template PO Revision.xlsm") 'File dari Local

    '            If Not fi.Exists Then
    '                'lblInfo.Text = "Excel Not Exist"
    '                'MsgBox("ga ada excel", MsgBoxStyle.Information)
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, bencause File Excel isn't Found " & vbCrLf & _
    '                                rtbProcess.Text
    '                Exit Sub
    '            End If
    '            NewFileCopy = Trim(txtAttachment.Text) & "\Template PO Revision.xlsm"
    '            'For Each fi In aryFi
    '            Dim xlApp = New Excel.Application
    '            Dim ls_file As String = NewFileCopy

    '            ExcelBook = xlApp.Workbooks.Open(ls_file)
    '            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

    '            pDate = Now
    '            pAffCode = ds.Tables(0).Rows(0)("AffiliateID")
    '            pPORevNo = ds.Tables(0).Rows(0)("PORevNo")
    '            pPONo = ds.Tables(0).Rows(0)("PONo")
    '            pSupplier = ds.Tables(0).Rows(0)("SupplierID")
    '            pPeriod = ds.Tables(0).Rows(0)("Period")

    '            dsHeader = bindDataHeaderPORev(pPeriod, pAffCode, pPORevNo, pPONo, pSupplier)
    '            dsDetail = BindDataExcel(pPeriod, pAffCode, pPORevNo, pPONo, pSupplier)
    '            dsEta = bindDataETA(pAffCode, pSupplier, pPeriod)
    '            If dsHeader.Tables(0).Rows.Count > 0 Then
    '                If dsHeader.Tables(0).Rows(0)("PODeliveryBy") = "1" Then
    '                    pDel = "PASI"
    '                Else
    '                    pDel = dsHeader.Tables(0).Rows(0)("AffiliateID")
    '                End If
    '                Dim dsEmail As New DataSet
    '                dsEmail = EmailToEmailCC(pAffCode, pDel, pSupplier)
    '                '1 CC Affiliate'2 CC PASI'3 CC & TO Supplier
    '                For i = 0 To dsEmail.Tables(0).Rows.Count - 1
    '                    'If receiptCCEmail = "" Then
    '                    '    receiptCCEmail = dsEmail.Tables(0).Rows(i)("affiliatepocc")
    '                    'Else
    '                    '    receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("affiliatepocc")
    '                    'End If
    '                    'receiptEmail = dsEmail.Tables(0).Rows(i)("affiliatepoto")
    '                    If receiptCCEmail = "" Then
    '                        receiptCCEmail = dsEmail.Tables(0).Rows(i)("affiliatepocc")
    '                    Else
    '                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("affiliatepocc")
    '                    End If
    '                    If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
    '                        fromEmail = dsEmail.Tables(0).Rows(i)("toEmail")
    '                    End If
    '                    If receiptEmail = "" Then
    '                        receiptEmail = dsEmail.Tables(0).Rows(i)("affiliatepoto")
    '                    Else
    '                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(i)("affiliatepoto")
    '                    End If
    '                Next
    '                receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '                receiptEmail = Replace(receiptEmail, ",", ";")

    '                ExcelSheet.Range("H1").Value = "POR"
    '                ExcelSheet.Range("H2").Value = fromEmail
    '                ExcelSheet.Range("H3").Value = dsHeader.Tables(0).Rows(0)("AffiliateID")
    '                ExcelSheet.Range("H5").Value = dsHeader.Tables(0).Rows(0)("SupplierID")

    '                ExcelSheet.Range("R8:X8").Merge()
    '                ExcelSheet.Range("R8:X8").Value = "PO REVISION NO:"
    '                ExcelSheet.Range("R8:X8").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
    '                ExcelSheet.Range("R8:X8").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
    '                ExcelSheet.Range("Y8:AF8").Merge()
    '                ExcelSheet.Range("Y8:AF8").Value = Trim(dsHeader.Tables(0).Rows(0)("PORevNo"))

    '                ExcelSheet.Range("I9").Value = dsHeader.Tables(0).Rows(0)("PONo")
    '                ExcelSheet.Range("T9").Value = dsHeader.Tables(0).Rows(0)("PODate")

    '                ExcelSheet.Range("Y2").Value = ""

    '                ExcelSheet.Range("I11").Value = dsHeader.Tables(0).Rows(0)("SupplierName")
    '                Dim dsSupp As New DataSet
    '                dsSupp = Supplier(Trim(pSupplier))
    '                ExcelSheet.Range("I12").Value = dsSupp.Tables(0).Rows(0)("Address")
    '                ExcelSheet.Range("I12:X14").WrapText = True

    '                'Buyer
    '                Dim dsAffp2 As New DataSet
    '                dsAffp2 = Affiliate(Trim(pDel))
    '                ExcelSheet.Range("I16").Value = dsAffp2.Tables(0).Rows(0)("AffiliateName")
    '                ExcelSheet.Range("I17").Value = dsAffp2.Tables(0).Rows(0)("Address")
    '                ExcelSheet.Range("I17:X19").WrapText = True

    '                ExcelSheet.Range("AE9").Value = dsHeader.Tables(0).Rows(0)("PORevDate")
    '                ExcelSheet.Range("AE12").Value = dsHeader.Tables(0).Rows(0)("CommercialCls")
    '                ExcelSheet.Range("AE14").Value = dsHeader.Tables(0).Rows(0)("ShipCls")

    '                'Consignee
    '                Dim dsAffp As New DataSet
    '                dsAffp = Affiliate(Trim(dsHeader.Tables(0).Rows(0)("AffiliateID")))
    '                ExcelSheet.Range("AE16").Value = dsHeader.Tables(0).Rows(0)("AffiliateName")
    '                ExcelSheet.Range("AE17").Value = dsAffp.Tables(0).Rows(0)("Address")
    '                ExcelSheet.Range("AE17:AT19").WrapText = True

    '                pPeriod = Format(dsHeader.Tables(0).Rows(0)("PORevDate"), "MMM-yyyy")
    '                pAffiliateName = dsAffp2.Tables(0).Rows(0)("AffiliateName")
    '                pDelivBy = dsHeader.Tables(0).Rows(0)("PODeliveryBy")
    '                pCommercialRev = dsHeader.Tables(0).Rows(0)("CommercialCls")
    '                pShipRev = dsHeader.Tables(0).Rows(0)("ShipCls")

    '                If dsEta.Tables(0).Rows.Count > 0 Then
    '                    For i = 1 To dsEta.Tables(0).Rows.Count - 1
    '                        ExcelSheet.Range("AX" & i + 33 & ": AY" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day1") '1
    '                        ExcelSheet.Range("AZ" & i + 33 & ": BA" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day2") '2
    '                        ExcelSheet.Range("BB" & i + 33 & ": BC" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day3") '3
    '                        ExcelSheet.Range("BD" & i + 33 & ": BE" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day4") '4
    '                        ExcelSheet.Range("BF" & i + 33 & ": BG" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day5") '5
    '                        ExcelSheet.Range("BH" & i + 33 & ": BI" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day6") '6
    '                        ExcelSheet.Range("BJ" & i + 33 & ": BK" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day7") '7
    '                        ExcelSheet.Range("BL" & i + 33 & ": BM" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day8") '8
    '                        ExcelSheet.Range("BN" & i + 33 & ": BO" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day9") '9
    '                        ExcelSheet.Range("BP" & i + 33 & ": BQ" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day10") '10
    '                        ExcelSheet.Range("BR" & i + 33 & ": BS" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day11") '11
    '                        ExcelSheet.Range("BT" & i + 33 & ": BU" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day12") '12
    '                        ExcelSheet.Range("BV" & i + 33 & ": BW" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day13") '13
    '                        ExcelSheet.Range("BX" & i + 33 & ": BY" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day14") '14
    '                        ExcelSheet.Range("BZ" & i + 33 & ": CA" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day15") '15
    '                        ExcelSheet.Range("CB" & i + 33 & ": CC" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day16") '16
    '                        ExcelSheet.Range("CD" & i + 33 & ": CE" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day17") '17
    '                        ExcelSheet.Range("CF" & i + 33 & ": CG" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day18") '18
    '                        ExcelSheet.Range("CH" & i + 33 & ": CI" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day19") '19
    '                        ExcelSheet.Range("CJ" & i + 33 & ": CK" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day20") '20
    '                        ExcelSheet.Range("CL" & i + 33 & ": CM" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day21") '21
    '                        ExcelSheet.Range("CN" & i + 33 & ": CO" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day22") '22
    '                        ExcelSheet.Range("CP" & i + 33 & ": CQ" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day23") '23
    '                        ExcelSheet.Range("CR" & i + 33 & ": CS" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day24") '24
    '                        ExcelSheet.Range("CT" & i + 33 & ": CU" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day25") '25
    '                        ExcelSheet.Range("CV" & i + 33 & ": CW" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day26") '26
    '                        ExcelSheet.Range("CX" & i + 33 & ": CY" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day27") '27
    '                        ExcelSheet.Range("CZ" & i + 33 & ": DA" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day28") '28
    '                        ExcelSheet.Range("DB" & i + 33 & ": DC" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day29") '29
    '                        ExcelSheet.Range("DD" & i + 33 & ": DE" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day30") '30
    '                        ExcelSheet.Range("DF" & i + 33 & ": DG" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day31") '31
    '                    Next
    '                End If

    '                If dsDetail.Tables(0).Rows.Count > 0 Then
    '                    For i = 0 To dsDetail.Tables(0).Rows.Count - 1
    '                        'If dsDetail.Tables(0).Rows(0)("cols") = "1" Then
    '                        'Header
    '                        ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).Merge()
    '                        ExcelSheet.Range("D" & i + 36 & ": H" & i + 36).Merge()
    '                        ExcelSheet.Range("I" & i + 36 & ": P" & i + 36).Merge()
    '                        ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).Merge()
    '                        ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).Merge()
    '                        ExcelSheet.Range("V" & i + 36 & ": W" & i + 36).Merge()
    '                        ExcelSheet.Range("X" & i + 36 & ": Y" & i + 36).Merge()
    '                        ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Merge()
    '                        ExcelSheet.Range("AC" & i + 36 & ": AE" & i + 36).Merge()
    '                        'ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Merge()
    '                        ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Merge()
    '                        ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Merge()
    '                        ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Merge()
    '                        ExcelSheet.Range("AR" & i + 36 & ": AW" & i + 36).Merge()
    '                        ExcelSheet.Range("AX" & i + 36 & ": AY" & i + 36).Merge() '1
    '                        ExcelSheet.Range("AZ" & i + 36 & ": BA" & i + 36).Merge() '2
    '                        ExcelSheet.Range("BB" & i + 36 & ": BC" & i + 36).Merge() '3
    '                        ExcelSheet.Range("BD" & i + 36 & ": BE" & i + 36).Merge() '4
    '                        ExcelSheet.Range("BF" & i + 36 & ": BG" & i + 36).Merge() '5
    '                        ExcelSheet.Range("BH" & i + 36 & ": BI" & i + 36).Merge() '6
    '                        ExcelSheet.Range("BJ" & i + 36 & ": BK" & i + 36).Merge() '7
    '                        ExcelSheet.Range("BL" & i + 36 & ": BM" & i + 36).Merge() '8
    '                        ExcelSheet.Range("BN" & i + 36 & ": BO" & i + 36).Merge() '9
    '                        ExcelSheet.Range("BP" & i + 36 & ": BQ" & i + 36).Merge() '10
    '                        ExcelSheet.Range("BR" & i + 36 & ": BS" & i + 36).Merge() '11
    '                        ExcelSheet.Range("BT" & i + 36 & ": BU" & i + 36).Merge() '12
    '                        ExcelSheet.Range("BV" & i + 36 & ": BW" & i + 36).Merge() '13
    '                        ExcelSheet.Range("BX" & i + 36 & ": BY" & i + 36).Merge() '14
    '                        ExcelSheet.Range("BZ" & i + 36 & ": CA" & i + 36).Merge() '15
    '                        ExcelSheet.Range("CB" & i + 36 & ": CC" & i + 36).Merge() '16
    '                        ExcelSheet.Range("CD" & i + 36 & ": CE" & i + 36).Merge() '17
    '                        ExcelSheet.Range("CF" & i + 36 & ": CG" & i + 36).Merge() '18
    '                        ExcelSheet.Range("CH" & i + 36 & ": CI" & i + 36).Merge() '19
    '                        ExcelSheet.Range("CJ" & i + 36 & ": CK" & i + 36).Merge() '20
    '                        ExcelSheet.Range("CL" & i + 36 & ": CM" & i + 36).Merge() '21
    '                        ExcelSheet.Range("CN" & i + 36 & ": CO" & i + 36).Merge() '22
    '                        ExcelSheet.Range("CP" & i + 36 & ": CQ" & i + 36).Merge() '23
    '                        ExcelSheet.Range("CR" & i + 36 & ": CS" & i + 36).Merge() '24
    '                        ExcelSheet.Range("CT" & i + 36 & ": CU" & i + 36).Merge() '25
    '                        ExcelSheet.Range("CV" & i + 36 & ": CW" & i + 36).Merge() '26
    '                        ExcelSheet.Range("CX" & i + 36 & ": CY" & i + 36).Merge() '27
    '                        ExcelSheet.Range("CZ" & i + 36 & ": DA" & i + 36).Merge() '28
    '                        ExcelSheet.Range("DB" & i + 36 & ": DC" & i + 36).Merge() '29
    '                        ExcelSheet.Range("DD" & i + 36 & ": DE" & i + 36).Merge() '30
    '                        ExcelSheet.Range("DF" & i + 36 & ": DG" & i + 36).Merge() '31

    '                        ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("NoUrut"))
    '                        ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                        ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
    '                        ExcelSheet.Range("D" & i + 36 & ": H" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
    '                        ExcelSheet.Range("I" & i + 36 & ": P" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName"))
    '                        ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("POKanbanCls"))
    '                        ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                        ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

    '                        ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("Description"))
    '                        ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                        ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

    '                        ExcelSheet.Range("V" & i + 36 & ": W" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("MOQ"))
    '                        ExcelSheet.Range("X" & i + 36 & ": Y" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("QtyBox"))
    '                        ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("Maker"))
    '                        ExcelSheet.Range("AC" & i + 36 & ": AE" & i + 36).Value = If(IsDBNull(dsDetail.Tables(0).Rows(i)("POQty")), 0, dsDetail.Tables(0).Rows(i)("POQty"))
    '                        ExcelSheet.Range("AC" & i + 36 & ": DE" & i + 36).NumberFormat = "#,##0"
    '                        'ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("colapprove"))
    '                        ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Value = dsDetail.Tables(0).Rows(i)("ForecastN1")
    '                        ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).NumberFormat = "#,##0"

    '                        ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Value = dsDetail.Tables(0).Rows(i)("ForecastN2")
    '                        ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).NumberFormat = "#,##0"

    '                        ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Value = dsDetail.Tables(0).Rows(i)("ForecastN3")
    '                        ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).NumberFormat = "#,##0"

    '                        ExcelSheet.Range("AR" & i + 36 & ": AW" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("BYWHAT"))

    '                        If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) = "BEFORE" Then
    '                            ExcelSheet.Range("AG" & i + 36).Value = "YES"
    '                            ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Interior.Color = ColorYellow
    '                        ElseIf Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) = "SUPPLIER APPROVAL" Then
    '                            ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Value = "" 'Maker
    '                            ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Value = "" 'Forecast1
    '                            ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Value = "" 'Forecast1
    '                            ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Value = "" 'Forecast1
    '                            ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).Interior.Color = ColorYellow
    '                        Else
    '                            ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Value = "" 'Maker
    '                            ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Value = "" 'Forecast1
    '                            ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Value = "" 'Forecast1
    '                            ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Value = "" 'Forecast1
    '                        End If

    '                        ExcelSheet.Range("AX" & i + 36 & ": AY" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD1") '1
    '                        ExcelSheet.Range("AZ" & i + 36 & ": BA" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD2") '2
    '                        ExcelSheet.Range("BB" & i + 36 & ": BC" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD3") '3
    '                        ExcelSheet.Range("BD" & i + 36 & ": BE" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD4") '4
    '                        ExcelSheet.Range("BF" & i + 36 & ": BG" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD5") '5
    '                        ExcelSheet.Range("BH" & i + 36 & ": BI" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD6") '6
    '                        ExcelSheet.Range("BJ" & i + 36 & ": BK" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD7") '7
    '                        ExcelSheet.Range("BL" & i + 36 & ": BM" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD8") '8
    '                        ExcelSheet.Range("BN" & i + 36 & ": BO" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD9") '9
    '                        ExcelSheet.Range("BP" & i + 36 & ": BQ" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD10") '10
    '                        ExcelSheet.Range("BR" & i + 36 & ": BS" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD11") '11
    '                        ExcelSheet.Range("BT" & i + 36 & ": BU" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD12") '12
    '                        ExcelSheet.Range("BV" & i + 36 & ": BW" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD13") '13
    '                        ExcelSheet.Range("BX" & i + 36 & ": BY" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD14") '14
    '                        ExcelSheet.Range("BZ" & i + 36 & ": CA" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD15") '15
    '                        ExcelSheet.Range("CB" & i + 36 & ": CC" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD16") '16
    '                        ExcelSheet.Range("CD" & i + 36 & ": CE" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD17") '17
    '                        ExcelSheet.Range("CF" & i + 36 & ": CG" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD18") '18
    '                        ExcelSheet.Range("CH" & i + 36 & ": CI" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD19") '19
    '                        ExcelSheet.Range("CJ" & i + 36 & ": CK" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD20") '20
    '                        ExcelSheet.Range("CL" & i + 36 & ": CM" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD21") '21
    '                        ExcelSheet.Range("CN" & i + 36 & ": CO" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD22") '22
    '                        ExcelSheet.Range("CP" & i + 36 & ": CQ" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD23") '23
    '                        ExcelSheet.Range("CR" & i + 36 & ": CS" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD24") '24
    '                        ExcelSheet.Range("CT" & i + 36 & ": CU" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD25") '25
    '                        ExcelSheet.Range("CV" & i + 36 & ": CW" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD26") '26
    '                        ExcelSheet.Range("CX" & i + 36 & ": CY" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD27") '27
    '                        ExcelSheet.Range("CZ" & i + 36 & ": DA" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD28") '28
    '                        ExcelSheet.Range("DB" & i + 36 & ": DC" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD29") '29
    '                        ExcelSheet.Range("DD" & i + 36 & ": DE" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD30") '30
    '                        ExcelSheet.Range("DF" & i + 36 & ": DG" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD31") '31
    '                        ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).NumberFormat = "#,##0"
    '                        ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                        ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
    '                        DrawAllBorders(ExcelSheet.Range("B" & i + 36 & ": AE" & i + 36))
    '                        DrawAllBorders(ExcelSheet.Range("AI" & i + 36 & ": DG" & i + 36))
    '                        ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                        ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                    Next
    '                End If

    '                ExcelSheet.Range("B39").Interior.Color = Color.White
    '                ExcelSheet.Range("B39").Font.Color = Color.Black
    '                ExcelSheet.Range("B" & i + 36).Value = "E"
    '                ExcelSheet.Range("B" & i + 36).Interior.Color = Color.Black
    '                ExcelSheet.Range("B" & i + 36).Font.Color = Color.White

    '                'ExcelBook.SaveAs("D:\PASI\Source Code Terakhir\PASISystem\PASISystem\Template\PO.xlsm")\

    '                'Save ke Server
    '                'ExcelBook.SaveAs(Server.MapPath("~\Template\Result\PO.xlsm"))

    '                'Save ke Local
    '                xlApp.DisplayAlerts = False
    '                'ExcelBook.SaveAs("D:\PASI EBWEB\PASISystem\Template\PO.xlsm")
    '                ExcelBook.SaveAs(Trim(txtSaveAs.Text) & "\PO Revision " & Trim(pPORevNo) & ".xlsm")
    '                'ExcelBook.Save()
    '                'ExcelBook.SaveAs(Server.MapPath("~\Template\PO.xlsm"))
    '                'Dim fStream As New FileStream("c:\data.xls", FileMode.Create)
    '                'ExcelBook.SaveAs(Server.MapPath("~\Template\Result\PO.xlsm"))
    '                'If System.IO.File.Exists("D:\Template\Result\PO.xlsm") = True Then
    '                '    'System.IO.File.Delete(NewFileCopy)
    '                '    System.IO.File.Copy(fileTocopy, NewFileCopy)
    '                'Else
    '                '    System.IO.File.Copy(fileTocopy, NewFileCopy)
    '                'End If


    '                '*****Copy Excel Local ke Server
    '                'If System.IO.File.Exists(fileTocopy) = True Then
    '                '    System.IO.File.Delete(fileTocopy)
    '                '    System.IO.File.Copy(NewFileCopy, fileTocopy)
    '                'Else
    '                '    System.IO.File.Copy(NewFileCopy, fileTocopy)
    '                'End If

    '                'System.IO.File.Delete(NewFileCopy)

    '                xlApp.Workbooks.Close()

    '                xlApp.Quit()

    '                'Call sendEmailPORev()
    '                Call sendEmailPORevisiontoSupllier()
    '                Call sendEmailPORevisiontoAffiliate()
    '                Call sendEmailPORevisionccPASI()
    '                'System.IO.File.Delete("D:\Template\PO.xlsm")
    '                Call UpdateExcelPORev(True, pAffCode, pPONo, pSupplier)
    '                'rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier SUCCESFULL" & Format(Now, "HH:mm:ss") & vbCrLf & _
    '                '                rtbProcess.Text

    '            End If

    'ErrHandler:
    '            'MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    '            xlApp.Workbooks.Close()
    '            xlApp.Quit()

    '        Else
    '            'xlApp.Workbooks.Close()

    '            'xlApp.Quit()
    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because there is nothing PO Revision to send " & vbCrLf & _
    '                            rtbProcess.Text

    '        End If

    '        Exit Sub
    '    End Sub

    '    Private Sub sendEmailPORev()
    '        Try
    '            Dim TempFilePath As String
    '            Dim TempFileName As String
    '            Dim receiptEmail As String = ""
    '            Dim receiptCCEmail As String = ""
    '            Dim fromEmail As String = ""


    '            'TempFilePath = Trim(txtAttachment.Text)
    '            TempFilePath = Trim(txtSaveAs.Text)

    '            '*******File di Server
    '            'TempFilePath = Server.MapPath("~/Template/")
    '            'TempFileName = "PO.xlsm"

    '            'File di Local
    '            'TempFilePath = "D:\Template\"
    '            'TempFileName = "PO.xlsm"

    '            TempFileName = "\PO Revision " & Trim(pPORevNo) & ".xlsm"

    '            'receiptEmail = "kristriyana@tos.co.id"
    '            'receiptEmail = "kris.trieyana@gmail.com"

    '            Dim dsEmail As New DataSet
    '            dsEmail = EmailToEmailCC(pAffCode, pDel, pSupplier)
    '            '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '                If receiptCCEmail = "" Then
    '                    receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '                Else
    '                    receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '                End If
    '                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
    '                End If
    '                If receiptEmail = "" Then
    '                    receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '                Else
    '                    receiptEmail = receiptEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '                End If
    '                'receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '            Next
    '            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '            receiptEmail = Replace(receiptEmail, ",", ";")

    '            receiptCCEmail = Replace(receiptCCEmail, " ", "")
    '            receiptEmail = Replace(receiptEmail, " ", "")
    '            fromEmail = Replace(fromEmail, " ", "")


    '            If receiptEmail = "" Then
    '                'MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because Recipient's e-mail address is not found " & vbCrLf & _
    '                                rtbProcess.Text
    '                Exit Sub
    '            End If

    '            If fromEmail = "" Then
    '                'MsgBox("Mailer's e-mail address is not found", vbCritical, "Warning")
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because Mailer's e-mail address is not found " & vbCrLf & _
    '                                rtbProcess.Text
    '                Exit Sub
    '            End If

    '            'Make a copy of the file/Open it/Mail it/Delete it
    '            'If you want to change the file name then change only TempFileName


    '            'Dim mailMessage As New Mail.MailMessage(fromEmail, receiptEmail)
    '            Dim mailMessage As New Mail.MailMessage()
    '            mailMessage.From = New MailAddress(fromEmail)
    '            mailMessage.Subject = "[TRIAL] PO Revision Template Testing " & pPORevNo & ""

    '            If receiptEmail <> "" Then
    '                For Each recipient In receiptEmail.Split(";"c)
    '                    If recipient <> "" Then
    '                        Dim mailAddress As New MailAddress(recipient)
    '                        mailMessage.To.Add(mailAddress)
    '                    End If
    '                Next
    '            End If
    '            If receiptCCEmail <> "" Then
    '                For Each recipientCC In receiptCCEmail.Split(";"c)
    '                    If recipientCC <> "" Then
    '                        Dim mailAddress As New MailAddress(recipientCC)
    '                        mailMessage.CC.Add(mailAddress)
    '                    End If
    '                Next
    '            End If
    '            GetSettingEmail("PO Revision")
    '            uf_GetNotification("21")
    '            ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
    '            mailMessage.Body = ls_Body
    '            Dim filename As String = TempFilePath & TempFileName
    '            mailMessage.Attachments.Add(New Attachment(filename))
    '            mailMessage.IsBodyHtml = False
    '            Dim smtp As New SmtpClient
    '            'smtp.Host = "smtp.atisicloud.com"
    '            'smtp.Host = "mail.fast.net.id"
    '            'smtp.EnableSsl = False
    '            'smtp.UseDefaultCredentials = True
    '            smtp.Host = smtpClient
    '            If smtp.UseDefaultCredentials = True Then
    '                smtp.EnableSsl = True
    '            Else
    '                smtp.EnableSsl = False
    '                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '                smtp.Credentials = myCredential
    '            End If
    '            smtp.Port = portClient
    '            smtp.Send(mailMessage)

    '            'Delete the file
    '            'Kill(TempFilePath & TempFileName)
    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier SUCCESSFULL" & vbCrLf & _
    '                            rtbProcess.Text

    '        Catch ex As Exception
    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & "" & vbCrLf & _
    '                            rtbProcess.Text

    '        End Try

    '    End Sub

    '    Private Sub UpdateExcelPORev(ByVal pIsNewData As Boolean, _
    '                         Optional ByVal pAffCode As String = "", _
    '                         Optional ByVal pPONo As String = "", _
    '                         Optional ByVal pSuppCode As String = "")

    '        Dim ls_SQL As String = "", ls_MsgID As String = ""
    '        Dim admin As String = "administrator"

    '        Try
    '            MdlConn.ReadConnection()
    '            Using sqlConn As New SqlConnection(MdlConn.uf_GetConString)
    '                sqlConn.Open()
    '                ls_SQL = " UPDATE dbo.AffiliateRev_Master " & vbCrLf & _
    '                      " SET ExcelCls='2'" & vbCrLf & _
    '                      " WHERE PONo='" & pPONo & "'  " & vbCrLf & _
    '                      " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
    '                      " AND SupplierID='" & pSuppCode & "' "
    '                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
    '                sqlComm.ExecuteNonQuery()
    '                sqlComm.Dispose()
    '                sqlConn.Close()
    '            End Using

    '        Catch ex As Exception
    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
    '                            rtbProcess.Text

    '        End Try
    '    End Sub

    '    Private Sub sendEmailPORevisiontoSupllier()
    '        Try
    '            Dim TempFilePath As String
    '            Dim TempFileName As String
    '            Dim receiptEmail As String = ""
    '            Dim receiptCCEmail As String = ""
    '            Dim fromEmail As String = ""


    '            TempFilePath = Trim(txtSaveAs.Text)

    '            TempFileName = "\PO Revision " & Trim(pPORevNo) & ".xlsm"

    '            Dim dsEmail As New DataSet
    '            dsEmail = EmailToEmailCCPORev(pAffCode, pDel, pSupplier)
    '            'To Supplier, CC Supplier
    '            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
    '                End If
    '                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
    '                    If receiptEmail = "" Then
    '                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionTo")
    '                    Else
    '                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionTo")
    '                    End If
    '                End If
    '                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
    '                    If receiptCCEmail = "" Then
    '                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionCC")
    '                    Else
    '                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionCC")
    '                    End If
    '                End If
    '            Next
    '            receiptCCEmail = Replace(receiptCCEmail, " ", "")
    '            receiptEmail = Replace(receiptEmail, " ", "")
    '            fromEmail = Replace(fromEmail, " ", "")

    '            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '            receiptEmail = Replace(receiptEmail, ",", ";")

    '            If receiptEmail = "" Then
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because Recipient's e-mail address is not found " & vbCrLf & _
    '                                rtbProcess.Text
    '                Exit Sub
    '            End If

    '            If fromEmail = "" Then
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because Mailer's e-mail address is not found " & vbCrLf & _
    '                                rtbProcess.Text
    '                Exit Sub
    '            End If

    '            Dim mailMessage As New Mail.MailMessage()
    '            mailMessage.From = New MailAddress(fromEmail)
    '            mailMessage.Subject = "[TRIAL] PO Revision Template Testing " & pPORevNo & ""

    '            If receiptEmail <> "" Then
    '                For Each recipient In receiptEmail.Split(";"c)
    '                    If recipient <> "" Then
    '                        Dim mailAddress As New MailAddress(recipient)
    '                        mailMessage.To.Add(mailAddress)
    '                    End If
    '                Next
    '            End If
    '            If receiptCCEmail <> "" Then
    '                For Each recipientCC In receiptCCEmail.Split(";"c)
    '                    If recipientCC <> "" Then
    '                        Dim mailAddress As New MailAddress(recipientCC)
    '                        mailMessage.CC.Add(mailAddress)
    '                    End If
    '                Next
    '            End If
    '            GetSettingEmail("PO Revision")
    '            uf_GetNotification("21")
    '            ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
    '            mailMessage.Body = ls_Body
    '            Dim filename As String = TempFilePath & TempFileName
    '            mailMessage.Attachments.Add(New Attachment(filename))
    '            mailMessage.IsBodyHtml = False
    '            Dim smtp As New SmtpClient
    '            smtp.Host = smtpClient
    '            If smtp.UseDefaultCredentials = True Then
    '                smtp.EnableSsl = True
    '            Else
    '                smtp.EnableSsl = False
    '                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '                smtp.Credentials = myCredential
    '            End If
    '            smtp.Port = portClient
    '            smtp.Send(mailMessage)

    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier SUCCESSFULL" & vbCrLf & _
    '                            rtbProcess.Text

    '        Catch ex As Exception
    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & "" & vbCrLf & _
    '                            rtbProcess.Text

    '        End Try

    '    End Sub

    '    Private Sub sendEmailPORevisionccPASI() 'Link Affiliate Order Revision Entry
    '        Try
    '            Dim TempFilePath As String
    '            Dim TempFileName As String
    '            Dim receiptEmail As String = ""
    '            Dim receiptCCEmail As String = ""
    '            Dim fromEmail As String = ""

    '            TempFilePath = Trim(txtSaveAs.Text)
    '            TempFileName = "\PO " & Trim(pPONo) & ".xlsm"

    '            Dim dsEmail As New DataSet
    '            dsEmail = EmailToEmailCCPORev(pAffCode, pDel, pSupplier)
    '            '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
    '                    If receiptCCEmail = "" Then
    '                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionCC")
    '                    Else
    '                        receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionCC")
    '                    End If
    '                    If receiptEmail = "" Then
    '                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionTo")
    '                    Else
    '                        receiptEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionTo")
    '                    End If
    '                End If
    '            Next
    '            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '            receiptEmail = Replace(receiptEmail, ",", ";")

    '            receiptCCEmail = Replace(receiptCCEmail, " ", "")
    '            receiptEmail = Replace(receiptEmail, " ", "")
    '            fromEmail = Replace(fromEmail, " ", "")

    '            If receiptEmail = "" Then
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
    '                                rtbProcess.Text
    '                Exit Sub
    '            End If

    '            If fromEmail = "" Then
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
    '                                rtbProcess.Text
    '                Exit Sub
    '            End If

    '            Dim mailMessage As New Mail.MailMessage()
    '            mailMessage.From = New MailAddress(fromEmail)
    '            mailMessage.Subject = "[TRIAL] Send To Supplier PO Rev No: " & pPORevNo & ""

    '            If receiptEmail <> "" Then
    '                For Each recipient In receiptEmail.Split(";"c)
    '                    If recipient <> "" Then
    '                        Dim mailAddress As New MailAddress(recipient)
    '                        mailMessage.To.Add(mailAddress)
    '                    End If
    '                Next
    '            End If
    '            If receiptCCEmail <> "" Then
    '                For Each recipientCC In receiptCCEmail.Split(";"c)
    '                    If recipientCC <> "" Then
    '                        Dim mailAddress As New MailAddress(recipientCC)
    '                        mailMessage.CC.Add(mailAddress)
    '                    End If
    '                Next
    '            End If
    '            GetSettingEmail("PO Rev")
    '            'uf_GetNotification("11")
    '            'ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
    '            '"AffiliateOrderevEntry.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetPeriod(Container)%>
    '            '&t2=<%#GetPORevNo(Container)%>&t3=<%#GetPONo(Container)%>
    '            '&t4=<%#GetCommercial(Container)%>&t5=<%#GetAffiliateID(Container)%>
    '            '&t6=<%#GetAffiliateName(Container)%>&t7=<%#GetSupplierID(Container)%>
    '            '&t8=<%#GetSupplierName(Container)%>&t9=<%#GetKanban(Container)%>
    '            '&t10=<%#GetShip(Container)%>>&t11=<%#GetSeq(Container)%>&Session=~/AffiliateRevision/AffiliateOrderRevList.aspx"
    '            Dim ls_URl As String = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateRevision/AffiliateOrderevEntry.aspx?id2=" & clsNotification.EncryptURL(pPONo.Trim) & _
    '                "&t1=" & clsNotification.EncryptURL(pPeriod) & "&t2=" & clsNotification.EncryptURL(pPORevNo.Trim) & _
    '                "&t3=" & clsNotification.EncryptURL(pPONo.Trim) & "t4=" & clsNotification.EncryptURL(pCommercialRev.Trim) & _
    '                "&t5=" & clsNotification.EncryptURL(pAffCode.Trim) & "&t6=" & clsNotification.EncryptURL(pAffiliateName) & _
    '                "&t7=" & clsNotification.EncryptURL(pSupplier.Trim) & "&t8=" & clsNotification.EncryptURL("2") & _
    '                "&t9=" & clsNotification.EncryptURL(pShipRev.Trim) & "&t10=" & clsNotification.EncryptURL("") & _
    '                "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderList.aspx")

    '            ls_Body = clsNotification.GetNotification("21", ls_URl, pPORevNo.Trim)

    '            mailMessage.Body = ls_Body
    '            mailMessage.IsBodyHtml = False
    '            Dim smtp As New SmtpClient
    '            smtp.Host = smtpClient
    '            If smtp.UseDefaultCredentials = True Then
    '                smtp.EnableSsl = True
    '            Else
    '                smtp.EnableSsl = False
    '                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '                smtp.Credentials = myCredential
    '            End If

    '            smtp.Port = portClient
    '            smtp.Send(mailMessage)

    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier SUCCESSFULL" & vbCrLf & _
    '                             rtbProcess.Text
    '        Catch ex As Exception
    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
    '                            rtbProcess.Text

    '        End Try

    '    End Sub

    '    Private Sub sendEmailPORevisiontoAffiliate() 'Link PO Revision Entry
    '        Try
    '            Dim receiptEmail As String = ""
    '            Dim receiptCCEmail As String = ""
    '            Dim fromEmail As String = ""

    '            Dim dsEmail As New DataSet
    '            dsEmail = EmailToEmailCCPORev(pAffCode, pDel, pSupplier)
    '            '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
    '                End If
    '                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
    '                    If receiptEmail = "" Then
    '                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionTo")
    '                    Else
    '                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionTo")
    '                    End If
    '                End If
    '                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
    '                    If receiptCCEmail = "" Then
    '                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionCC")
    '                    Else
    '                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionCC")
    '                    End If
    '                End If
    '            Next

    '            receiptCCEmail = Replace(receiptCCEmail, " ", "")
    '            receiptEmail = Replace(receiptEmail, " ", "")
    '            fromEmail = Replace(fromEmail, " ", "")

    '            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '            receiptEmail = Replace(receiptEmail, ",", ";")



    '            If receiptEmail = "" Then
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
    '                                rtbProcess.Text

    '                Exit Sub
    '            End If

    '            If fromEmail = "" Then
    '                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Revision to Supplier STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
    '                                rtbProcess.Text
    '                Exit Sub
    '            End If

    '            Dim mailMessage As New Mail.MailMessage()
    '            mailMessage.From = New MailAddress(fromEmail)
    '            mailMessage.Subject = "[TRIAL] Send To Supplier PO Rev" & pPORevNo & ""

    '            If receiptEmail <> "" Then
    '                For Each recipient In receiptEmail.Split(";"c)
    '                    If recipient <> "" Then
    '                        Dim mailAddress As New MailAddress(recipient)
    '                        mailMessage.To.Add(mailAddress)
    '                    End If
    '                Next
    '            End If
    '            If receiptCCEmail <> "" Then
    '                For Each recipientCC In receiptCCEmail.Split(";"c)
    '                    If recipientCC <> "" Then
    '                        Dim mailAddress As New MailAddress(recipientCC)
    '                        mailMessage.CC.Add(mailAddress)
    '                    End If
    '                Next
    '            End If
    '            GetSettingEmail("PO Rev")
    '            '"PORevEntry.aspx?id=<%#GetRowValue(Container)%>
    '            '&t1=<%#GetAffiliateID(Container)%>&t2=<%#GetAffiliateName(Container)%>
    '            '&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%>
    '            '&t5=<%#GetPORevNo(Container)%>&Session=~/PurchaseOrderRevision/PORevList.aspx"

    '            Dim ls_URl As String = "http://" & clsNotification.pub_ServerName & "/PurchaseOrderRevision/PORevEntry.aspx?id2=" & clsNotification.EncryptURL(pPONo.Trim) & _
    '                "&t1=" & clsNotification.EncryptURL(pAffCode) & "&t2=" & clsNotification.EncryptURL(pAffiliateName.Trim) & _
    '                "&t3=" & clsNotification.EncryptURL(pPeriod.Trim) & "&t4=" & clsNotification.EncryptURL(pSupplier.Trim) & _
    '                "&t5=" & clsNotification.EncryptURL(pPORevNo.Trim) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrderRevision/PORevFinalApprovalList.aspx")

    '            ls_Body = clsNotification.GetNotification("21", ls_URl, pPONo.Trim, "", "", pPORevNo.Trim)

    '            mailMessage.Body = ls_Body

    '            mailMessage.IsBodyHtml = False
    '            Dim smtp As New SmtpClient
    '            smtp.Host = smtpClient
    '            If smtp.UseDefaultCredentials = True Then
    '                smtp.EnableSsl = True
    '            Else
    '                smtp.EnableSsl = False
    '                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '                smtp.Credentials = myCredential
    '            End If

    '            smtp.Port = portClient
    '            smtp.Send(mailMessage)

    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier SUCCESSFULL" & vbCrLf & _
    '                             rtbProcess.Text
    '        Catch ex As Exception
    '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
    '                            rtbProcess.Text

    '        End Try

    '    End Sub

    '    Private Function EmailToEmailCCPORev(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
    '        Dim ls_SQL As String = ""
    '        MdlConn.ReadConnection()
    '        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
    '                " select 'AFF' flag,AffiliatePORevisionCC,AffiliatePORevisionTo,toEmail = AffiliatePORevisionTo  from ms_emailAffiliate where AffiliateID='" & pAfffCode & "'" & vbCrLf & _
    '                " union all " & vbCrLf & _
    '                " --PASI TO -CC " & vbCrLf & _
    '                " select 'PASI' flag,AffiliatePORevisionCC,AffiliatePORevisionTo,toEmail = AffiliatePORevisionTo  from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
    '                " union all " & vbCrLf & _
    '                " --Supplier TO- CC " & vbCrLf & _
    '                " select 'SUPP' flag,AffiliatePORevisionCC,AffiliatePORevisionTo,toEmail='' from ms_emailSupplier where SupplierID='" & Trim(pSupplierID) & "'"
    '        Dim ds As New DataSet
    '        ds = uf_GetDataSet(ls_SQL)

    '        If ds.Tables(0).Rows.Count > 0 Then
    '            Return ds
    '        End If
    '    End Function
End Class
