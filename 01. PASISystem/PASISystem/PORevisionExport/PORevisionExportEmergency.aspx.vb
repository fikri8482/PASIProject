Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports OfficeOpenXml
Imports Microsoft.Office.Interop
Imports System.Net
Imports System.Net.Mail
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO

Public Class PORevisionExportEmergency
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "B04"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_AffiliateID As String, pub_AffiliateName As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String, pub_SupplierName As String, pub_Remarks As String
    Dim pub_FinalApproval As String, pub_DeliveyBy As String
    Dim pub_Period As Date
    Dim pub_HeijunkaSttus As Boolean

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String

    Dim flag As Boolean = True

    Dim pStatus As Boolean

    Dim pPeriod As String
    Dim pCommercial As String
    Dim pDeliveryCode As String
    Dim pDeliveryName As String
    Dim pPOEmergency As String
    Dim pShipBy As String
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pSupplierCode As String
    Dim pSupplierName As String
    Dim pPORevNo As String
    Dim pPO As String

    Dim pFilter As String
    Dim pub_Param As String
    Dim pstatusInsert As String
#End Region

#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim param As String = ""
        Dim filterQty As String = ""


        Try
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                'If Not IsNothing(Request.QueryString("prm")) Then
                Session("MenuDesc") = "DELIVERY TO AFFILIATE ENTRY"

                If Session("PORevExportList") <> "" Then
                    param = Session("PORevExportList").ToString()
                ElseIf Session("TampungDelivery") <> "" Then
                    param = Session("TampungDelivery").ToString()
                Else
                    param = Request.QueryString("prm").ToString
                End If

                If param = "  'back'" Then
                    btnSubMenu.Text = "BACK"
                Else
                    If pStatus = False Then
                        pPORevNo = Split(param, "|")(0)
                        pAffiliateCode = Split(param, "|")(1)
                        pAffiliateName = Split(param, "|")(2)
                        pSupplierCode = Split(param, "|")(3)
                        pSupplierName = Split(param, "|")(4)
                        pDeliveryCode = Split(param, "|")(5)
                        pDeliveryName = Split(param, "|")(6)
                        pCommercial = Split(param, "|")(7)
                        pPOEmergency = Split(param, "|")(8)
                        pShipBy = Split(param, "|")(9)

                        If pAffiliateCode <> "" Then btnSubMenu.Text = "BACK"
                        If Trim(pPeriod) = "01 Jan 1900" Then pPeriod = Format(Now, "dd MMM yyyy")
                        If Trim(pPeriod) = "" Then pPeriod = Format(Now, "dd MMM yyyy")
                        dtPeriodFrom.Text = pPeriod
                        rblCommercial.Value = pCommercial
                        cboAffiliateCode.Text = pAffiliateCode
                        txtAffiliateName.Text = pAffiliateName
                        cboDeliveryLoc.Text = pDeliveryCode
                        txtDeliveryLoc.Text = pDeliveryName
                        cboSupplierCode.Text = pSupplierCode
                        txtSupplierName.Text = pSupplierName
                        txtRevisionNo.Text = pPORevNo
                        pStatus = True

                        Session("pFilter") = pFilter
                        Session.Remove("POList")
                    End If
                End If
                btnSubMenu.Text = "BACK"
                'End If
            End If
            '===============================================================================

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblInfo.Text = ""
                'dt1.Value = Format(txtkanbandate.text, "MMM yyyy")
            End If

            'Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            grid.JSProperties("cpMessage") = lblInfo.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        If Session("M01Url") <> "" Then
            Session.Remove("M01Url")
            Session.Remove("Mode")
            'Session.Remove("SupplierID")
            Response.Redirect("~/PORevisionExport/PORevisionExportList.aspx")
        Else
            Session.Remove("M01Url")
            Session.Remove("Mode")
            'Session.Remove("SupplierID")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub up_Fillcombo()
        Dim ls_SQL As String = ""
        'Combo Affiliate
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' AffiliateID, '" & clsGlobal.gs_All & "' AffiliateName UNION ALL SELECT RTRIM(AffiliateID) AffiliateID,AffiliateName FROM dbo.MS_Affiliate" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 50
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 120

                .TextField = "AffiliateID"
                .DataBind()
                .SelectedIndex = 0
                txtAffiliateName.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataHeader(ByVal pDate As Date, ByVal pPORevNo As String, ByVal pPOEmergency As String, ByVal pAffCode As String, ByVal pSupplierID As String, ByVal pDelivery As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT POM.OrderNo1,POM.OrderNo2,POM.OrderNo3,POM.OrderNo4,POM.OrderNo5 " & vbCrLf & _
                  " ,POM.ETDVendor1 AS ETDVendorOld1,POM.ETDVendor2 AS ETDVendorOld2,POM.ETDVendor3 AS ETDVendorOld3,POM.ETDVendor4 AS ETDVendorOld4,POM.ETDVendor5 AS ETDVendorOld5 " & vbCrLf & _
                  " ,PORM.ETDVendor1,PORM.ETDVendor2,PORM.ETDVendor3,PORM.ETDVendor4,PORM.ETDVendor5 " & vbCrLf & _
                  " ,POM.ETDPort1 AS ETDPortOld1,POM.ETDPort2 AS ETDPortOld2,POM.ETDPort3 AS ETDPortOld3,POM.ETDPort4 AS ETDPortOld4,POM.ETDPort5 AS ETDPortOld5 " & vbCrLf & _
                  " ,PORM.ETDPort1,PORM.ETDPort2,PORM.ETDPort3,PORM.ETDPort4,PORM.ETDPort5 " & vbCrLf & _
                  " ,POM.ETAPort1 AS ETAPortOld1,POM.ETAPort2 AS ETAPortOld2,POM.ETAPort3 AS ETAPortOld3,POM.ETAPort4 AS ETDPortOld4,POM.ETAPort5 AS ETDPortOld5 " & vbCrLf & _
                  " ,PORM.ETAPort1,PORM.ETAPort2,PORM.ETAPort3,PORM.ETAPort4,PORM.ETAPort5 " & vbCrLf & _
                  " ,POM.ETAFactory1 AS ETAFactoryOld1,POM.ETAFactory2 AS ETAFactoryOld2,POM.ETAFactory3 AS ETAFactoryOld3,POM.ETAFactory4 AS ETAFactoryOld4,POM.ETAFactory5 AS ETAFactoryOld5 " & vbCrLf & _
                  " ,PORM.ETAFactory1,PORM.ETAFactory2,PORM.ETAFactory3,PORM.ETAFactory4,PORM.ETAFactory5 " & vbCrLf & _
                  " FROM PO_Master_Export POM  " & vbCrLf & _
                  " LEFT JOIN PORev_Master_Export PORM ON POM.PONo = PORM.PONo AND POM.AffiliateID = PORM.AffiliateID  "

            ls_SQL = ls_SQL + " AND POM.SupplierID = PORM.SupplierID " & vbCrLf & _
                              " WHERE MONTH(PORM.Period) = MONTH('" & pDate & "') AND YEAR(PORM.Period) = YEAR('" & pDate & "')" & vbCrLf & _
                              " AND POM.EmergencyCls=CASE WHEN '" & Trim(pPOEmergency) & "'='E' THEN '1' ELSE '0' END --AND POM.ForwarderID='" & Trim(pDelivery) & "' AND POM.AffiliateID ='" & Trim(pAffCode) & "' AND POM.SupplierID='" & Trim(pSupplierID) & "' "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then

                grid.JSProperties("cpOrder1") = ds.Tables(0).Rows(0)("OrderNo1")
                grid.JSProperties("cpOrder1") = ds.Tables(0).Rows(0)("OrderNo1")
                grid.JSProperties("cpOrder1") = ds.Tables(0).Rows(0)("OrderNo1")
                grid.JSProperties("cpOrder1") = ds.Tables(0).Rows(0)("OrderNo1")
                grid.JSProperties("cpOrder1") = ds.Tables(0).Rows(0)("OrderNo1")
                grid.JSProperties("cpOrder1") = ds.Tables(0).Rows(0)("OrderNo1")
                grid.JSProperties("cpOrder2") = ds.Tables(0).Rows(0)("OrderNo2")
                grid.JSProperties("cpOrder3") = ds.Tables(0).Rows(0)("OrderNo3")
                grid.JSProperties("cpOrder4") = ds.Tables(0).Rows(0)("OrderNo4")
                grid.JSProperties("cpOrder5") = ds.Tables(0).Rows(0)("OrderNo5")
                grid.JSProperties("ETDVendorOld1") = If(IsDBNull(ds.Tables(0).Rows(0)("ETDVendorOld1")), "", Format(ds.Tables(0).Rows(0)("ETDVendorOld1"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("ETDVendorOld1") = If(IsDBNull(ds.Tables(0).Rows(0)("ETDVendorOld1")), "", Format(ds.Tables(0).Rows(0)("ETDVendorOld1"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpEntryUser") = ds.Tables(0).Rows(0)("EntryUser")
                grid.JSProperties("cpAffAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")), "", Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpAffAppUser") = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                grid.JSProperties("cpSendDate") = If(IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")), "", Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpSendUser") = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                grid.JSProperties("cpSuppAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpSuppAppUser") = ds.Tables(0).Rows(0)("SupplierApproveUser")
                grid.JSProperties("cpSuppAppPendingDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpSuppAppPendingUser") = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                grid.JSProperties("cpSuppUnApproveDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpSuppUnApproveUser") = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                grid.JSProperties("cpPASIAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")), "", Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpPASIAppUser") = ds.Tables(0).Rows(0)("PASIApproveUser")
                grid.JSProperties("cpFinalAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")), "", Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                grid.JSProperties("cpFinalAppUser") = ds.Tables(0).Rows(0)("FinalApproveUser")

                Call clsMsg.DisplayMessage(lblInfo, "1008", clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                Session("YA010IsSubmit") = lblInfo.Text
            End If
            sqlConn.Close()
        End Using
    End Sub


    Private Sub bindDataDetail(ByVal pDate As Date, ByVal pPORevNo As String, ByVal pPONo As String, ByVal pAffCode As String, ByVal pSupplierID As String, ByVal pKanban As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "    IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(pPORevNo) & "' AND PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSupplierID) & "')   " & vbCrLf & _
                  "    BEGIN   " & vbCrLf & _
                  "    SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description    " & vbCrLf & _
                  "       ,MOQ = LEFT(MOQ,LEN(MOQ)-3) ,MinOrderQty,SeqNo, QtyBox = LEFT(QtyBox,LEN(QtyBox)-3) ,Maker  ,ISNULL(MonthlyProductionCapacity,0)MonthlyProductionCapacity   " & vbCrLf & _
                  "       ,BYWHAT,POQty ,ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                  "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                  "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                  "       FROM (    " & vbCrLf

            ls_SQL = ls_SQL + "  		SELECT CONVERT(CHAR,row_number() over (order by PORD.PONo)) as NoUrut,PORD.PartNo,PORD.PartNo PartNos,PartName       " & vbCrLf & _
                              "        	,CASE WHEN MPART.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,MU.DESCRIPTION  " & vbCrLf & _
                              "        	,MOQ =CONVERT(CHAR,MOQ),MinOrderQty = MOQ,PORM.SeqNo,QtyBox = CONVERT(CHAR,QtyBox),ISNULL(MPART.Maker,'')Maker       " & vbCrLf & _
                              "          ,(SELECT ISNULL(QtyRemaining, MonthlyProductionCapacity) from MS_SupplierCapacity A  " & vbCrLf & _
                              "             LEFT JOIN RemainingCapacity B ON A.PartNo = B.PartNo AND A.SupplierID = B.SupplierID AND PORD.PartNo = B.PartNo  " & vbCrLf & _
                              "             WHERE B.Period = '" & Format(pDate, "yyyyMM") & "' AND A.SupplierID='" & pSupplierID.Trim & "') MonthlyProductionCapacity   " & vbCrLf & _
                              "         ,'REV. BY AFFILIATE' BYWHAT,PORD.POQty ,hari = CEILING((SUM(PORD.POQty)/QtyBox))  " & vbCrLf & _
                              "       	,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,PORM.Period))),0)    " & vbCrLf & _
                              "    	    ,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,PORM.Period))),0)    " & vbCrLf & _
                              "    		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,PORM.Period))),0)                              " & vbCrLf & _
                              "      	,PORD.DeliveryD1 ,PORD.DeliveryD2 ,PORD.DeliveryD3 ,PORD.DeliveryD4 ,PORD.DeliveryD5 ,PORD.DeliveryD6 ,PORD.DeliveryD7 ,PORD.DeliveryD8 ,PORD.DeliveryD9 ,PORD.DeliveryD10  " & vbCrLf

            ls_SQL = ls_SQL + "      	,PORD.DeliveryD11 ,PORD.DeliveryD12 ,PORD.DeliveryD13 ,PORD.DeliveryD14 ,PORD.DeliveryD15 ,PORD.DeliveryD16 ,PORD.DeliveryD17 ,PORD.DeliveryD18 ,PORD.DeliveryD19 ,PORD.DeliveryD20  " & vbCrLf & _
                              "      	,PORD.DeliveryD21 ,PORD.DeliveryD22 ,PORD.DeliveryD23 ,PORD.DeliveryD24 ,PORD.DeliveryD25 ,PORD.DeliveryD26 ,PORD.DeliveryD27 ,PORD.DeliveryD28 ,PORD.DeliveryD29 ,PORD.DeliveryD30 ,PORD.DeliveryD31  " & vbCrLf & _
                              "      	,row_number() over (order by PORD.PONo) as Sort      " & vbCrLf & _
                              "      	FROM dbo.PORev_Master PORM      " & vbCrLf & _
                              "  		INNER JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "  		LEFT JOIN dbo.MS_Parts MPART ON PORD.PartNo = MPART.PartNo         " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID          " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls                 " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_CurrCls MCUR1 ON PORD.CurrCls = MCUR1.CurrCls          " & vbCrLf & _
                              "         WHERE MONTH(PORM.Period) = MONTH('" & pDate & "') AND YEAR(PORM.Period) = YEAR('" & pDate & "')  " & vbCrLf & _
                              "         AND PORM.PORevNo='" & Trim(pPORevNo) & "' AND PORM.PONo='" & pPONo.Trim & "' AND PORM.SupplierID='" & Trim(pSupplierID) & "'  " & vbCrLf

            ls_SQL = ls_SQL + "            GROUP BY PORD.PONo,PORD.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ " & vbCrLf & _
                              "            ,PORM.SeqNo,QtyBox,PORD.poqty,MPART.Maker,MonthlyProductionCapacity,PORM.Period,MSC.PartNo,PORM.AffiliateID    " & vbCrLf & _
                              "      	   ,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10          " & vbCrLf & _
                              "      	   ,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20        		    " & vbCrLf & _
                              "      	   ,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31     " & vbCrLf & _
                              " 	) Detail1    " & vbCrLf & _
                              "  	UNION ALL    " & vbCrLf

            ls_SQL = ls_SQL + "      SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description    " & vbCrLf & _
                              "       ,MOQ = MOQ ,MinOrderQty,SeqNo,QtyBox = QtyBox ,Maker ,MonthlyProductionCapacity ,BYWHAT    " & vbCrLf & _
                              "       ,POQty, ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                              "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                              "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              "       FROM (    " & vbCrLf & _
                              "      		SELECT '' as NoUrut,'' PartNo,PORD.PartNo PartNos,''PartName,'' KanbanCls,''Description,MOQ = '',MinOrderQty = MOQ ,PORM.SeqNo    " & vbCrLf

            ls_SQL = ls_SQL + "  			,'' QtyBox,ISNULL(MPART.Maker,'')Maker,0 MonthlyProductionCapacity  " & vbCrLf & _
                              "  			,'REV. BY PASI' BYWHAT,PORD.POQty  " & vbCrLf & _
                              "      		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,PORM.Period))),0)    " & vbCrLf & _
                              "    			,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,PORM.Period))),0)    " & vbCrLf & _
                              "    			,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = PORD.PartNo AND MF.AffiliateID = PORM.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,PORM.Period))),0)     " & vbCrLf & _
                              "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6 ,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10 " & vbCrLf & _
                              "      		,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20 " & vbCrLf & _
                              "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31 " & vbCrLf & _
                              "      		,row_number() over (order by PORD.PONo) as Sort      " & vbCrLf & _
                              "      	FROM dbo.PORev_Master PORM      " & vbCrLf & _
                              "  		 LEFT JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID   " & vbCrLf

            ls_SQL = ls_SQL + "  		 LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID   " & vbCrLf & _
                              "  		 LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Parts MPART ON PORD.PartNo = MPART.PartNo         " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID          " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls     " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_CurrCls MCUR1 ON PORD.CurrCls = MCUR1.CurrCls        " & vbCrLf & _
                              "         WHERE MONTH(PORM.Period) = MONTH('" & pDate & "') AND YEAR(PORM.Period) = YEAR('" & pDate & "')  " & vbCrLf & _
                              "         AND PORM.PORevNo='" & pPORevNo.Trim & "' AND PORM.PONo='" & pPONo.Trim & "' AND PORM.SupplierID='" & pSupplierID.Trim & "'  " & vbCrLf
            ls_SQL = ls_SQL + _
                              "            GROUP BY PORD.PONo,PORD.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ " & vbCrLf

            ls_SQL = ls_SQL + "             ,PORM.SeqNo,QtyBox,PORD.POQty,MPART.Maker,MonthlyProductionCapacity ,PORM.Period,MSC.PartNo,PORM.AffiliateID   " & vbCrLf & _
                              "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10        " & vbCrLf & _
                              "      		,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20          " & vbCrLf & _
                              "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31   " & vbCrLf & _
                              "  	)detail2    " & vbCrLf & _
                              "  ORDER BY sort, PartNo DESC    " & vbCrLf & _
                              "  END   " & vbCrLf & _
                              "  ELSE   " & vbCrLf & _
                              "  BEGIN   " & vbCrLf & _
                              "  SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description     " & vbCrLf & _
                              "    ,MOQ = LEFT(MOQ,LEN(MOQ)-3) ,MinOrderQty,SeqNo,QtyBox = LEFT(QtyBox,LEN(QtyBox)-3) ,Maker ,ISNULL(MonthlyProductionCapacity,0)MonthlyProductionCapacity ,BYWHAT     " & vbCrLf & _
                              "    ,POQty ,ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25    " & vbCrLf & _
                              "    ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31    " & vbCrLf & _
                              "     FROM (   " & vbCrLf & _
                              "        SELECT row_number() over (order by AD.PONo) as Sort   ,CONVERT(CHAR,row_number() over (order by AD.PONo)) as NoUrut     " & vbCrLf

            ls_SQL = ls_SQL + "    		,AD.PartNo as PartNo ,AD.PartNo AS PartNos,PartName ,CASE WHEN AD.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls ,MU.Description     " & vbCrLf & _
                              "    		,MOQ =CONVERT(CHAR,MOQ),MinOrderQty = MOQ,AD.SeqNo,QtyBox = CONVERT(CHAR,QtyBox) ,AD.Maker    " & vbCrLf & _
                              "         ,(SELECT ISNULL(QtyRemaining, MonthlyProductionCapacity) from MS_SupplierCapacity A   " & vbCrLf & _
                              "            LEFT JOIN RemainingCapacity B ON A.PartNo = B.PartNo AND A.SupplierID = B.SupplierID AND AD.PartNo = B.PartNo   " & vbCrLf & _
                              "             WHERE B.Period = '" & Format(pDate, "yyyyMM") & "' AND A.SupplierID='" & pSupplierID.Trim & "') MonthlyProductionCapacity    " & vbCrLf & _
                              "         ,'REV. BY AFFILIATE' BYWHAT     " & vbCrLf & _
                              "    		,POQtyOld POqty   " & vbCrLf & _
                              "    		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,PORM.Period))),0)     " & vbCrLf & _
                              "     	,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,PORM.Period))),0)     " & vbCrLf & _
                              "     	,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,PORM.Period))),0)                               " & vbCrLf & _
                              "       	,DeliveryD1Old DeliveryD1,DeliveryD2Old DeliveryD2,DeliveryD3Old DeliveryD3,DeliveryD4Old DeliveryD4,DeliveryD5Old DeliveryD5    " & vbCrLf

            ls_SQL = ls_SQL + "    		,DeliveryD6Old DeliveryD6,DeliveryD7Old DeliveryD7,DeliveryD8Old DeliveryD8,DeliveryD9Old DeliveryD9,DeliveryD10Old DeliveryD10    " & vbCrLf & _
                              "    		,DeliveryD11Old DeliveryD11,DeliveryD12Old DeliveryD12,DeliveryD13Old DeliveryD13,DeliveryD14Old DeliveryD14,DeliveryD15Old DeliveryD15    " & vbCrLf & _
                              "    		,DeliveryD16Old DeliveryD16,DeliveryD17Old DeliveryD17,DeliveryD18 DeliveryD18,DeliveryD19Old DeliveryD19,DeliveryD20Old DeliveryD20    " & vbCrLf & _
                              "    		,DeliveryD21Old DeliveryD21,DeliveryD22Old DeliveryD22,DeliveryD23Old DeliveryD23,DeliveryD24Old DeliveryD24,DeliveryD25Old DeliveryD25    " & vbCrLf & _
                              "    		,DeliveryD26Old DeliveryD26,DeliveryD27Old DeliveryD27,DeliveryD28Old DeliveryD28,DeliveryD29Old DeliveryD29,DeliveryD30Old DeliveryD30,DeliveryD31Old DeliveryD31    " & vbCrLf & _
                              "    		FROM dbo.AffiliateRev_Detail AD " & vbCrLf & _
                              "    		LEFT JOIN PORev_Master PORM ON PORM.PONo = AD.PONo AND PORM.PORevNo = AD.PORevNo AND PORM.AffiliateID = AD.AffiliateID AND PORM.SupplierID = AD.SupplierID " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo    " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID           " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID       " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "    		LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls           " & vbCrLf & _
                              "    		LEFT JOIN dbo.MS_CurrCls MCUR1 ON AD.CurrCls = MCUR1.CurrCls           " & vbCrLf & _
                              "    		WHERE AD.PORevNo='" & pPORevNo.Trim & "' AND AD.PONo='" & pPONo.Trim & "' AND AD.SupplierID='" & pSupplierID.Trim & "'   " & vbCrLf

            ls_SQL = ls_SQL + "    		GROUP BY AD.PONo,AD.PartNo,PartName,AD.KanbanCls,POQtyOld,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity,AD.SeqNo,AD.AffiliateID     " & vbCrLf & _
                              "   		,MSC.PartNo,PORM.Period,DeliveryD1Old,DeliveryD2Old,DeliveryD3Old,DeliveryD4Old,DeliveryD5Old    " & vbCrLf & _
                              "    		,DeliveryD6Old,DeliveryD7Old,DeliveryD8Old,DeliveryD9Old,DeliveryD10Old    " & vbCrLf & _
                              "    		,DeliveryD11Old,DeliveryD12Old,DeliveryD13Old,DeliveryD14Old,DeliveryD15Old    " & vbCrLf & _
                              "    		,DeliveryD16Old,DeliveryD17Old,DeliveryD18,DeliveryD19Old,DeliveryD20Old    " & vbCrLf & _
                              "    		,DeliveryD21Old,DeliveryD22Old,DeliveryD23Old,DeliveryD24Old,DeliveryD25Old    " & vbCrLf & _
                              "    		,DeliveryD26Old,DeliveryD27Old,DeliveryD28Old,DeliveryD29Old,DeliveryD30Old,DeliveryD31Old  )detail1    " & vbCrLf & _
                              "   	UNION ALL    "

            ls_SQL = ls_SQL + "   	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,KanbanCls = KanbanCls ,Description     " & vbCrLf & _
                              "        ,MOQ = MOQ ,MinOrderQty,SeqNo,QtyBox = QtyBox ,Maker ,MonthlyProductionCapacity ,BYWHAT     " & vbCrLf & _
                              "        ,POqty ,ForecastN1 ,ForecastN2 ,ForecastN3    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25    " & vbCrLf & _
                              "        ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                              "        FROM (     " & vbCrLf & _
                              "    		SELECT row_number() over (order by AD.PONo) as Sort,'' as NoUrut,'' PartNo,AD.PartNo PartNos,''PartName,'' KanbanCls,''Description,MOQ = ''      "

            ls_SQL = ls_SQL + "    		,MinOrderQty = MOQ,AD.SeqNo,'' QtyBox,ISNULL(AD.Maker,'')Maker ,0 MonthlyProductionCapacity,'REV. BY PASI' BYWHAT ,POQty   " & vbCrLf & _
                              "    		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,1,PORM.Period))),0)     " & vbCrLf & _
                              "     	,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,2,PORM.Period))),0)     " & vbCrLf & _
                              "     	,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = AD.PartNo AND MF.AffiliateID = AD.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,PORM.Period)) and MONTH(Period) = MONTH(DATEADD(MONTH,3,PORM.Period))),0)                               " & vbCrLf & _
                              "       	,DeliveryD1 ,DeliveryD2 ,DeliveryD3 ,DeliveryD4 ,DeliveryD5     " & vbCrLf & _
                              "   		,DeliveryD6 ,DeliveryD7 ,DeliveryD8 ,DeliveryD9 ,DeliveryD10     " & vbCrLf & _
                              "   		,DeliveryD11 ,DeliveryD12 ,DeliveryD13 ,DeliveryD14     " & vbCrLf & _
                              "   		,DeliveryD15 ,DeliveryD16 ,DeliveryD17 ,DeliveryD18,DeliveryD19 ,DeliveryD20 ,DeliveryD21     " & vbCrLf & _
                              "   		,DeliveryD22 ,DeliveryD23 ,DeliveryD24 ,DeliveryD25 ,DeliveryD26 ,DeliveryD27 ,DeliveryD28 ,DeliveryD29     " & vbCrLf & _
                              "   		,DeliveryD30 ,DeliveryD31     " & vbCrLf & _
                              "   		FROM dbo.AffiliateRev_Detail AD  "

            ls_SQL = ls_SQL + "   		 LEFT JOIN PORev_Master PORM ON PORM.PONo = AD.PONo AND PORM.PORevNo = AD.PORevNo AND PORM.AffiliateID = AD.AffiliateID AND PORM.SupplierID = AD.SupplierID   " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo    " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID    " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID       " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID           " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls    " & vbCrLf & _
                              "   		 LEFT JOIN dbo.MS_CurrCls MCUR1 ON AD.CurrCls = MCUR1.CurrCls   " & vbCrLf & _
                              "          WHERE AD.PORevNo='" & pPORevNo.Trim & "' AND AD.PONo='" & pPONo.Trim & "' AND AD.SupplierID='" & pSupplierID.Trim & "'   " & vbCrLf

            ls_SQL = ls_SQL + "   		 GROUP BY AD.PONo,AD.PartNo,PartName,AD.KanbanCls,POQty,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity,AD.SeqNo,AD.AffiliateID   " & vbCrLf & _
                              "       		,MSC.PartNo,PORM.Period,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5    " & vbCrLf & _
                              "    			,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10    "

            ls_SQL = ls_SQL + "    			,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15    " & vbCrLf & _
                              "    			,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20    " & vbCrLf & _
                              "    			,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25    " & vbCrLf & _
                              "    			,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31   " & vbCrLf & _
                              "   		)detail2   	  " & vbCrLf & _
                              "  	ORDER BY sort, PartNo DESC   " & vbCrLf & _
                              "  END   " & vbCrLf & _
                              "    "




            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Dim pDateDay As DateTime = CDate(Format(pDate, "MM") + "/01/" + Format(pDate, "yyyy"))
                Select Case Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, pDateDay)))
                    Case 28
                        grid.Columns("DeliveryD29").Visible = False
                        grid.Columns("DeliveryD30").Visible = False
                        grid.Columns("DeliveryD31").Visible = False

                    Case 29
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = False
                        grid.Columns("DeliveryD31").Visible = False

                    Case 30
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = True
                        grid.Columns("DeliveryD31").Visible = False

                    Case 31
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = True
                        grid.Columns("DeliveryD31").Visible = True
                End Select
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            End With
            sqlConn.Close()
        End Using
    End Sub

    '    Private Sub SaveDataMaster(ByVal pIsNewData As Boolean, _
    '                         Optional ByVal pDate As String = "", _
    '                         Optional ByVal pPORevNo As String = "", _
    '                         Optional ByVal pPONo As String = "", _
    '                         Optional ByVal pAffCode As String = "", _
    '                         Optional ByVal pSuppCode As String = "", _
    '                         Optional ByVal pComm As String = "", _
    '                         Optional ByVal pKanban As String = "", _
    '                         Optional ByVal pShipBy As String = "", _
    '                         Optional ByVal pSeqNo As String = "")

    '        Dim ls_SQL As String = "", ls_MsgID As String = ""
    '        Dim admin As String = Session("UserID").ToString

    '        Try
    '            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '                sqlConn.Open()
    '                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("PO")
    '                    Dim sqlComm As New SqlCommand()
    '                    ls_SQL = "  IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Master WHERE PORevNo='" & Trim(pPORevNo) & "' AND PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSuppCode) & "')  " & vbCrLf & _
    '                  "  BEGIN  " & vbCrLf & _
    '                  "  INSERT INTO dbo.AffiliateRev_Master " & vbCrLf & _
    '                  "          ( PORevNo ,PONo ,AffiliateID ,SupplierID ,SeqNo ,EntryDate ,EntryUser ,UpdateDate ,UpdateUSer) " & vbCrLf & _
    '                  "  VALUES  ( '" & Trim(pPORevNo) & "' , '" & Trim(pPONo) & "' , '" & Trim(pAffCode) & "' ,'" & Trim(pSuppCode) & "', '" & Trim(pSeqNo) & "' , GETDATE(), '" & Session("UserID") & "' , getdate() ,  '" & Session("UserID") & "')  " & vbCrLf & _
    '                  "          END  " & vbCrLf & _
    '                  "          ELSE  " & vbCrLf & _
    '                  "          BEGIN  " & vbCrLf & _
    '                  "          UPDATE dbo.AffiliateRev_Master  " & vbCrLf & _
    '                  "          SET UpdateDate = GETDATE() " & vbCrLf & _
    '                  "          ,UpdateUSer= '" & Session("UserID") & "' " & vbCrLf

    '                    ls_SQL = ls_SQL + "  WHERE PORevNo='" & Trim(pPORevNo) & "' AND PONo='" & Trim(pPONo) & "' AND AffiliateID='" & Trim(pAffCode) & "' AND SupplierID='" & Trim(pSuppCode) & "' " & vbCrLf & _
    '                                      " END "

    '                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
    '                    sqlComm.ExecuteNonQuery()

    '                    sqlComm.Dispose()
    '                    sqlTran.Commit()
    '                End Using
    '                sqlConn.Close()
    '            End Using
    '        Catch ex As Exception
    '            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '        End Try
    '    End Sub

    '    Private Sub SaveDataDetail(ByVal pIsNewData As Boolean, _
    '                         Optional ByVal pDate As String = "", _
    '                         Optional ByVal pPORevNo As String = "", _
    '                         Optional ByVal pPONo As String = "", _
    '                         Optional ByVal pAffCode As String = "", _
    '                         Optional ByVal pSuppCode As String = "", _
    '                         Optional ByVal pComm As String = "", _
    '                         Optional ByVal pKanban As String = "", _
    '                         Optional ByVal pShipBy As String = "")

    '        Dim ls_SQL As String = "", ls_MsgID As String = ""
    '        Dim ls_DeliveryD1 As Double = 0 : Dim ls_DeliveryD2 As Double = 0 : Dim ls_DeliveryD3 As Double = 0 : Dim ls_DeliveryD4 As Double = 0 : Dim ls_DeliveryD5 As Double = 0
    '        Dim ls_DeliveryD6 As Double = 0 : Dim ls_DeliveryD7 As Double = 0 : Dim ls_DeliveryD8 As Double = 0 : Dim ls_DeliveryD9 As Double = 0 : Dim ls_DeliveryD10 As Double = 0
    '        Dim ls_DeliveryD11 As Double = 0 : Dim ls_DeliveryD12 As Double = 0 : Dim ls_DeliveryD13 As Double = 0 : Dim ls_DeliveryD14 As Double = 0 : Dim ls_DeliveryD15 As Double = 0
    '        Dim ls_DeliveryD16 As Double = 0 : Dim ls_DeliveryD17 As Double = 0 : Dim ls_DeliveryD18 As Double = 0 : Dim ls_DeliveryD19 As Double = 0 : Dim ls_DeliveryD20 As Double = 0
    '        Dim ls_DeliveryD21 As Double = 0 : Dim ls_DeliveryD22 As Double = 0 : Dim ls_DeliveryD23 As Double = 0 : Dim ls_DeliveryD24 As Double = 0 : Dim ls_DeliveryD25 As Double = 0
    '        Dim ls_DeliveryD26 As Double = 0 : Dim ls_DeliveryD27 As Double = 0 : Dim ls_DeliveryD28 As Double = 0 : Dim ls_DeliveryD29 As Double = 0 : Dim ls_DeliveryD30 As Double = 0
    '        Dim ls_DeliveryD31 As Double = 0

    '        Dim ls_POQty As Double = 0
    '        Dim ls_POQtyOld As Double = 0

    '        Dim ls_diffCls As String = ""

    '        Dim ls_SeqNo As String


    '        Dim admin As String = Session("UserID").ToString

    '        Try
    '            Dim iLoop As Long = 0, jLoop As Long = 0
    '            Dim ls_UserID As String = ""

    '            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '                sqlConn.Open()
    '                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("SaveDetail")
    '                    If grid.VisibleRowCount = 0 Then
    '                        ls_MsgID = "6011"
    '                        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
    '                        Session("ZZ010Msg") = lblInfo.Text
    '                        Exit Sub
    '                    End If
    '                    With grid
    '                        For iLoop = 0 To grid.VisibleRowCount - 1
    '                            Dim ls_Kanban As String = .GetRowValues(iLoop, "KanbanCls").ToString()
    '                            'If FlagGrid = 1 Then
    '                            '    GoTo EndNext
    '                            'End If
    '                            ls_SeqNo = .GetRowValues(iLoop, "SeqNo")
    '                            ls_DeliveryD1 = .GetRowValues(iLoop, "DeliveryD1")
    '                            ls_DeliveryD2 = .GetRowValues(iLoop, "DeliveryD2")
    '                            ls_DeliveryD3 = .GetRowValues(iLoop, "DeliveryD3")
    '                            ls_DeliveryD4 = .GetRowValues(iLoop, "DeliveryD4")
    '                            ls_DeliveryD5 = .GetRowValues(iLoop, "DeliveryD5")
    '                            ls_DeliveryD6 = .GetRowValues(iLoop, "DeliveryD6")
    '                            ls_DeliveryD7 = .GetRowValues(iLoop, "DeliveryD7")
    '                            ls_DeliveryD8 = .GetRowValues(iLoop, "DeliveryD8")
    '                            ls_DeliveryD9 = .GetRowValues(iLoop, "DeliveryD9")
    '                            ls_DeliveryD10 = .GetRowValues(iLoop, "DeliveryD10")
    '                            ls_DeliveryD11 = .GetRowValues(iLoop, "DeliveryD11")
    '                            ls_DeliveryD12 = .GetRowValues(iLoop, "DeliveryD12")
    '                            ls_DeliveryD13 = .GetRowValues(iLoop, "DeliveryD13")
    '                            ls_DeliveryD14 = .GetRowValues(iLoop, "DeliveryD14")
    '                            ls_DeliveryD15 = .GetRowValues(iLoop, "DeliveryD15")
    '                            ls_DeliveryD16 = .GetRowValues(iLoop, "DeliveryD16")
    '                            ls_DeliveryD17 = .GetRowValues(iLoop, "DeliveryD17")
    '                            ls_DeliveryD18 = .GetRowValues(iLoop, "DeliveryD18")
    '                            ls_DeliveryD19 = .GetRowValues(iLoop, "DeliveryD19")
    '                            ls_DeliveryD20 = .GetRowValues(iLoop, "DeliveryD20")
    '                            ls_DeliveryD21 = .GetRowValues(iLoop, "DeliveryD21")
    '                            ls_DeliveryD22 = .GetRowValues(iLoop, "DeliveryD22")
    '                            ls_DeliveryD23 = .GetRowValues(iLoop, "DeliveryD23")
    '                            ls_DeliveryD24 = .GetRowValues(iLoop, "DeliveryD24")
    '                            ls_DeliveryD25 = .GetRowValues(iLoop, "DeliveryD25")
    '                            ls_DeliveryD26 = .GetRowValues(iLoop, "DeliveryD26")
    '                            ls_DeliveryD27 = .GetRowValues(iLoop, "DeliveryD27")
    '                            ls_DeliveryD28 = .GetRowValues(iLoop, "DeliveryD28")
    '                            ls_DeliveryD29 = .GetRowValues(iLoop, "DeliveryD29")
    '                            ls_DeliveryD30 = .GetRowValues(iLoop, "DeliveryD30")
    '                            ls_DeliveryD31 = .GetRowValues(iLoop, "DeliveryD31")


    '                            If ls_Kanban = "YES" Then ls_Kanban = "1" Else ls_Kanban = "0"
    '                            Dim byWhat As String = .GetRowValues(iLoop, "BYWHAT")
    '                            If byWhat = "REV. BY AFFILIATE" Then 'OLD
    '                                ls_POQtyOld = .GetRowValues(iLoop, "POQty")
    '                                If ls_POQty = ls_POQtyOld Then
    '                                    ls_diffCls = "0"
    '                                Else
    '                                    ls_diffCls = "1"
    '                                End If
    '                                'Dim ls_AmountAff As Double = .GetRowValues(iLoop, "PriceAff") * .GetRowValues(iLoop, "POQty")
    '                                Dim ls_AmountAff As Double = 0
    '                                ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(txtPORev.Text) & "' AND PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' AND PartNo='" & .GetRowValues(iLoop, "PartNos").ToString & "')  " & vbCrLf & _
    '                                  " BEGIN  " & vbCrLf & _
    '                                  " 	INSERT INTO dbo.AffiliateRev_Detail " & vbCrLf & _
    '                                  "         ( PORevNo, " & vbCrLf & _
    '                                  "           PONo , " & vbCrLf & _
    '                                  "           AffiliateID , " & vbCrLf & _
    '                                  "           SupplierID , " & vbCrLf & _
    '                                  "           PartNo , " & vbCrLf & _
    '                                  "           SeqNo , " & vbCrLf & _
    '                                  "           DifferenceCls , " & vbCrLf & _
    '                                  "           KanbanCls , " & vbCrLf & _
    '                                  "           Maker , " & vbCrLf & _
    '                                  "           POQtyOld , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           CurrCls , " & vbCrLf & _
    '                                                  "           Price , " & vbCrLf & _
    '                                                  "           Amount , " & vbCrLf & _
    '                                                  "           DeliveryD1Old , " & vbCrLf & _
    '                                                  "           DeliveryD2Old , " & vbCrLf & _
    '                                                  "           DeliveryD3Old , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD4Old , " & vbCrLf & _
    '                                                  "           DeliveryD5Old , " & vbCrLf & _
    '                                                  "           DeliveryD6Old , " & vbCrLf & _
    '                                                  "           DeliveryD7Old , " & vbCrLf & _
    '                                                  "           DeliveryD8Old , " & vbCrLf & _
    '                                                  "           DeliveryD9Old , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD10Old , " & vbCrLf & _
    '                                                  "           DeliveryD11Old , " & vbCrLf & _
    '                                                  "           DeliveryD12Old , " & vbCrLf & _
    '                                                  "           DeliveryD13Old , " & vbCrLf & _
    '                                                  "           DeliveryD14Old , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD15Old , " & vbCrLf & _
    '                                                  "           DeliveryD16Old , " & vbCrLf & _
    '                                                  "           DeliveryD17Old , " & vbCrLf & _
    '                                                  "           DeliveryD18Old , " & vbCrLf & _
    '                                                  "           DeliveryD19Old , " & vbCrLf & _
    '                                                  "           DeliveryD20Old , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD21Old , " & vbCrLf & _
    '                                                  "           DeliveryD22Old , " & vbCrLf & _
    '                                                  "           DeliveryD23Old , " & vbCrLf & _
    '                                                  "           DeliveryD24Old , " & vbCrLf & _
    '                                                  "           DeliveryD25Old , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD26Old , " & vbCrLf & _
    '                                                  "           DeliveryD27Old , " & vbCrLf & _
    '                                                  "           DeliveryD28Old , " & vbCrLf & _
    '                                                  "           DeliveryD29Old , " & vbCrLf & _
    '                                                  "           DeliveryD30Old , " & vbCrLf & _
    '                                                  "           DeliveryD31Old , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           EntryDate , " & vbCrLf & _
    '                                                  "           EntryUser , " & vbCrLf & _
    '                                                  "           UpdateDate , " & vbCrLf & _
    '                                                  "           UpdateUser " & vbCrLf & _
    '                                                  "         ) " & vbCrLf & _
    '                                                  " 	VALUES  ( '" & Trim(pPORevNo) & "' , -- PORevNo - char(20) " & vbCrLf & _
    '                                                  "           '" & Trim(pPONo) & "' , -- PONo - char(20) " & vbCrLf & _
    '                                                  "           '" & Trim(pAffCode) & "' , -- AffiliateID - char(20) " & vbCrLf & _
    '                                                  "           '" & Trim(pSuppCode) & "' , -- SupplierID - char(20) " & vbCrLf & _
    '                                                  "           '" & .GetRowValues(iLoop, "PartNos").ToString & "' , -- PartNo - char(25) " & vbCrLf & _
    '                                                  "           '" & ls_diffCls & "' , -- PartNo - char(25) " & vbCrLf & _
    '                                                  "           '" & ls_SeqNo & "' , -- PartNo - char(25) " & vbCrLf & _
    '                                                  "           '" & ls_Kanban & "' , -- KanbanCls - char(1) " & vbCrLf


    '                                ls_SQL = ls_SQL + "           '" & .GetRowValues(iLoop, "Maker") & "', -- Maker - char(20) " & vbCrLf & _
    '                                                  "           " & ls_POQtyOld & " , -- POQtyOld - numeric " & vbCrLf & _
    '                                                  "           '' , -- CurrCls - char(2) " & vbCrLf & _
    '                                                  "           '0' , -- Price - numeric " & vbCrLf & _
    '                                                  "           " & ls_AmountAff & " , -- Amount - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD1 & " , -- DeliveryD1Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD2 & " , -- DeliveryD2Old - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD3 & " , -- DeliveryD3Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD4 & " , -- DeliveryD4Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD5 & " , -- DeliveryD5Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD6 & " , -- DeliveryD6Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD7 & " , -- DeliveryD7Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD8 & " , -- DeliveryD8Old - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD9 & " , -- DeliveryD9Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD10 & " , -- DeliveryD10Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD11 & " , -- DeliveryD11Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD12 & " , -- DeliveryD12Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD13 & " , -- DeliveryD13Old - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD14 & " , -- DeliveryD14Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD15 & " , -- DeliveryD15Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD16 & " , -- DeliveryD16Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD17 & " , -- DeliveryD17Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD18 & " , -- DeliveryD18Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD19 & " , -- DeliveryD19Old - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD20 & " , -- DeliveryD20Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD21 & " , -- DeliveryD21Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD22 & " , -- DeliveryD22Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD23 & " , -- DeliveryD23Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD24 & " , -- DeliveryD24Old - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD25 & " , -- DeliveryD25Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD26 & " , -- DeliveryD26Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD27 & " , -- DeliveryD27Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD28 & " , -- DeliveryD28Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD29 & " , -- DeliveryD29Old - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD30 & " , -- DeliveryD30Old - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD31 & " , -- DeliveryD31Old - numeric " & vbCrLf & _
    '                                                  "           getdate() , -- EntryDate - datetime " & vbCrLf & _
    '                                                  "           '" & Session("UserID") & "' , -- EntryUser - char(15) " & vbCrLf & _
    '                                                  "           getdate() , -- UpdateDate - datetime " & vbCrLf & _
    '                                                  "           '" & Session("UserID") & "'  -- UpdateUser - char(15) " & vbCrLf & _
    '                                                  "         ) " & vbCrLf & _
    '                                                  "         END	 " & vbCrLf & _
    '                                                  "         ELSE	 " & vbCrLf & _
    '                                                  "         BEGIN  " & vbCrLf & _
    '                                                  "            UPDATE [dbo].[AffiliateRev_Detail] " & vbCrLf

    '                                ls_SQL = ls_SQL + " 		   SET [KanbanCls] = '" & ls_Kanban & "' " & vbCrLf & _
    '                                                  "               ,[Maker] = '" & .GetRowValues(iLoop, "Maker") & "' " & vbCrLf & _
    '                                                  " 			  ,[POQtyOld] = " & ls_POQtyOld & " " & vbCrLf & _
    '                                                  " 			  ,[CurrCls] = '' " & vbCrLf & _
    '                                                  " 			  ,[Price] = 0 " & vbCrLf & _
    '                                                  " 			  ,[Amount] = " & ls_AmountAff & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD1Old] = " & ls_DeliveryD1 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD2Old] = " & ls_DeliveryD2 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD3Old] = " & ls_DeliveryD3 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD4Old] = " & ls_DeliveryD4 & "" & vbCrLf & _
    '                                                  " 			  ,[DeliveryD5Old] =" & ls_DeliveryD5 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD6Old] = " & ls_DeliveryD6 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD7Old] =" & ls_DeliveryD7 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD8Old] = " & ls_DeliveryD8 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD9Old] = " & ls_DeliveryD9 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD10Old] = " & ls_DeliveryD10 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD11Old] = " & ls_DeliveryD11 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD12Old] = " & ls_DeliveryD12 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD13Old] = " & ls_DeliveryD13 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD14Old] = " & ls_DeliveryD14 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD15Old] = " & ls_DeliveryD15 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD16Old] = " & ls_DeliveryD16 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD17Old] = " & ls_DeliveryD17 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD18Old] = " & ls_DeliveryD18 & " "

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD19Old] = " & ls_DeliveryD19 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD20Old] = " & ls_DeliveryD20 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD21Old] = " & ls_DeliveryD21 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD22Old] = " & ls_DeliveryD22 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD23Old] = " & ls_DeliveryD23 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD24Old] = " & ls_DeliveryD24 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD25Old] = " & ls_DeliveryD25 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD26Old] = " & ls_DeliveryD26 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD27Old] = " & ls_DeliveryD27 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD28Old] = " & ls_DeliveryD28 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD29Old] = " & ls_DeliveryD29 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD30Old] = " & ls_DeliveryD30 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD31Old] = " & ls_DeliveryD31 & "  " & vbCrLf & _
    '                                                  " 			  ,[UpdateDate] = getdate() " & vbCrLf & _
    '                                                  " 			  ,[UpdateUser] = '" & Session("UserID") & "' " & vbCrLf & _
    '                                                  " 			WHERE PORevNo='" & Trim(txtPORev.Text) & "' " & vbCrLf & _
    '                                                  "               AND [PONo] = '" & Trim(txtPONo.Text) & "' " & vbCrLf & _
    '                                                  " 			  AND [AffiliateID] ='" & Trim(txtAffiliateID.Text) & "' " & vbCrLf & _
    '                                                  " 			  AND [SupplierID] = '" & Trim(txtSupplierCode.Text) & "'" & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  AND [PartNo] = '" & .GetRowValues(iLoop, "PartNos") & "' " & vbCrLf & _
    '                                                  " 		 END  "


    '                                ls_MsgID = "1002"

    '                                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
    '                                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
    '                                sqlComm.ExecuteNonQuery()
    '                                sqlComm.Dispose()
    '                            Else
    '                                'BY PASI New
    '                                ls_POQty = .GetRowValues(iLoop, "POQty")
    '                                If ls_POQty = ls_POQtyOld Then
    '                                    ls_diffCls = "1"
    '                                Else
    '                                    ls_diffCls = "0"
    '                                End If
    '                                ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(txtPORev.Text) & "' AND PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' AND PartNo='" & .GetRowValues(iLoop, "PartNos").ToString & "')  " & vbCrLf & _
    '                                  " BEGIN  " & vbCrLf & _
    '                                  " 	INSERT INTO dbo.AffiliateRev_Detail " & vbCrLf & _
    '                                  "         ( PORevNo, " & vbCrLf & _
    '                                  "           PONo , " & vbCrLf & _
    '                                  "           AffiliateID , " & vbCrLf & _
    '                                  "           SupplierID , " & vbCrLf & _
    '                                  "           PartNo , " & vbCrLf & _
    '                                  "           --KanbanCls , " & vbCrLf & _
    '                                  "           Maker , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           POQty , " & vbCrLf & _
    '                                                  "           CurrCls , " & vbCrLf & _
    '                                                  "           Price , " & vbCrLf & _
    '                                                  "           Amount , " & vbCrLf & _
    '                                                  "           DeliveryD1 , " & vbCrLf & _
    '                                                  "           DeliveryD2 , " & vbCrLf & _
    '                                                  "           DeliveryD3 , " & vbCrLf & _
    '                                                  "           DeliveryD4 , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD5 , " & vbCrLf & _
    '                                                  "           DeliveryD6 , " & vbCrLf & _
    '                                                  "           DeliveryD7 , " & vbCrLf & _
    '                                                  "           DeliveryD8 , " & vbCrLf & _
    '                                                  "           DeliveryD9 , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD10 , " & vbCrLf & _
    '                                                  "           DeliveryD11 , " & vbCrLf & _
    '                                                  "           DeliveryD12 , " & vbCrLf & _
    '                                                  "           DeliveryD13 , " & vbCrLf & _
    '                                                  "           DeliveryD14 , " & vbCrLf & _
    '                                                  "           DeliveryD15 , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD16 , " & vbCrLf & _
    '                                                  "           DeliveryD17 , " & vbCrLf & _
    '                                                  "           DeliveryD18 , " & vbCrLf & _
    '                                                  "           DeliveryD19 , " & vbCrLf & _
    '                                                  "           DeliveryD20 , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD21 , " & vbCrLf & _
    '                                                  "           DeliveryD22 , " & vbCrLf & _
    '                                                  "           DeliveryD23 , " & vbCrLf & _
    '                                                  "           DeliveryD24 , " & vbCrLf & _
    '                                                  "           DeliveryD25 , " & vbCrLf & _
    '                                                  "           DeliveryD26 , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           DeliveryD27 , " & vbCrLf & _
    '                                                  "           DeliveryD28 , " & vbCrLf & _
    '                                                  "           DeliveryD29 , " & vbCrLf & _
    '                                                  "           DeliveryD30 , " & vbCrLf & _
    '                                                  "           DeliveryD31 , " & vbCrLf

    '                                ls_SQL = ls_SQL + "           EntryDate , " & vbCrLf & _
    '                                                  "           EntryUser , " & vbCrLf & _
    '                                                  "           UpdateDate , " & vbCrLf & _
    '                                                  "           UpdateUser " & vbCrLf & _
    '                                                  "         ) " & vbCrLf & _
    '                                                  " 	VALUES  ( '" & Trim(pPORevNo) & "' , -- PORevNo - char(20) " & vbCrLf & _
    '                                                  "           '" & Trim(pPONo) & "' , -- PONo - char(20) " & vbCrLf & _
    '                                                  "           '" & Trim(pAffCode) & "' , -- AffiliateID - char(20) " & vbCrLf & _
    '                                                  "           '" & Trim(pSuppCode) & "' , -- SupplierID - char(20) " & vbCrLf & _
    '                                                  "           '" & .GetRowValues(iLoop, "PartNos").ToString & "' , -- PartNo - char(25) " & vbCrLf & _
    '                                                  "           --'" & Trim(ls_Kanban) & "' , -- KanbanCls - char(1) " & vbCrLf

    '                                ls_SQL = ls_SQL + "           '" & .GetRowValues(iLoop, "Maker") & "', -- Maker - char(20) " & vbCrLf & _
    '                                                  "           " & ls_POQty & " , -- POQtyOld - numeric " & vbCrLf & _
    '                                                  "           '' , -- CurrCls - char(2) " & vbCrLf & _
    '                                                  "           0 , -- Price - numeric " & vbCrLf & _
    '                                                  "           0 , -- Amount - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD1 & " , -- DeliveryD1 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD2 & " , -- DeliveryD2 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD3 & " , -- DeliveryD3 - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD4 & " , -- DeliveryD4 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD5 & " , -- DeliveryD5 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD6 & " , -- DeliveryD6 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD7 & " , -- DeliveryD7 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD8 & " , -- DeliveryD8 - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD9 & " , -- DeliveryD9 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD10 & " , -- DeliveryD10 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD11 & " , -- DeliveryD11 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD12 & " , -- DeliveryD12 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD13 & " , -- DeliveryD13 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD14 & " , -- DeliveryD14 - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD15 & " , -- DeliveryD15 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD16 & " , -- DeliveryD16 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD17 & " , -- DeliveryD17 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD18 & " , -- DeliveryD18 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD19 & " , -- DeliveryD19 - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD20 & " , -- DeliveryD20 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD21 & " , -- DeliveryD21 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD22 & " , -- DeliveryD22 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD23 & " , -- DeliveryD23 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD24 & " , -- DeliveryD24 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD25 & " , -- DeliveryD25 - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD26 & " , -- DeliveryD26 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD27 & " , -- DeliveryD27 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD28 & " , -- DeliveryD28 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD29 & " , -- DeliveryD29 - numeric " & vbCrLf & _
    '                                                  "           " & ls_DeliveryD30 & " , -- DeliveryD30 - numeric " & vbCrLf

    '                                ls_SQL = ls_SQL + "           " & ls_DeliveryD31 & " , -- DeliveryD31 - numeric " & vbCrLf & _
    '                                                  "           getdate() , -- EntryDate - datetime " & vbCrLf & _
    '                                                  "           '" & Session("UserID") & "' , -- EntryUser - char(15) " & vbCrLf & _
    '                                                  "           getdate() , -- UpdateDate - datetime " & vbCrLf & _
    '                                                  "           '" & Session("UserID") & "'  -- UpdateUser - char(15) " & vbCrLf & _
    '                                                  "         ) " & vbCrLf & _
    '                                                  "         END	 " & vbCrLf & _
    '                                                  "         ELSE	 " & vbCrLf & _
    '                                                  "         BEGIN  " & vbCrLf & _
    '                                                  "            UPDATE [dbo].[AffiliateRev_Detail] " & vbCrLf

    '                                ls_SQL = ls_SQL + " 		   SET --[KanbanCls] = '" & ls_Kanban & "' " & vbCrLf & _
    '                                                  " 			  [Maker] = '" & .GetRowValues(iLoop, "Maker") & "' " & vbCrLf & _
    '                                                  " 			  ,[POQty] = " & ls_POQty & " " & vbCrLf & _
    '                                                  " 			  ,[CurrCls] = '' " & vbCrLf & _
    '                                                  " 			  ,[Price] = 0 " & vbCrLf & _
    '                                                  " 			  ,[Amount] = 0 " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD1] = " & ls_DeliveryD1 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD2] = " & ls_DeliveryD2 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD3] = " & ls_DeliveryD3 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD4] = " & ls_DeliveryD4 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD5] = " & ls_DeliveryD5 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD6] =  " & ls_DeliveryD6 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD7] =  " & ls_DeliveryD7 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD8] =  " & ls_DeliveryD8 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD9] =  " & ls_DeliveryD9 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD10] = " & ls_DeliveryD10 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD11] = " & ls_DeliveryD11 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD12] = " & ls_DeliveryD12 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD13] = " & ls_DeliveryD13 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD14] = " & ls_DeliveryD14 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD15] = " & ls_DeliveryD15 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD16] = " & ls_DeliveryD16 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD17] = " & ls_DeliveryD17 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD18] = " & ls_DeliveryD18 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD19] = " & ls_DeliveryD19 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD20] = " & ls_DeliveryD20 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD21] = " & ls_DeliveryD21 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD22] = " & ls_DeliveryD22 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD23] = " & ls_DeliveryD23 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD24] = " & ls_DeliveryD24 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD25] = " & ls_DeliveryD25 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD26] = " & ls_DeliveryD26 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD27] = " & ls_DeliveryD27 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD28] = " & ls_DeliveryD28 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD29] = " & ls_DeliveryD29 & " " & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  ,[DeliveryD30] = " & ls_DeliveryD30 & " " & vbCrLf & _
    '                                                  " 			  ,[DeliveryD31] = " & ls_DeliveryD31 & " " & vbCrLf & _
    '                                                  " 			  ,[UpdateDate] = getdate() " & vbCrLf & _
    '                                                  " 			  ,[UpdateUser] = '" & Session("UserID") & "' " & vbCrLf & _
    '                                                  " 			WHERE PORevNo='" & Trim(txtPORev.Text) & "' " & vbCrLf & _
    '                                                  "               AND [PONo] = '" & Trim(txtPONo.Text) & "' " & vbCrLf & _
    '                                                  " 			  AND [AffiliateID] ='" & Trim(txtAffiliateID.Text) & "' " & vbCrLf & _
    '                                                  " 			  AND [SupplierID] = '" & Trim(txtSupplierCode.Text) & "'" & vbCrLf

    '                                ls_SQL = ls_SQL + " 			  AND [PartNo] = '" & .GetRowValues(iLoop, "PartNos") & "' " & vbCrLf & _
    '                                                  " 		 END  "


    '                                ls_MsgID = "1002"

    '                                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
    '                                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
    '                                sqlComm.ExecuteNonQuery()
    '                                sqlComm.Dispose()
    '                            End If



    'EndNext:
    '                        Next iLoop


    '                        sqlTran.Commit()
    '                        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
    '                        If lblInfo.Text = "[] " Then lblInfo.Text = ""
    '                        Session("ZZ010Msg") = lblInfo.Text
    '                    End With
    '                End Using

    '                sqlConn.Close()


    '            End Using
    '        Catch ex As Exception
    '            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '        End Try
    '    End Sub

    '    Private Sub UpdatePO(ByVal pIsNewData As Boolean, _
    '                         Optional ByVal pAffCode As String = "", _
    '                         Optional ByVal pPORevNo As String = "", _
    '                         Optional ByVal pPONo As String = "", _
    '                         Optional ByVal pSuppCode As String = "")

    '        Dim ls_SQL As String = "", ls_MsgID As String = ""
    '        Dim admin As String = Session("UserID").ToString

    '        Try
    '            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '                sqlConn.Open()

    '                ls_SQL = " UPDATE dbo.PORev_Master " & vbCrLf & _
    '                          " SET PASISendAffiliateUser='" & admin & "' " & vbCrLf & _
    '                          " ,PASISendAffiliateDate=getdate() " & vbCrLf & _
    '                          " WHERE PORevNo='" & pPORevNo & "' " & vbCrLf & _
    '                          " AND PONo='" & pPONo & "'  " & vbCrLf & _
    '                          " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
    '                          " AND SupplierID='" & pSuppCode & "' "

    '                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
    '                sqlComm.ExecuteNonQuery()
    '                sqlComm.Dispose()
    '                sqlConn.Close()
    '            End Using
    '        Catch ex As Exception
    '            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '        End Try
    '    End Sub

    '    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
    '        Try
    '            'Dim ls_SQL As String = ""
    '            'Dim ls_MsgID As String = ""

    'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '    sqlConn.Open()

    '    ls_SQL = "SELECT AffiliateID" & vbCrLf & _
    '                " FROM MS_Affiliate " & _
    '                " WHERE AffiliateID= '" & Trim(pAffiliate) & "'"

    '    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '    Dim ds As New DataSet
    '    sqlDA.Fill(ds)

    '    If ds.Tables(0).Rows.Count > 0 And grid.FocusedRowIndex = -1 Then
    '        ls_MsgID = "6018"
    '        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
    '        AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
    '        flag = False
    '        Return False
    '    ElseIf ds.Tables(0).Rows.Count > 0 Then
    '        lblInfo.Text = "Affiliate ID with ID " & txtAffiliateID.Text & " already exists in the database."
    '        Return False
    '    End If
    '    Return True
    '    sqlConn.Close()
    'End Using
    '    Catch ex As Exception
    '        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '    End Try

    'End Function

    'Private Function BindDataExcel() As DataSet
    '    Dim ls_SQL As String = ""
    '    Dim tanggal As Date = FormatDateTime(Trim(dtPeriodFrom.Text), DateFormat.ShortDate)

    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.AffiliateRev_Detail WHERE PORevNo='" & Trim(txtPORev.Text) & "' AND PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "')   " & vbCrLf & _
    '              " BEGIN  " & vbCrLf & _
    '              " 	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
    '              " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
    '              "       ,MOQ = LEFT(MOQ,LEN(MOQ)-3) , QtyBox = LEFT(QtyBox,LEN(QtyBox)-3) ,Maker   " & vbCrLf & _
    '              "       ,POQty     " & vbCrLf & _
    '              "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
    '              "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
    '              "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
    '              "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
    '              "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf

    '        ls_SQL = ls_SQL + "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
    '                          "       FROM (    " & vbCrLf & _
    '                          " 			SELECT CONVERT(CHAR,row_number() over (order by PMU.PONo)) as NoUrut,PDU.PartNo,PDU.PartNo PartNos,PartName ,PMU.PONo     " & vbCrLf & _
    '                          "        		,CASE WHEN MPART.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,MU.DESCRIPTION  " & vbCrLf & _
    '                          "        		,MOQ =CONVERT(CHAR,MOQ),QtyBox = CONVERT(CHAR,QtyBox),ISNULL(MPART.Maker,'')Maker       " & vbCrLf & _
    '                          " 			,PDU.POQty  " & vbCrLf & _
    '                          " 			,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    			,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    			,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)                              " & vbCrLf & _
    '                          "      		,'BEFORE' BYWHAT " & vbCrLf & _
    '                          "      		,PDU.DeliveryD1 ,PDU.DeliveryD2 ,PDU.DeliveryD3 ,PDU.DeliveryD4 ,PDU.DeliveryD5 ,PDU.DeliveryD6 ,PDU.DeliveryD7 ,PDU.DeliveryD8 ,PDU.DeliveryD9 ,PDU.DeliveryD10  " & vbCrLf

    '        ls_SQL = ls_SQL + "      		,PDU.DeliveryD11 ,PDU.DeliveryD12 ,PDU.DeliveryD13 ,PDU.DeliveryD14 ,PDU.DeliveryD15 ,PDU.DeliveryD16 ,PDU.DeliveryD17 ,PDU.DeliveryD18 ,PDU.DeliveryD19 ,PDU.DeliveryD20  " & vbCrLf & _
    '                          "      	,PDU.DeliveryD21 ,PDU.DeliveryD22 ,PDU.DeliveryD23 ,PDU.DeliveryD24 ,PDU.DeliveryD25 ,PDU.DeliveryD26 ,PDU.DeliveryD27 ,PDU.DeliveryD28 ,PDU.DeliveryD29 ,PDU.DeliveryD30 ,PDU.DeliveryD31  " & vbCrLf & _
    '                          "      	,row_number() over (order by PDU.PONo) as Sort      " & vbCrLf & _
    '                          "      	FROM dbo.PO_MasterUpload PMU " & vbCrLf & _
    '                          "  		INNER JOIN dbo.PO_DetailUpload PDU ON PMU.PONo = PDU.PONo  AND PMU.AffiliateID = PDU.AffiliateID AND PMU.SupplierID = PDU.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN PO_Master POM ON PDU.AffiliateID = POM.AffiliateID AND PDU.PONo = POM.PONo AND PDU.SupplierID = POM.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = PDU.PartNo and MP.AffiliateID = PDU.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_Parts MPART ON PDU.PartNo = MPART.PartNo         " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_Supplier MS ON PDU.SupplierID = MS.SupplierID          " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_Affiliate MA ON PDU.AffiliateID = MA.AffiliateID      " & vbCrLf

    '        ls_SQL = ls_SQL + "          LEFT JOIN dbo.MS_SupplierCapacity MSC ON PDU.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PDU.SupplierID=MSC.SupplierID          " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls                LEFT JOIN dbo.MS_CurrCls MCUR1 ON PDU.CurrCls = MCUR1.CurrCls          " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf & _
    '                          "         WHERE  PMU.SupplierID='" & Trim(cboSupplierCode.Text) & "'  " & vbCrLf & _
    '                          "            GROUP BY PMU.PONo,PDU.PONo,PDU.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ,QtyBox,PDU.poqty,MPART.Maker,MonthlyProductionCapacity       " & vbCrLf & _
    '                          "      		,PDU.CurrCls,MCUR1.Description,PDU.Price,PDU.Amount,MP.CurrCls,MCUR2.Description,MP.Price   " & vbCrLf & _
    '                          "       		,MSC.PartNo    " & vbCrLf & _
    '                          "      		,PDU.DeliveryD1,PDU.DeliveryD2,PDU.DeliveryD3,PDU.DeliveryD4,PDU.DeliveryD5,PDU.DeliveryD6,PDU.DeliveryD7,PDU.DeliveryD8,PDU.DeliveryD9,PDU.DeliveryD10          " & vbCrLf & _
    '                          "      		,PDU.DeliveryD11,PDU.DeliveryD12,PDU.DeliveryD13,PDU.DeliveryD14,PDU.DeliveryD15,PDU.DeliveryD16,PDU.DeliveryD17,PDU.DeliveryD18,PDU.DeliveryD19,PDU.DeliveryD20        		    " & vbCrLf & _
    '                          "      		,PDU.DeliveryD21,PDU.DeliveryD22,PDU.DeliveryD23,PDU.DeliveryD24,PDU.DeliveryD25,PDU.DeliveryD26,PDU.DeliveryD27,PDU.DeliveryD28,PDU.DeliveryD29,PDU.DeliveryD30,PDU.DeliveryD31     " & vbCrLf & _
    '                          " 	)detail1 " & vbCrLf

    '        ls_SQL = ls_SQL + " 	UNION ALL  " & vbCrLf & _
    '                          " 	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
    '                          " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
    '                          "       ,MOQ = MOQ , QtyBox = QtyBox ,Maker   " & vbCrLf & _
    '                          "       ,POQty     " & vbCrLf & _
    '                          "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf

    '        ls_SQL = ls_SQL + "       FROM (   " & vbCrLf & _
    '                          "  		SELECT '' NoUrut,'' PartNo,PORD.PartNo PartNos,''PartName ,'' PONo,'' KanbanCls,'' DESCRIPTION  " & vbCrLf & _
    '                          "  		,''MOQ,''QtyBox,ISNULL(MPART.Maker,'')Maker       " & vbCrLf & _
    '                          "         ,PORD.POQty ,'AFTER' BYWHAT " & vbCrLf & _
    '                          "         ,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    	    ,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)                              " & vbCrLf & _
    '                          "      	,PORD.DeliveryD1 ,PORD.DeliveryD2 ,PORD.DeliveryD3 ,PORD.DeliveryD4 ,PORD.DeliveryD5 ,PORD.DeliveryD6 ,PORD.DeliveryD7 ,PORD.DeliveryD8 ,PORD.DeliveryD9 ,PORD.DeliveryD10  " & vbCrLf & _
    '                          "      	,PORD.DeliveryD11 ,PORD.DeliveryD12 ,PORD.DeliveryD13 ,PORD.DeliveryD14 ,PORD.DeliveryD15 ,PORD.DeliveryD16 ,PORD.DeliveryD17 ,PORD.DeliveryD18 ,PORD.DeliveryD19 ,PORD.DeliveryD20  " & vbCrLf & _
    '                          "      	,PORD.DeliveryD21 ,PORD.DeliveryD22 ,PORD.DeliveryD23 ,PORD.DeliveryD24 ,PORD.DeliveryD25 ,PORD.DeliveryD26 ,PORD.DeliveryD27 ,PORD.DeliveryD28 ,PORD.DeliveryD29 ,PORD.DeliveryD30 ,PORD.DeliveryD31  " & vbCrLf & _
    '                          "      	,row_number() over (order by PORD.PONo) as Sort      " & vbCrLf

    '        ls_SQL = ls_SQL + "      	FROM dbo.PORev_Master PORM      " & vbCrLf & _
    '                          "  		INNER JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = PORD.PartNo and MP.AffiliateID = PORD.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_Parts MPART ON PORD.PartNo = MPART.PartNo         " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID          " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID          " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls    " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_CurrCls MCUR1 ON PORD.CurrCls = MCUR1.CurrCls  " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf

    '        ls_SQL = ls_SQL + "         WHERE MONTH(PORM.Period) = MONTH('" & tanggal & "') AND YEAR(PORM.Period) = YEAR('" & tanggal & "')  " & vbCrLf & _
    '                          "         AND PORM.PORevNo='" & Trim(txtRevisionNo.Text) & "' AND PORM.SupplierID='" & Trim(cboSupplierCode.Text) & "'  " & vbCrLf & _
    '                          "         GROUP BY PORD.PONo,PORD.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ,PORM.SeqNo,QtyBox,PORD.poqty,MPART.Maker,MonthlyProductionCapacity       " & vbCrLf & _
    '                          "      		,PORD.CurrCls,MCUR1.Description,PORD.Price,PORD.Amount,MP.CurrCls,MCUR2.Description,MP.Price   " & vbCrLf & _
    '                          "       		,PORM.Period,MSC.PartNo    " & vbCrLf & _
    '                          "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10          " & vbCrLf & _
    '                          "      		,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20        		    " & vbCrLf & _
    '                          "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31     " & vbCrLf & _
    '                          " 	) Detail2 " & vbCrLf & _
    '                          " 	UNION ALL    " & vbCrLf & _
    '                          "     SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf

    '        ls_SQL = ls_SQL + " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
    '                          "       ,MOQ = MOQ , QtyBox = QtyBox ,Maker   " & vbCrLf & _
    '                          "       ,POQty     " & vbCrLf & _
    '                          "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
    '                          "       FROM (    " & vbCrLf & _
    '                          "      		SELECT '' as NoUrut,'' PartNo,PORD.PartNo PartNos,''PartName,''PONo,'' KanbanCls,''Description,MOQ = '',MinOrderQty = MOQ ,PORM.SeqNo    " & vbCrLf

    '        ls_SQL = ls_SQL + "  			,'' QtyBox,ISNULL(MPART.Maker,'')Maker,'' MonthlyProductionCapacity  " & vbCrLf & _
    '                          "  			,'SUPPLIER APPROVAL' BYWHAT  " & vbCrLf & _
    '                          "      		,PORD.POQty  " & vbCrLf & _
    '                          "      		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    			,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    			,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)     " & vbCrLf & _
    '                          "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10 " & vbCrLf & _
    '                          " 			,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20 " & vbCrLf & _
    '                          "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31 " & vbCrLf & _
    '                          "      		,row_number() over (order by PORD.PONo) as Sort      " & vbCrLf & _
    '                          "      	FROM dbo.PORev_Master PORM      " & vbCrLf

    '        ls_SQL = ls_SQL + "  		INNER JOIN dbo.PORev_Detail PORD ON PORM.PONo = PORD.PONo AND PORM.PORevNo = PORD.PORevNo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = PORD.PartNo  and MP.AffiliateID = PORD.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_Parts MPART ON PORD.PartNo = MPART.PartNo         " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID          " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID  " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls     " & vbCrLf & _
    '                          "          LEFT JOIN dbo.MS_CurrCls MCUR1 ON PORD.CurrCls = MCUR1.CurrCls      " & vbCrLf & _
    '                          "    		LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls     " & vbCrLf

    '        ls_SQL = ls_SQL + "         WHERE MONTH(PORM.Period) = MONTH('" & tanggal & "') AND YEAR(PORM.Period) = YEAR('" & tanggal & "')  " & vbCrLf & _
    '                          "         AND PORM.PORevNo='" & Trim(txtRevisionNo.Text) & "' AND PORM.SupplierID='" & Trim(cboSupplierCode.Text) & "'  " & vbCrLf & _
    '                          "            GROUP BY PORD.PONo,PORD.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ,PORM.SeqNo,QtyBox,PORD.POQty,MPART.Maker,MonthlyProductionCapacity  " & vbCrLf & _
    '                          "      		,PORD.CurrCls,MCUR1.Description,PORD.Price,PORD.Amount,MP.CurrCls,MCUR2.Description,MP.Price   " & vbCrLf & _
    '                          "              ,PORM.Period,MSC.PartNo   " & vbCrLf & _
    '                          "      		,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10        " & vbCrLf & _
    '                          "      		,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20          " & vbCrLf & _
    '                          "      		,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31   " & vbCrLf & _
    '                          " 		)detail3 " & vbCrLf & _
    '                          "      	ORDER BY sort, PartNo DESC  " & vbCrLf & _
    '                          " END   " & vbCrLf

    '        ls_SQL = ls_SQL + " ELSE   " & vbCrLf & _
    '                          " BEGIN   " & vbCrLf & _
    '                          " 	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
    '                          " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
    '                          "       ,MOQ = LEFT(MOQ,LEN(MOQ)-3) , QtyBox = LEFT(QtyBox,LEN(QtyBox)-3) ,Maker   " & vbCrLf & _
    '                          "       ,POQty     " & vbCrLf & _
    '                          "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf

    '        ls_SQL = ls_SQL + "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
    '                          "       FROM (    " & vbCrLf & _
    '                          " 			SELECT CONVERT(CHAR,row_number() over (order by PMU.PONo)) as NoUrut,PDU.PartNo,PDU.PartNo PartNos,PartName ,PMU.PONo     " & vbCrLf & _
    '                          "        		,CASE WHEN MPART.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,MU.DESCRIPTION  " & vbCrLf & _
    '                          "        		,MOQ =CONVERT(CHAR,MOQ),QtyBox = CONVERT(CHAR,QtyBox),ISNULL(MPART.Maker,'')Maker       " & vbCrLf & _
    '                          " 			,PDU.POQty  " & vbCrLf & _
    '                          " 			,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    			,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    			,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)                              " & vbCrLf & _
    '                          "      		,'BEFORE' BYWHAT " & vbCrLf & _
    '                          "      		,PDU.DeliveryD1 ,PDU.DeliveryD2 ,PDU.DeliveryD3 ,PDU.DeliveryD4 ,PDU.DeliveryD5 ,PDU.DeliveryD6 ,PDU.DeliveryD7 ,PDU.DeliveryD8 ,PDU.DeliveryD9 ,PDU.DeliveryD10  " & vbCrLf

    '        ls_SQL = ls_SQL + "      		,PDU.DeliveryD11 ,PDU.DeliveryD12 ,PDU.DeliveryD13 ,PDU.DeliveryD14 ,PDU.DeliveryD15 ,PDU.DeliveryD16 ,PDU.DeliveryD17 ,PDU.DeliveryD18 ,PDU.DeliveryD19 ,PDU.DeliveryD20  " & vbCrLf & _
    '                          "      	,PDU.DeliveryD21 ,PDU.DeliveryD22 ,PDU.DeliveryD23 ,PDU.DeliveryD24 ,PDU.DeliveryD25 ,PDU.DeliveryD26 ,PDU.DeliveryD27 ,PDU.DeliveryD28 ,PDU.DeliveryD29 ,PDU.DeliveryD30 ,PDU.DeliveryD31  " & vbCrLf & _
    '                          "      	,row_number() over (order by PDU.PONo) as Sort      " & vbCrLf & _
    '                          "      	FROM dbo.PO_MasterUpload PMU " & vbCrLf & _
    '                          "  		INNER JOIN dbo.PO_DetailUpload PDU ON PMU.PONo = PDU.PONo  AND PMU.AffiliateID = PDU.AffiliateID AND PMU.SupplierID = PDU.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN PO_Master POM ON PDU.AffiliateID = POM.AffiliateID AND PDU.PONo = POM.PONo AND PDU.SupplierID = POM.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
    '                          "  		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = PDU.PartNo and MP.AffiliateID = PDU.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_Parts MPART ON PDU.PartNo = MPART.PartNo         " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_Supplier MS ON PDU.SupplierID = MS.SupplierID          " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_Affiliate MA ON PDU.AffiliateID = MA.AffiliateID      " & vbCrLf

    '        ls_SQL = ls_SQL + "         LEFT JOIN dbo.MS_SupplierCapacity MSC ON PDU.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PDU.SupplierID=MSC.SupplierID          " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls                LEFT JOIN dbo.MS_CurrCls MCUR1 ON PDU.CurrCls = MCUR1.CurrCls          " & vbCrLf & _
    '                          "         LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf & _
    '                          "         WHERE PMU.SupplierID='" & Trim(cboSupplierCode.Text) & "'  " & vbCrLf & _
    '                          "            GROUP BY PMU.PONo,PDU.PONo,PDU.PartNo,PartName,MPART.KanbanCls,MU.Description,MOQ,QtyBox,PDU.poqty,MPART.Maker,MonthlyProductionCapacity       " & vbCrLf & _
    '                          "      		,PDU.CurrCls,MCUR1.Description,PDU.Price,PDU.Amount,MP.CurrCls,MCUR2.Description,MP.Price   " & vbCrLf & _
    '                          "       		,MSC.PartNo    " & vbCrLf & _
    '                          "      		,PDU.DeliveryD1,PDU.DeliveryD2,PDU.DeliveryD3,PDU.DeliveryD4,PDU.DeliveryD5,PDU.DeliveryD6,PDU.DeliveryD7,PDU.DeliveryD8,PDU.DeliveryD9,PDU.DeliveryD10          " & vbCrLf & _
    '                          "      		,PDU.DeliveryD11,PDU.DeliveryD12,PDU.DeliveryD13,PDU.DeliveryD14,PDU.DeliveryD15,PDU.DeliveryD16,PDU.DeliveryD17,PDU.DeliveryD18,PDU.DeliveryD19,PDU.DeliveryD20        		    " & vbCrLf & _
    '                          "      		,PDU.DeliveryD21,PDU.DeliveryD22,PDU.DeliveryD23,PDU.DeliveryD24,PDU.DeliveryD25,PDU.DeliveryD26,PDU.DeliveryD27,PDU.DeliveryD28,PDU.DeliveryD29,PDU.DeliveryD30,PDU.DeliveryD31     " & vbCrLf & _
    '                          " 	)detail1 " & vbCrLf

    '        ls_SQL = ls_SQL + " 	UNION ALL   " & vbCrLf & _
    '                          " 	SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
    '                          " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
    '                          "       ,MOQ = MOQ , QtyBox = QtyBox,Maker   " & vbCrLf & _
    '                          "       ,POQty     " & vbCrLf & _
    '                          "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15        ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf

    '        ls_SQL = ls_SQL + "       FROM (   " & vbCrLf & _
    '                          "       SELECT row_number() over (order by AD.PONo) as Sort ,'' as NoUrut ,'' PartNo ,AD.PartNo AS PartNos,'' PartName ,'' PONo, '' KanbanCls ,''Description ,'' MOQ,'' QtyBox ,AD.Maker   " & vbCrLf & _
    '                          "         ,POQtyOld POqty " & vbCrLf & _
    '                          "   		,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    		,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)                              " & vbCrLf & _
    '                          "    		,'AFTER' BYWHAT  " & vbCrLf & _
    '                          "      	,DeliveryD1Old DeliveryD1,DeliveryD2Old DeliveryD2,DeliveryD3Old DeliveryD3,DeliveryD4Old DeliveryD4,DeliveryD5Old DeliveryD5   " & vbCrLf & _
    '                          "   		,DeliveryD6Old DeliveryD6,DeliveryD7Old DeliveryD7,DeliveryD8Old DeliveryD8,DeliveryD9Old DeliveryD9,DeliveryD10Old DeliveryD10   " & vbCrLf & _
    '                          "   		,DeliveryD11Old DeliveryD11,DeliveryD12Old DeliveryD12,DeliveryD13Old DeliveryD13,DeliveryD14Old DeliveryD14,DeliveryD15Old DeliveryD15   " & vbCrLf & _
    '                          "   		,DeliveryD16Old DeliveryD16,DeliveryD17Old DeliveryD17,DeliveryD18 DeliveryD18,DeliveryD19Old DeliveryD19,DeliveryD20Old DeliveryD20   " & vbCrLf

    '        ls_SQL = ls_SQL + "   		,DeliveryD21Old DeliveryD21,DeliveryD22Old DeliveryD22,DeliveryD23Old DeliveryD23,DeliveryD24Old DeliveryD24,DeliveryD25Old DeliveryD25   " & vbCrLf & _
    '                          "   		,DeliveryD26Old DeliveryD26,DeliveryD27Old DeliveryD27,DeliveryD28Old DeliveryD28,DeliveryD29Old DeliveryD29,DeliveryD30Old DeliveryD30,DeliveryD31Old DeliveryD31   " & vbCrLf & _
    '                          "   		FROM dbo.AffiliateRev_Detail AD   " & vbCrLf & _
    '                          "   		LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo   " & vbCrLf & _
    '                          "   		LEFT JOIN dbo.MS_Price MP ON MP.PartNo = AD.PartNo and MP.AffiliateID = AD.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
    '                          "   		LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID          " & vbCrLf & _
    '                          "   		LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
    '                          "   		LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID          " & vbCrLf & _
    '                          "   		LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls          " & vbCrLf & _
    '                          "   		LEFT JOIN dbo.MS_CurrCls MCUR1 ON AD.CurrCls = MCUR1.CurrCls          " & vbCrLf & _
    '                          "   		LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf

    '        ls_SQL = ls_SQL + "         WHERE AD.PORevNo='" & Trim(txtRevisionNo.Text) & "' AND AD.SupplierID='" & Trim(cboSupplierCode.Text) & "'  " & vbCrLf & _
    '                          "   		GROUP BY PONo,AD.PartNo,PartName,AD.KanbanCls,POQtyOld,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity   ,SeqNo    " & vbCrLf & _
    '                          "  		,AD.CurrCls,MCUR1.Description,AD.Price,Amount,MP.CurrCls,MCUR2.Description,MP.Price,MSC.PartNo    " & vbCrLf & _
    '                          "      	,DeliveryD1Old,DeliveryD2Old,DeliveryD3Old,DeliveryD4Old,DeliveryD5Old   " & vbCrLf & _
    '                          "   		,DeliveryD6Old,DeliveryD7Old,DeliveryD8Old,DeliveryD9Old,DeliveryD10Old   " & vbCrLf & _
    '                          "   		,DeliveryD11Old,DeliveryD12Old,DeliveryD13Old,DeliveryD14Old,DeliveryD15Old   " & vbCrLf & _
    '                          "   		,DeliveryD16Old,DeliveryD17Old,DeliveryD18,DeliveryD19Old,DeliveryD20Old   " & vbCrLf & _
    '                          "   		,DeliveryD21Old,DeliveryD22Old,DeliveryD23Old,DeliveryD24Old,DeliveryD25Old   " & vbCrLf & _
    '                          "   		,DeliveryD26Old,DeliveryD27Old,DeliveryD28Old,DeliveryD29Old,DeliveryD30Old,DeliveryD31Old   " & vbCrLf & _
    '                          "   	 )detail2 " & vbCrLf & _
    '                          " 	 UNION ALL " & vbCrLf

    '        ls_SQL = ls_SQL + "   	 SELECT Sort,NoUrut , PartNo = PartNo ,PartNos,PartName = PartName ,PONo = PONo " & vbCrLf & _
    '                          " 	  ,POKanbanCls = KanbanCls ,Description    " & vbCrLf & _
    '                          "       ,MOQ = MOQ , QtyBox = QtyBox,Maker   " & vbCrLf & _
    '                          "       ,POQty     " & vbCrLf & _
    '                          "       ,ForecastN1 ,ForecastN2 ,ForecastN3,BYWHAT    " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15 " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
    '                          "       ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
    '                          "       FROM (   " & vbCrLf

    '        ls_SQL = ls_SQL + "       SELECT row_number() over (order by AD.PONo) as Sort ,'' as NoUrut ,'' PartNo ,AD.PartNo AS PartNos,'' PartName ,'' PONo, '' KanbanCls ,''Description ,'' MOQ,'' QtyBox ,AD.Maker   " & vbCrLf & _
    '                          "         ,POQtyOld POqty " & vbCrLf & _
    '                          "         ,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    		,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    		,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & tanggal & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & tanggal & "'))),0)    " & vbCrLf & _
    '                          "    		,'SUPPLIER APPROVAL' BYWHAT  " & vbCrLf & _
    '                          "  		,DeliveryD1 ,DeliveryD2 ,DeliveryD3 ,DeliveryD4 ,DeliveryD5    " & vbCrLf & _
    '                          "  		,DeliveryD6 ,DeliveryD7 ,DeliveryD8 ,DeliveryD9 ,DeliveryD10    " & vbCrLf & _
    '                          "  		,DeliveryD11 ,DeliveryD12 ,DeliveryD13 ,DeliveryD14    " & vbCrLf & _
    '                          "  		,DeliveryD15 ,DeliveryD16 ,DeliveryD17 ,DeliveryD18,DeliveryD19 ,DeliveryD20 ,DeliveryD21    " & vbCrLf & _
    '                          "  		,DeliveryD22 ,DeliveryD23 ,DeliveryD24 ,DeliveryD25 ,DeliveryD26 ,DeliveryD27 ,DeliveryD28 ,DeliveryD29    " & vbCrLf

    '        ls_SQL = ls_SQL + "  		,DeliveryD30 ,DeliveryD31    " & vbCrLf & _
    '                          "  		FROM dbo.AffiliateRev_Detail AD   " & vbCrLf & _
    '                          "  		 LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo   " & vbCrLf & _
    '                          "  		 LEFT JOIN dbo.MS_Price MP ON MP.PartNo = AD.PartNo and MP.AffiliateID = AD.AffiliateID and ('" & tanggal & "' between StartDate and EndDate)     " & vbCrLf & _
    '                          "  		 LEFT JOIN dbo.MS_Supplier MS ON AD.SupplierID = MS.SupplierID   " & vbCrLf & _
    '                          "  		 LEFT JOIN dbo.MS_Affiliate MA ON AD.AffiliateID = MA.AffiliateID      " & vbCrLf & _
    '                          "  		 LEFT JOIN dbo.MS_SupplierCapacity MSC ON AD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND AD.SupplierID=MSC.SupplierID          " & vbCrLf & _
    '                          "  		 LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls   " & vbCrLf & _
    '                          "  		 LEFT JOIN dbo.MS_CurrCls MCUR1 ON AD.CurrCls = MCUR1.CurrCls  " & vbCrLf & _
    '                          "  		 LEFT JOIN dbo.MS_CurrCls MCUR2 ON MP.CurrCls = MCUR2.CurrCls    " & vbCrLf & _
    '                          "         WHERE AD.PORevNo='" & Trim(txtRevisionNo.Text) & "' AND AD.SupplierID='" & Trim(cboSupplierCode.Text) & "'  " & vbCrLf

    '        ls_SQL = ls_SQL + "  		 GROUP BY PONo,AD.PartNo,PartName,AD.KanbanCls,POQty,MU.Description,MOQ,QtyBox,AD.Maker,MonthlyProductionCapacity ,POQtyOld,MSC.PartNo    " & vbCrLf & _
    '                          "      		,DeliveryD1,DeliveryD2,DeliveryD3,DeliveryD4,DeliveryD5   " & vbCrLf & _
    '                          "   			,DeliveryD6,DeliveryD7,DeliveryD8,DeliveryD9,DeliveryD10   " & vbCrLf & _
    '                          "   			,DeliveryD11,DeliveryD12,DeliveryD13,DeliveryD14,DeliveryD15   " & vbCrLf & _
    '                          "   			,DeliveryD16,DeliveryD17,DeliveryD18,DeliveryD19,DeliveryD20   " & vbCrLf & _
    '                          "   			,DeliveryD21,DeliveryD22,DeliveryD23,DeliveryD24,DeliveryD25   " & vbCrLf & _
    '                          "   			,DeliveryD26,DeliveryD27,DeliveryD28,DeliveryD29,DeliveryD30,DeliveryD31  " & vbCrLf & _
    '                          " 	)detail3 " & vbCrLf & _
    '                          " 	ORDER BY sort, PartNo DESC  " & vbCrLf & _
    '                          " END   " & vbCrLf & _
    '                          "    "



    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)
    '        Return ds
    '    End Using
    'End Function

    Private Function Supplier(ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID='" & ls_value & "'"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Function Affiliate(ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID='" & ls_value & "'"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    'Private Function EmailToEmailCC() As DataSet
    '    Dim ls_SQL As String = ""

    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()
    '        'ls_SQL = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID='" & ls_value & "'"

    '        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
    '                 " select 'AFF' flag,affiliatepocc, affiliatepoto='',toEmail='' from ms_emailaffiliate where AffiliateID='" & Trim(txtAffiliateID.Text) & "'" & vbCrLf & _
    '                 " union all " & vbCrLf & _
    '                 " --PASI TO -CC " & vbCrLf & _
    '                 " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI where AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
    '                 " union all " & vbCrLf & _
    '                 " --Supplier TO- CC " & vbCrLf & _
    '                 " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail='' from ms_emailSupplier where SupplierID='" & Trim(txtSupplierCode.Text) & "'"

    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)
    '        If ds.Tables(0).Rows.Count > 0 Then
    '            Return ds
    '        End If
    '    End Using
    'End Function

    Private Sub UpdateExcel(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffCode As String = "", _
                         Optional ByVal pRevNo As String = "", _
                         Optional ByVal pPONo As String = "", _
                         Optional ByVal pSuppCode As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " UPDATE dbo.AffiliateRev_Master " & vbCrLf & _
                          " SET ExcelCls='1'" & vbCrLf & _
                          " WHERE PORevNo = '" & pRevNo & "' " & vbCrLf & _
                          " AND PONo='" & pPONo & "'  " & vbCrLf & _
                          " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                          " AND SupplierID='" & pSuppCode & "' "

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Function getApp(ByVal pPORevNo As String, ByVal pPONo As String) As Boolean
        Dim ls_SQL As String = ""
        Dim doneApp As Boolean = False
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = " SELECT * FROM PORev_Master " & vbCrLf & _
                  " WHERE PORevNo='" & pPORevNo & "' AND PONO='" & pPONo & "' AND  " & vbCrLf & _
                  " (ISNULL(SupplierApproveDate,'') <> '' OR " & vbCrLf & _
                  " ISNULL(SupplierApprovePendingDate,'') <> '' OR  " & vbCrLf & _
                  " ISNULL(SupplierUnApproveDate,'') <> '' OR  " & vbCrLf & _
                  " ISNULL(PASIApproveDate,'') <> '' OR  " & vbCrLf & _
                  " ISNULL(FinalApproveDate,'') <> '') "
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                doneApp = True
            End If
        End Using
        Return doneApp
    End Function
#End Region

End Class