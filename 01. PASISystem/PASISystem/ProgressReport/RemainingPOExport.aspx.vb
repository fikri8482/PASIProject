Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports DevExpress.Web.ASPxMenu
Imports OfficeOpenXml
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.Net

Public Class RemainingPOExport1
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "M00"

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim dtHeader As DataTable
    Dim dtDetail As DataTable
#End Region

#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_fillcombo()
                dtPeriodFrom1.Value = Now
                dtPeriodTo1.Value = Now
                txtAffiliateName.Text = "==ALL=="
                txtSupplierName.Text = "==ALL=="
                txtForwarder.Text = "==ALL=="
                txtPartName.Text = "==ALL=="
                lblErrMsg.Text = ""
            End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "PROCEDURE"
    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""

        'Affiliate ID
        ls_sql = "SELECT [Affiliate Code] = '==ALL==' , [Affiliate Name] = '==ALL==' UNION ALL SELECT [Affiliate Code] = RTRIM(AffiliateID) ,[Affiliate Name] = RTRIM(AffiliateName) FROM MS_Affiliate " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Affiliate Code")
                .Columns(0).Width = 100
                .Columns.Add("Affiliate Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0

                .TextField = "Affiliate Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Forwarder ID
        ls_sql = "SELECT [Forwarder Code] = '==ALL==' , [Forwarder Name] = '==ALL==' UNION ALL SELECT [Forwarder Code] = RTRIM(ForwarderID) ,[Forwarder Name] = RTRIM(ForwarderName) FROM MS_Forwarder " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboForwarder
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Forwarder Code")
                .Columns(0).Width = 100
                .Columns.Add("Forwarder Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0

                .TextField = "Forwarder Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Supplier Code
        ls_sql = "SELECT [Supplier Code] = '==ALL==' , [Supplier Name] = '==ALL==' UNION ALL SELECT [Supplier Code] = RTRIM(supplierID) ,[Supplier Name] = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Supplier Code")
                .Columns(0).Width = 100
                .Columns.Add("Supplier Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0

                .TextField = "Supplier Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Part No
        ls_sql = "SELECT [Part No] = '==ALL==' , [Part Name] = '==ALL==' UNION ALL SELECT [Part No] = RTRIM(PartNo) ,[Part Name] = RTRIM(PartName) FROM MS_Parts where FinishGoodCls = '2' " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Part No")
                .Columns(0).Width = 100
                .Columns.Add("Part Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0

                .TextField = "Part No"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            'ls_SQL = "  SELECT *,CONVERT(CHAR,row_number() over (order by OrderNo)) as NoUrut FROM  " & vbCrLf & _
            '      " ( " & vbCrLf & _
            '      " 	SELECT '1' Urut," & vbCrLf & _
            '      " 		PONo, AffiliateID = CONVERT(CHAR,AffiliateID), SupplierID = CONVERT(CHAR,SupplierID), " & vbCrLf & _
            '      " 		ForwarderID = CONVERT(CHAR,ForwarderID), Period = CONVERT(CHAR(7),Period), EmergencyCls, OrderSeq, " & vbCrLf & _
            '      " 		CASE WHEN OrderSeq = 'OrderNo1' then ISNULL(OrderNo1,'') " & vbCrLf & _
            '      " 			 WHEN OrderSeq = 'OrderNo2' then ISNULL(OrderNo2,'') " & vbCrLf & _
            '      " 			 WHEN OrderSeq = 'OrderNo3' then ISNULL(OrderNo3,'') " & vbCrLf & _
            '      " 			 WHEN OrderSeq = 'OrderNo4' then ISNULL(OrderNo4,'') " & vbCrLf & _
            '      " 			 WHEN OrderSeq = 'OrderNo5' then ISNULL(OrderNo5,'') " & vbCrLf & _
            '      " 		END OrderNo, " & vbCrLf

            'ls_SQL = ls_SQL + " 		CASE WHEN OrderSeq = 'OrderNo1' then ISNULL(CONVERT(varchar,ETDVendor1),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo2' then ISNULL(CONVERT(varchar,ETDVendor2),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo3' then ISNULL(CONVERT(varchar,ETDVendor3),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo4' then ISNULL(CONVERT(varchar,ETDVendor4),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo5' then ISNULL(CONVERT(varchar,ETDVendor5),'') " & vbCrLf & _
            '                  " 		END ETDVendor, " & vbCrLf & _
            '                  " 		CASE WHEN OrderSeq = 'OrderNo1' then ISNULL(CONVERT(varchar,ETDPort1),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo2' then ISNULL(CONVERT(varchar,ETDPort2),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo3' then ISNULL(CONVERT(varchar,ETDPort3),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo4' then ISNULL(CONVERT(varchar,ETDPort4),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo5' then ISNULL(CONVERT(varchar,ETDPort5),'') " & vbCrLf

            'ls_SQL = ls_SQL + " 		END ETDPort, '' PartNo, '' OrderQty, " & vbCrLf & _
            '                  " 		'' SuppDelQty, '' GoodRecQty, '' DefRecQty, '' RemainRecQty " & vbCrLf & _
            '                  " 	FROM " & vbCrLf & _
            '                  " 	( " & vbCrLf & _
            '                  " 		SELECT * from " & vbCrLf & _
            '                  " 		( " & vbCrLf & _
            '                  " 			SELECT 'OrderNo1' OrderSeq " & vbCrLf & _
            '                  " 			union all " & vbCrLf & _
            '                  " 			SELECT 'OrderNo2' OrderSeq " & vbCrLf & _
            '                  " 			union all " & vbCrLf & _
            '                  " 			SELECT 'OrderNo3' OrderSeq " & vbCrLf

            'ls_SQL = ls_SQL + " 			union all " & vbCrLf & _
            '                  " 			SELECT 'OrderNo4' OrderSeq " & vbCrLf & _
            '                  " 			union all " & vbCrLf & _
            '                  " 			SELECT 'OrderNo5' OrderSeq " & vbCrLf & _
            '                  " 		)x cross join " & vbCrLf & _
            '                  " 		( " & vbCrLf & _
            '                  " 			SELECT *  " & vbCrLf & _
            '                  " 			FROM PO_Master_Export a " & vbCrLf & _
            '                  " 			--left join PO_Master_Export b " & vbCrLf & _
            '                  " 			WHERE Period BETWEEN '" & Format(dtPeriodFrom1.Value, "yyyy-MM-01") & "' AND '" & Format(dtPeriodTo1.Value, "yyyy-MM-01") & "' --PONo = '12346' " & vbCrLf

            ''AffiliateID
            'If Trim(cboAffiliateCode.Text) <> "" And Trim(cboAffiliateCode.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            'End If

            ''SupplierCode
            'If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            'End If

            ''Forwarder
            'If Trim(cboForwarder.Text) <> "" And Trim(cboForwarder.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND ForwarderID = '" & Trim(cboForwarder.Text) & "' " & vbCrLf
            'End If

            ''EmergencyCls
            'If rbEmergency.Checked = True Then
            '    ls_SQL = ls_SQL + _
            '        "           AND EmergencyCls = 'E' " & vbCrLf
            'End If

            ''EmergencyCls
            'If rbMonthly.Checked = True Then
            '    ls_SQL = ls_SQL + _
            '        "           AND EmergencyCls = 'M' " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + " 		)y " & vbCrLf & _
            '                  " 	)z " & vbCrLf & _
            '                  " )w where OrderNo <> '' " & vbCrLf & _
            '                  "  " & vbCrLf & _
            '                  " UNION ALL " & vbCrLf & _
            '                  " SELECT * FROM " & vbCrLf & _
            '                  " ( " & vbCrLf & _
            '                  " 	SELECT DISTINCT '2' Urut, " & vbCrLf

            'ls_SQL = ls_SQL + " 		PONo, AffiliateID = '', SupplierID = '', " & vbCrLf & _
            '                  " 		ForwarderID = '', Period = '', '' EmergencyCls, OrderSeq, " & vbCrLf & _
            '                  " 		CASE WHEN OrderSeq = 'OrderNo1' then ISNULL(OrderNo1,'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo2' then ISNULL(OrderNo2,'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo3' then ISNULL(OrderNo3,'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo4' then ISNULL(OrderNo4,'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo5' then ISNULL(OrderNo5,'') " & vbCrLf & _
            '                  " 		END OrderNo, " & vbCrLf & _
            '                  " 		'' ETDVendor, " & vbCrLf & _
            '                  " 		'' ETDPort, PartNo,  " & vbCrLf & _
            '                  " 		CASE WHEN OrderSeq = 'OrderNo1' then ISNULL(CONVERT(varchar,Week1),'') " & vbCrLf

            'ls_SQL = ls_SQL + " 			 WHEN OrderSeq = 'OrderNo2' then ISNULL(CONVERT(varchar,Week2),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo3' then ISNULL(CONVERT(varchar,Week3),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo4' then ISNULL(CONVERT(varchar,Week4),'') " & vbCrLf & _
            '                  " 			 WHEN OrderSeq = 'OrderNo5' then ISNULL(CONVERT(varchar,Week5),'') " & vbCrLf & _
            '                  " 		END OrderQty, " & vbCrLf & _
            '                  " 		SuppDelQty = ISNULL(CONVERT(CHAR,DOQty),0), GoodRecQty = ISNULL(CONVERT(CHAR,GoodRecQty),0), " & vbCrLf & _
            '                  " 		DefRecQty = ISNULL(CONVERT(CHAR,DefectRecQty),0), RemainRecQty =ISNULL(CONVERT(CHAR,RemainRecQty),0),'' NoUrut " & vbCrLf & _
            '                  " 	from " & vbCrLf & _
            '                  " 	( " & vbCrLf & _
            '                  " 		SELECT * from " & vbCrLf & _
            '                  " 		( " & vbCrLf

            'ls_SQL = ls_SQL + " 			SELECT 'OrderNo1' OrderSeq " & vbCrLf & _
            '                  " 			union all " & vbCrLf & _
            '                  " 			SELECT 'OrderNo2' OrderSeq " & vbCrLf & _
            '                  " 			union all " & vbCrLf & _
            '                  " 			SELECT 'OrderNo3' OrderSeq " & vbCrLf & _
            '                  " 			union all " & vbCrLf & _
            '                  " 			SELECT 'OrderNo4' OrderSeq " & vbCrLf & _
            '                  " 			union all " & vbCrLf & _
            '                  " 			SELECT 'OrderNo5' OrderSeq " & vbCrLf & _
            '                  " 		)x cross join " & vbCrLf & _
            '                  " 		( " & vbCrLf

            'ls_SQL = ls_SQL + " 			SELECT PONo, AffiliateID, SupplierID, PartNo, " & vbCrLf & _
            '                  " 				Week1 = Convert(Numeric(18,0),Week1), " & vbCrLf & _
            '                  " 				Week2 = Convert(Numeric(18,0),Week2), " & vbCrLf & _
            '                  " 				Week3 = Convert(Numeric(18,0),Week3), " & vbCrLf & _
            '                  " 				Week4 = Convert(Numeric(18,0),Week4), " & vbCrLf & _
            '                  " 				Week5 = Convert(Numeric(18,0),Week5), " & vbCrLf & _
            '                  " 				TotalPOQty = Convert(Numeric(18,0),TotalPOQty),  " & vbCrLf & _
            '                  " 				OrderNo1, OrderNo2, OrderNo3, OrderNo4, OrderNo5, " & vbCrLf & _
            '                  " 				DOQty = Convert(Numeric(18,0),DOQty), GoodRecQty = Convert(Numeric(18,0),GoodRecQty), DefectRecQty = Convert(Numeric(18,0),DefectRecQty),  " & vbCrLf & _
            '                  " 				RemainRecQty = ((Convert(Numeric(18,0),GoodRecQty) + Convert(Numeric(18,0),DefectRecQty)) - Convert(Numeric(18,0),DOQty))  " & vbCrLf & _
            '                  " 			FROM( " & vbCrLf & _
            '                  " 			SELECT a.PONo, a.AffiliateID, a.SupplierID, a.PartNo, " & vbCrLf & _
            '                  " 				Week1, Week2, Week3, Week4, Week5, TotalPOQty, " & vbCrLf & _
            '                  " 				OrderNo1, OrderNo2, OrderNo3, OrderNo4, OrderNo5, " & vbCrLf & _
            '                  " 				DOQty, GoodRecQty, DefectRecQty " & vbCrLf & _
            '                  " 			FROM PO_Detail_Export a " & vbCrLf

            'ls_SQL = ls_SQL + " 			left join Po_master_Export b on a.pono=b.pono " & vbCrLf & _
            '                  " 			and a.affiliateID = b.AffiliateID " & vbCrLf & _
            '                  " 			left join dbo.DOSupplier_Detail_Export c on (c.OrderNo = b.OrderNo1 OR " & vbCrLf & _
            '                  " 					c.OrderNo = b.OrderNo2 OR " & vbCrLf & _
            '                  " 					c.OrderNo = b.OrderNo3 OR " & vbCrLf & _
            '                  " 					c.OrderNo = b.OrderNo4 OR " & vbCrLf & _
            '                  " 					c.OrderNo = b.OrderNo5) " & vbCrLf & _
            '                  " 					AND c.PartNo = a.PartNo " & vbCrLf & _
            '                  " 			left join dbo.ReceiveForwarder_Detail d on (d.OrderNo = b.OrderNo1 OR " & vbCrLf & _
            '                  " 					d.OrderNo = b.OrderNo2 OR " & vbCrLf & _
            '                  " 					d.OrderNo = b.OrderNo3 OR " & vbCrLf

            'ls_SQL = ls_SQL + " 					d.OrderNo = b.OrderNo4 OR " & vbCrLf & _
            '                  " 					d.OrderNo = b.OrderNo5) " & vbCrLf & _
            '                  " 					AND d.PartNo = a.PartNo " & vbCrLf & _
            '                  " 			WHERE Period BETWEEN '" & Format(dtPeriodFrom1.Value, "yyyy-MM-01") & "' AND '" & Format(dtPeriodTo1.Value, "yyyy-MM-01") & "' --PONo = '12346' " & vbCrLf

            ''AffiliateID
            'If Trim(cboAffiliateCode.Text) <> "" And Trim(cboAffiliateCode.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND a.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            'End If

            ''SupplierCode
            'If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND b.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            'End If

            ''Forwarder
            'If Trim(cboForwarder.Text) <> "" And Trim(cboForwarder.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND ForwarderID = '" & Trim(cboForwarder.Text) & "' " & vbCrLf
            'End If

            ''EmergencyCls
            'If rbEmergency.Checked = True Then
            '    ls_SQL = ls_SQL + _
            '        "           AND EmergencyCls = 'E' " & vbCrLf
            'End If

            ''EmergencyCls
            'If rbMonthly.Checked = True Then
            '    ls_SQL = ls_SQL + _
            '        "           AND EmergencyCls = 'M' " & vbCrLf
            'End If

            ''PartNo
            'If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND a.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
            'End If

            ''PartNo
            'If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND a.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
            'End If
            'ls_SQL = ls_SQL + ") XX " & vbCrLf & _
            '                  " 		)y "

            'ls_SQL = ls_SQL + " 	)z " & vbCrLf & _
            '                  " )w WHERE OrderNo <> '' " & vbCrLf & _
            '                  " ORDER BY 9,1 "

            '8 Desember 2015

            ls_SQL = "  " & vbCrLf & _
                  "  	SELECT '1' Urut, " & vbCrLf & _
                  "  		PONo, AffiliateID = CONVERT(CHAR,AffiliateID), SupplierID = CONVERT(CHAR,SupplierID),  " & vbCrLf & _
                  "  		ForwarderID = CONVERT(CHAR,ForwarderID), " & vbCrLf & _
                  "  		Period = CONVERT(CHAR(7),Period), EmergencyCls,  " & vbCrLf & _
                  "  		OrderNo = OrderNo1,  " & vbCrLf & _
                  "  		ETDVendor = ISNULL(CONVERT(varchar,ETDVendor1),''), " & vbCrLf & _
                  "  		ETDPort = ISNULL(CONVERT(varchar,ETDPort1),''), '' PartNo, '' OrderQty,  " & vbCrLf & _
                  "  		'' SuppDelQty, '' GoodRecQty, '' DefRecQty, '' RemainRecQty, CONVERT(CHAR,row_number() over (order by OrderNo1)) as NoUrut   " & vbCrLf & _
                  "  		FROM PO_Master_Export a   " & vbCrLf & _
                  "  		WHERE Period BETWEEN '" & Format(dtPeriodFrom1.Value, "yyyy-MM-01") & "' AND '" & Format(dtPeriodTo1.Value, "yyyy-MM-01") & "' --PONo = '12346' " & vbCrLf

            'AffiliateID
            If Trim(cboAffiliateCode.Text) <> "" And Trim(cboAffiliateCode.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND a.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            End If

            'SupplierCode
            If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            End If

            'Forwarder
            If Trim(cboForwarder.Text) <> "" And Trim(cboForwarder.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND ForwarderID = '" & Trim(cboForwarder.Text) & "' " & vbCrLf
            End If

            'EmergencyCls
            If rbEmergency.Checked = True Then
                ls_SQL = ls_SQL + _
                    "           AND EmergencyCls = 'E' " & vbCrLf
            End If

            'EmergencyCls
            If rbMonthly.Checked = True Then
                ls_SQL = ls_SQL + _
                    "           AND EmergencyCls = 'M' " & vbCrLf
            End If

            ''PartNo
            'If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND a.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
            'End If

            'OrderNo
            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "            AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "    " & vbCrLf & _
                              "  UNION ALL  " & vbCrLf & _
                              "  SELECT * FROM  " & vbCrLf & _
                              "  (  " & vbCrLf & _
                              "  	SELECT DISTINCT '2' Urut,  " & vbCrLf & _
                              "  		PONo, AffiliateID = '', SupplierID = '',  " & vbCrLf & _
                              "  		ForwarderID = '', Period = '', '' EmergencyCls, " & vbCrLf & _
                              "  		OrderNo = ISNULL(OrderNo1,''),  " & vbCrLf & _
                              "  		ETDVendor = '',  " & vbCrLf & _
                              "  		ETDPort = '', PartNo,  "

            ls_SQL = ls_SQL + "  		OrderQty = ISNULL(CONVERT(varchar,Week1),''),  " & vbCrLf & _
                              "  		SuppDelQty = ISNULL(CONVERT(CHAR,DOQty),0),  " & vbCrLf & _
                              "  		GoodRecQty = ISNULL(CONVERT(CHAR,GoodRecQty),0),  " & vbCrLf & _
                              "  		DefRecQty = ISNULL(CONVERT(CHAR,DefectRecQty),0),  " & vbCrLf & _
                              "  		RemainRecQty = ISNULL(CONVERT(CHAR,RemainRecQuantity),0),'' NoUrut  " & vbCrLf & _
                              "  	from  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		 " & vbCrLf & _
                              "  			SELECT a.PONo, a.AffiliateID, a.SupplierID, a.PartNo,  " & vbCrLf & _
                              "  				Week1, Week2, Week3, Week4, Week5, TotalPOQty,  " & vbCrLf & _
                              "  				OrderNo1, OrderNo2, OrderNo3, OrderNo4, OrderNo5,  "

            ls_SQL = ls_SQL + "  				DOQty = Convert(Numeric(18,0),DOQty), GoodRecQty = Convert(Numeric(18,0),GoodRecQty), " & vbCrLf & _
                              "  				DefectRecQty = Convert(Numeric(18,0),DefectRecQty), " & vbCrLf & _
                              "  				RemainRecQuantity = ((Convert(Numeric(18,0),GoodRecQty) + Convert(Numeric(18,0),DefectRecQty)) - Convert(Numeric(18,0),DOQty))  " & vbCrLf & _
                              "  			FROM PO_Detail_Export a  " & vbCrLf & _
                              "  			left join Po_master_Export b on a.pono=b.pono  " & vbCrLf & _
                              "  			and a.affiliateID = b.AffiliateID  " & vbCrLf & _
                              "  			left join dbo.DOSupplier_Detail_Export c on c.OrderNo = b.OrderNo1  " & vbCrLf & _
                              "  					AND c.PartNo = a.PartNo  " & vbCrLf & _
                              "  			left join dbo.ReceiveForwarder_Detail d on d.OrderNo = b.OrderNo1 " & vbCrLf & _
                              "  					AND d.PartNo = a.PartNo  " & vbCrLf & _
                              "  			WHERE Period BETWEEN '" & Format(dtPeriodFrom1.Value, "yyyy-MM-01") & "' AND '" & Format(dtPeriodTo1.Value, "yyyy-MM-01") & "' " & vbCrLf

            'AffiliateID
            If Trim(cboAffiliateCode.Text) <> "" And Trim(cboAffiliateCode.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND a.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            End If

            'SupplierCode
            If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND b.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            End If

            'Forwarder
            If Trim(cboForwarder.Text) <> "" And Trim(cboForwarder.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND ForwarderID = '" & Trim(cboForwarder.Text) & "' " & vbCrLf
            End If

            'EmergencyCls
            If rbEmergency.Checked = True Then
                ls_SQL = ls_SQL + _
                    "           AND EmergencyCls = 'E' " & vbCrLf
            End If

            'EmergencyCls
            If rbMonthly.Checked = True Then
                ls_SQL = ls_SQL + _
                    "           AND EmergencyCls = 'M' " & vbCrLf
            End If

            'PartNo
            If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND a.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
            End If

            'OrderNo
            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "           AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "  --PONo = '12346'  " & vbCrLf & _
                              " ) XX  " & vbCrLf & _
                              "  		)y  " & vbCrLf & _
                              "  		WHERE OrderNo <> ''  " & vbCrLf & _
                              "  ORDER BY 8,1  "

            
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

        End Using
    End Sub


    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction

                Case "loaddata"

                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblErrMsg, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblErrMsg.Text
                    Else
                        lblErrMsg.Text = ""
                        grid.JSProperties("cpMessage") = ""
                        grid.FocusedRowIndex = -1
                    End If

                Case "excelremaining"
                    Call ExcelRemainingPO()

            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblErrMsg, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
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

    Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
        With Rg
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
    End Sub

    Private Sub GridLoadExcel()

        Dim ds As New DataSet
        Dim ls_sql As String = ""

        ls_SQL = "  " & vbCrLf & _
                  "  	SELECT '1' Urut, " & vbCrLf & _
                  "  		PONo, AffiliateID = CONVERT(CHAR,AffiliateID), SupplierID = CONVERT(CHAR,SupplierID),  " & vbCrLf & _
                  "  		ForwarderID = CONVERT(CHAR,ForwarderID), " & vbCrLf & _
                  "  		Period = CONVERT(CHAR(7),Period), EmergencyCls,  " & vbCrLf & _
                  "  		OrderNo = OrderNo1,  " & vbCrLf & _
                  "  		ETDVendor = ISNULL(CONVERT(varchar,ETDVendor1),''), " & vbCrLf & _
                  "  		ETDPort = ISNULL(CONVERT(varchar,ETDPort1),''), '' PartNo, '' OrderQty,  " & vbCrLf & _
                  "  		'' SuppDelQty, '' GoodRecQty, '' DefRecQty, '' RemainRecQty, CONVERT(CHAR,row_number() over (order by OrderNo1)) as NoUrut   " & vbCrLf & _
                  "  		FROM PO_Master_Export a   " & vbCrLf & _
                  "  		WHERE Period BETWEEN '" & Format(dtPeriodFrom1.Value, "yyyy-MM-01") & "' AND '" & Format(dtPeriodTo1.Value, "yyyy-MM-01") & "' --PONo = '12346' " & vbCrLf

        'AffiliateID
        If Trim(cboAffiliateCode.Text) <> "" And Trim(cboAffiliateCode.Text) <> "==ALL==" Then
            ls_sql = ls_sql + _
                "           AND a.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
        End If

        'SupplierCode
        If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
            ls_sql = ls_sql + _
                "           AND SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
        End If

        'Forwarder
        If Trim(cboForwarder.Text) <> "" And Trim(cboForwarder.Text) <> "==ALL==" Then
            ls_sql = ls_sql + _
                "           AND ForwarderID = '" & Trim(cboForwarder.Text) & "' " & vbCrLf
        End If

        'EmergencyCls
        If rbEmergency.Checked = True Then
            ls_sql = ls_sql + _
                "           AND EmergencyCls = 'E' " & vbCrLf
        End If

        'EmergencyCls
        If rbMonthly.Checked = True Then
            ls_sql = ls_sql + _
                "           AND EmergencyCls = 'M' " & vbCrLf
        End If

        ''PartNo
        'If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
        '    ls_SQL = ls_SQL + _
        '        "           AND a.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
        'End If

        'OrderNo
        If Trim(txtOrderNo.Text) <> "" Then
            ls_sql = ls_sql + _
                "            AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
        End If

        ls_sql = ls_sql + "    " & vbCrLf & _
                          "  UNION ALL  " & vbCrLf & _
                          "  SELECT * FROM  " & vbCrLf & _
                          "  (  " & vbCrLf & _
                          "  	SELECT DISTINCT '2' Urut,  " & vbCrLf & _
                          "  		PONo, AffiliateID = '', SupplierID = '',  " & vbCrLf & _
                          "  		ForwarderID = '', Period = '', '' EmergencyCls, " & vbCrLf & _
                          "  		OrderNo = ISNULL(OrderNo1,''),  " & vbCrLf & _
                          "  		ETDVendor = '',  " & vbCrLf & _
                          "  		ETDPort = '', PartNo,  "

        ls_sql = ls_sql + "  		OrderQty = ISNULL(CONVERT(varchar,Week1),''),  " & vbCrLf & _
                          "  		SuppDelQty = ISNULL(CONVERT(CHAR,DOQty),0),  " & vbCrLf & _
                          "  		GoodRecQty = ISNULL(CONVERT(CHAR,GoodRecQty),0),  " & vbCrLf & _
                          "  		DefRecQty = ISNULL(CONVERT(CHAR,DefectRecQty),0),  " & vbCrLf & _
                          "  		RemainRecQty = ISNULL(CONVERT(CHAR,RemainRecQuantity),0),'' NoUrut  " & vbCrLf & _
                          "  	from  " & vbCrLf & _
                          "  	(  " & vbCrLf & _
                          "  		 " & vbCrLf & _
                          "  			SELECT a.PONo, a.AffiliateID, a.SupplierID, a.PartNo,  " & vbCrLf & _
                          "  				Week1, Week2, Week3, Week4, Week5, TotalPOQty,  " & vbCrLf & _
                          "  				OrderNo1, OrderNo2, OrderNo3, OrderNo4, OrderNo5,  "

        ls_sql = ls_sql + "  				DOQty = Convert(Numeric(18,0),DOQty), GoodRecQty = Convert(Numeric(18,0),GoodRecQty), " & vbCrLf & _
                          "  				DefectRecQty = Convert(Numeric(18,0),DefectRecQty), " & vbCrLf & _
                          "  				RemainRecQuantity = ((Convert(Numeric(18,0),GoodRecQty) + Convert(Numeric(18,0),DefectRecQty)) - Convert(Numeric(18,0),DOQty))  " & vbCrLf & _
                          "  			FROM PO_Detail_Export a  " & vbCrLf & _
                          "  			left join Po_master_Export b on a.pono=b.pono  " & vbCrLf & _
                          "  			and a.affiliateID = b.AffiliateID  " & vbCrLf & _
                          "  			left join dbo.DOSupplier_Detail_Export c on c.OrderNo = b.OrderNo1  " & vbCrLf & _
                          "  					AND c.PartNo = a.PartNo  " & vbCrLf & _
                          "  			left join dbo.ReceiveForwarder_Detail d on d.OrderNo = b.OrderNo1 " & vbCrLf & _
                          "  					AND d.PartNo = a.PartNo  " & vbCrLf & _
                          "  			WHERE Period BETWEEN '" & Format(dtPeriodFrom1.Value, "yyyy-MM-01") & "' AND '" & Format(dtPeriodTo1.Value, "yyyy-MM-01") & "' " & vbCrLf

        'AffiliateID
        If Trim(cboAffiliateCode.Text) <> "" And Trim(cboAffiliateCode.Text) <> "==ALL==" Then
            ls_sql = ls_sql + _
                "           AND a.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
        End If

        'SupplierCode
        If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
            ls_sql = ls_sql + _
                "           AND b.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
        End If

        'Forwarder
        If Trim(cboForwarder.Text) <> "" And Trim(cboForwarder.Text) <> "==ALL==" Then
            ls_sql = ls_sql + _
                "           AND ForwarderID = '" & Trim(cboForwarder.Text) & "' " & vbCrLf
        End If

        'EmergencyCls
        If rbEmergency.Checked = True Then
            ls_sql = ls_sql + _
                "           AND EmergencyCls = 'E' " & vbCrLf
        End If

        'EmergencyCls
        If rbMonthly.Checked = True Then
            ls_sql = ls_sql + _
                "           AND EmergencyCls = 'M' " & vbCrLf
        End If

        'PartNo
        If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
            ls_sql = ls_sql + _
                "           AND a.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
        End If

        'OrderNo
        If Trim(txtOrderNo.Text) <> "" Then
            ls_sql = ls_sql + _
                "           AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
        End If

        ls_sql = ls_sql + "  --PONo = '12346'  " & vbCrLf & _
                          " ) XX  " & vbCrLf & _
                          "  		)y  " & vbCrLf & _
                          "  		WHERE OrderNo <> ''  " & vbCrLf & _
                          "  ORDER BY 8,1  "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using
        dtHeader = ds.Tables(0)
    End Sub

    Private Sub epplusExportHeaderExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try

            Dim NewFileName As String = Server.MapPath("~\ProgressReport\TemplateRemainingPOExportReport.xlsx")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet

            ws = exl.Workbook.Worksheets("RemainingPOExportReport")
            Dim irow As Integer = 0
            Dim iRowTmp As Integer = 0
            Dim icol As Integer = 0

            iRowTmp = 5
            For irow = 0 To pData.Rows.Count - 1
                If pData.Rows.Count > 0 Then
                    ws.Cells("A" & iRowTmp).Value = pData.Rows(irow)("NoUrut")
                    ws.Cells("B" & iRowTmp).Value = pData.Rows(irow)("Period")
                    ws.Cells("C" & iRowTmp).Value = pData.Rows(irow)("AffiliateID")
                    ws.Cells("D" & iRowTmp).Value = pData.Rows(irow)("ForwarderID")
                    ws.Cells("E" & iRowTmp).Value = pData.Rows(irow)("OrderNo")
                    ws.Cells("F" & iRowTmp).Value = pData.Rows(irow)("EmergencyCls")
                    ws.Cells("G" & iRowTmp).Value = pData.Rows(irow)("SupplierID")
                    ws.Cells("H" & iRowTmp).Value = pData.Rows(irow)("ETDVendor")
                    ws.Cells("I" & iRowTmp).Value = pData.Rows(irow)("ETDPort")
                    ws.Cells("J" & iRowTmp).Value = pData.Rows(irow)("PartNo")
                    ws.Cells("K" & iRowTmp).Value = pData.Rows(irow)("OrderQty")
                    ws.Cells("L" & iRowTmp).Value = pData.Rows(irow)("SuppDelQty")
                    ws.Cells("M" & iRowTmp).Value = pData.Rows(irow)("GoodRecQty")
                    ws.Cells("N" & iRowTmp).Value = pData.Rows(irow)("DefRecQty")
                    ws.Cells("O" & iRowTmp).Value = pData.Rows(irow)("RemainRecQty")

                    'ALIGNMENT
                    ws.Cells("A" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                    ws.Cells("B" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("C" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("D" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("E" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("F" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("G" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("H" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("I" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("J" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("K" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("L" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("M" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("N" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("O" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'FORMAT
                    ws.Cells("A" & iRowTmp).Style.Numberformat.Format = "dd MMM yyyy"
                    ws.Cells("H" & iRowTmp).Style.Numberformat.Format = "dd MMM yyyy"
                    ws.Cells("I" & iRowTmp).Style.Numberformat.Format = "dd MMM yyyy"

                    'WIDTH
                    ws.Column(2).Width = 11
                End If
                iRowTmp = iRowTmp + 1
            Next

            Dim rgAll As ExcelRange = ws.Cells(5, 1, iRowTmp - 1, 15)
            EpPlusDrawAllBorders(rgAll)

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub ExcelRemainingPO()
        Call GridLoadExcel()
        FileName = "TemplateRemainingPOExportReport.xlsx"
        FilePath = Server.MapPath("~\Template\" & FileName)
        Call epplusExportHeaderExcel(FilePath, "", dtHeader, "A:5", "")
    End Sub
#End Region
End Class