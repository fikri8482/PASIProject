Imports System.Data.SqlClient
Imports OfficeOpenXml
Imports System.IO

Public Class DeliveryPerformReport
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
#End Region

    Private Sub up_fillcombo()
        Dim ls_sql As String

        dtPeriod.Value = Date.Now

        ls_sql = ""
        'SUPPLIER
        ls_sql = "/*SELECT SupplierGroupCode = '" & clsGlobal.gs_All & "', Description = '" & clsGlobal.gs_All & "' " & vbCrLf & _
                 "UNION */" & vbCrLf & _
                 "SELECT SupplierGroupCode,Description FROM dbo.MS_SupplierGroup " & vbCrLf

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierGroup
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierGroupCode")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 240
                .SelectedIndex = -1
                .TextField = "SupplierGroupCode"
                .DataBind()
            End With
            txtSupplierGroup.Text = ""
            sqlConn.Close()
        End Using

        'PART
        ls_sql = "/*SELECT PartNo = '" & clsGlobal.gs_All & "', PartName = '" & clsGlobal.gs_All & "' " & vbCrLf & _
                 "UNION*/ " & vbCrLf & _
                 "SELECT PartNo,PartName FROM dbo.MS_Parts " & vbCrLf

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 70
                .Columns.Add("PartName")
                .Columns(1).Width = 240
                .SelectedIndex = -1

                .TextField = "PartNo"
                .DataBind()
            End With
            txtPartCode.Text = ""
            sqlConn.Close()
        End Using

        'AFFILIATE
        ls_sql = "/*SELECT AffiliateID = '" & clsGlobal.gs_All & "', AffiliateName = '" & clsGlobal.gs_All & "' " & vbCrLf & _
                 "UNION*/ " & vbCrLf & _
                 "SELECT AffiliateID,AffiliateName FROM MS_Affiliate " & vbCrLf

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 70
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240
                .SelectedIndex = -1

                .TextField = "AffiliateID"
                .DataBind()
            End With
            txtAffiliateCode.Text = ""
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  SELECT No = '1',  " & vbCrLf & _
                              "  	   PerformanceCls = 'EXPECTATION',  " & vbCrLf & _
                              "  	   PlanActual = 'PLAN',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0 ,PFC = 'EXP' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '',  " & vbCrLf & _
                              "  	   PerformanceCls = '',  " & vbCrLf & _
                              "  	   PlanActual = 'ACTUAL',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0 ,PFC = 'EXP' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '2',  " & vbCrLf & _
                              "  	   PerformanceCls = 'DELAY',  " & vbCrLf & _
                              "  	   PlanActual = 'PLAN',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0  ,PFC = 'DLY' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '',  " & vbCrLf & _
                              "  	   PerformanceCls = '',  " & vbCrLf & _
                              "  	   PlanActual = 'ACTUAL',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0 ,PFC = 'DLY' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '3',  " & vbCrLf & _
                              "  	   PerformanceCls = 'PARTIAL',  " & vbCrLf & _
                              "  	   PlanActual = 'PLAN',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0  ,PFC = 'PTL' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '',  " & vbCrLf & _
                              "  	   PerformanceCls = '',  " & vbCrLf & _
                              "  	   PlanActual = 'ACTUAL',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0 ,PFC = 'PTL' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '4',  " & vbCrLf & _
                              "  	   PerformanceCls = 'OVER SUPPLY',  " & vbCrLf & _
                              "  	   PlanActual = 'PLAN',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0  ,PFC = 'OVS' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '',  " & vbCrLf & _
                              "  	   PerformanceCls = '',  " & vbCrLf & _
                              "  	   PlanActual = 'ACTUAL',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0  ,PFC = 'OVS' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '5',  " & vbCrLf & _
                              "  	   PerformanceCls = 'ADVANCE',  " & vbCrLf & _
                              "  	   PlanActual = 'PLAN',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0  ,PFC = 'ADV' " & vbCrLf & _
                              "  UNION ALL "

            ls_SQL = ls_SQL + "  SELECT No = '',  " & vbCrLf & _
                              "  	   PerformanceCls = '',  " & vbCrLf & _
                              "  	   PlanActual = 'ACTUAL',  " & vbCrLf & _
                              "  	   SupplierGroup ='Supplier001',  " & vbCrLf & _
                              "  	   PartNo = '7009-2191-02',  " & vbCrLf & _
                              "  	   ItemNo = '001',  " & vbCrLf & _
                              "  	   Month01 = 0,Month02 = 0,Month03 = 0,Month04 = 0,Month05 = 0,Month06 = 0,  " & vbCrLf & _
                              "  	   Month07 = 0,Month08 = 0,Month09 = 0,Month10 = 0,Month11 = 0,Month12 = 0,  " & vbCrLf & _
                              "  	   OntimeQty = 0,PlannedQty = 0,OntimeLine = 0,PlannedLine = 0,CummQty = 0,  " & vbCrLf & _
                              "  	   AQ = 0,AL = 0,CQ = 0   ,PFC = 'ADV' "




            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
            sqlConn.Close()
            up_GridHeader()

        End Using
    End Sub

    Private Sub up_GridLoadParam()
        Dim ls_sql As String = ""
        Dim ls_supp As String = Trim(cboSupplierGroup.Text)
        Dim ls_part As String = Trim(cboPartCode.Text)
        Dim ls_Aff As String = Trim(cboAffiliateCode.Text)
        Dim ls_month As String = Left(dtPeriod.Text, 3)
        Dim ls_year As Double = Right(dtPeriod.Text, 4)
        Dim ls_Month01 As String = ""
        Dim ls_Month02 As String = ""
        Dim ls_Month03 As String = ""
        Dim ls_Month04 As String = ""
        Dim ls_Month05 As String = ""
        Dim ls_Month06 As String = ""
        Dim ls_Month07 As String = ""
        Dim ls_Month08 As String = ""
        Dim ls_Month09 As String = ""
        Dim ls_Month10 As String = ""
        Dim ls_Month11 As String = ""
        Dim ls_Month12 As String = ""

        Select Case ls_month
            Case "Jan"
                ls_Month01 = ls_year & "-" & "01"
                ls_Month02 = ls_year & "-" & "02"
                ls_Month03 = ls_year & "-" & "03"
                ls_Month04 = ls_year & "-" & "04"
                ls_Month05 = ls_year & "-" & "05"
                ls_Month06 = ls_year & "-" & "06"
                ls_Month07 = ls_year & "-" & "07"
                ls_Month08 = ls_year & "-" & "08"
                ls_Month09 = ls_year & "-" & "09"
                ls_Month10 = ls_year & "-" & "10"
                ls_Month11 = ls_year & "-" & "11"
                ls_Month12 = ls_year & "-" & "12"
            Case "Feb"
                ls_Month01 = ls_year & "-" & "02"
                ls_Month02 = ls_year & "-" & "03"
                ls_Month03 = ls_year & "-" & "04"
                ls_Month04 = ls_year & "-" & "05"
                ls_Month05 = ls_year & "-" & "06"
                ls_Month06 = ls_year & "-" & "07"
                ls_Month07 = ls_year & "-" & "08"
                ls_Month08 = ls_year & "-" & "09"
                ls_Month09 = ls_year & "-" & "10"
                ls_Month10 = ls_year & "-" & "11"
                ls_Month11 = ls_year & "-" & "12"
                ls_Month12 = ls_year + 1 & "-" & "01"
            Case "Mar"
                ls_Month01 = ls_year & "-" & "03"
                ls_Month02 = ls_year & "-" & "04"
                ls_Month03 = ls_year & "-" & "05"
                ls_Month04 = ls_year & "-" & "06"
                ls_Month05 = ls_year & "-" & "07"
                ls_Month06 = ls_year & "-" & "08"
                ls_Month07 = ls_year & "-" & "09"
                ls_Month08 = ls_year & "-" & "10"
                ls_Month09 = ls_year & "-" & "11"
                ls_Month10 = ls_year & "-" & "12"
                ls_Month11 = ls_year + 1 & "-" & "01"
                ls_Month12 = ls_year + 1 & "-" & "02"
            Case "Apr"
                ls_Month01 = ls_year & "-" & "04"
                ls_Month02 = ls_year & "-" & "05"
                ls_Month03 = ls_year & "-" & "06"
                ls_Month04 = ls_year & "-" & "07"
                ls_Month05 = ls_year & "-" & "08"
                ls_Month06 = ls_year & "-" & "09"
                ls_Month07 = ls_year & "-" & "10"
                ls_Month08 = ls_year & "-" & "11"
                ls_Month09 = ls_year & "-" & "12"
                ls_Month10 = ls_year + 1 & "-" & "01"
                ls_Month11 = ls_year + 1 & "-" & "02"
                ls_Month12 = ls_year + 1 & "-" & "03"
            Case "May"
                ls_Month01 = ls_year & "-" & "05"
                ls_Month02 = ls_year & "-" & "06"
                ls_Month03 = ls_year & "-" & "07"
                ls_Month04 = ls_year & "-" & "08"
                ls_Month05 = ls_year & "-" & "09"
                ls_Month06 = ls_year & "-" & "10"
                ls_Month07 = ls_year & "-" & "11"
                ls_Month08 = ls_year & "-" & "12"
                ls_Month09 = ls_year + 1 & "-" & "01"
                ls_Month10 = ls_year + 1 & "-" & "02"
                ls_Month11 = ls_year + 1 & "-" & "03"
                ls_Month12 = ls_year + 1 & "-" & "04"
            Case "Jun"
                ls_Month01 = ls_year & "-" & "06"
                ls_Month02 = ls_year & "-" & "07"
                ls_Month03 = ls_year & "-" & "08"
                ls_Month04 = ls_year & "-" & "09"
                ls_Month05 = ls_year & "-" & "10"
                ls_Month06 = ls_year & "-" & "11"
                ls_Month07 = ls_year & "-" & "12"
                ls_Month08 = ls_year + 1 & "-" & "01"
                ls_Month09 = ls_year + 1 & "-" & "02"
                ls_Month10 = ls_year + 1 & "-" & "03"
                ls_Month11 = ls_year + 1 & "-" & "04"
                ls_Month12 = ls_year + 1 & "-" & "05"
            Case "Jul"
                ls_Month01 = ls_year & "-" & "07"
                ls_Month02 = ls_year & "-" & "08"
                ls_Month03 = ls_year & "-" & "09"
                ls_Month04 = ls_year & "-" & "10"
                ls_Month05 = ls_year & "-" & "11"
                ls_Month06 = ls_year & "-" & "12"
                ls_Month07 = ls_year + 1 & "-" & "01"
                ls_Month08 = ls_year + 1 & "-" & "02"
                ls_Month09 = ls_year + 1 & "-" & "03"
                ls_Month10 = ls_year + 1 & "-" & "04"
                ls_Month11 = ls_year + 1 & "-" & "05"
                ls_Month12 = ls_year + 1 & "-" & "06"
            Case "Aug"
                ls_Month01 = ls_year & "-" & "08"
                ls_Month02 = ls_year & "-" & "09"
                ls_Month03 = ls_year & "-" & "10"
                ls_Month04 = ls_year & "-" & "11"
                ls_Month05 = ls_year & "-" & "12"
                ls_Month06 = ls_year + 1 & "-" & "01"
                ls_Month07 = ls_year + 1 & "-" & "02"
                ls_Month08 = ls_year + 1 & "-" & "03"
                ls_Month09 = ls_year + 1 & "-" & "04"
                ls_Month10 = ls_year + 1 & "-" & "05"
                ls_Month11 = ls_year + 1 & "-" & "06"
                ls_Month12 = ls_year + 1 & "-" & "07"
            Case "Sep"
                ls_Month01 = ls_year & "-" & "09"
                ls_Month02 = ls_year & "-" & "10"
                ls_Month03 = ls_year & "-" & "11"
                ls_Month04 = ls_year & "-" & "12"
                ls_Month05 = ls_year + 1 & "-" & "01"
                ls_Month06 = ls_year + 1 & "-" & "02"
                ls_Month07 = ls_year + 1 & "-" & "03"
                ls_Month08 = ls_year + 1 & "-" & "04"
                ls_Month09 = ls_year + 1 & "-" & "05"
                ls_Month10 = ls_year + 1 & "-" & "06"
                ls_Month11 = ls_year + 1 & "-" & "07"
                ls_Month12 = ls_year + 1 & "-" & "08"
            Case "Oct"
                ls_Month01 = ls_year & "-" & "10"
                ls_Month02 = ls_year & "-" & "11"
                ls_Month03 = ls_year & "-" & "12"
                ls_Month04 = ls_year + 1 & "-" & "01"
                ls_Month05 = ls_year + 1 & "-" & "02"
                ls_Month06 = ls_year + 1 & "-" & "03"
                ls_Month07 = ls_year + 1 & "-" & "04"
                ls_Month08 = ls_year + 1 & "-" & "05"
                ls_Month09 = ls_year + 1 & "-" & "06"
                ls_Month10 = ls_year + 1 & "-" & "07"
                ls_Month11 = ls_year + 1 & "-" & "08"
                ls_Month12 = ls_year + 1 & "-" & "09"
            Case "Nov"
                ls_Month01 = ls_year & "-" & "11"
                ls_Month02 = ls_year & "-" & "12"
                ls_Month03 = ls_year + 1 & "-" & "01"
                ls_Month04 = ls_year + 1 & "-" & "02"
                ls_Month05 = ls_year + 1 & "-" & "03"
                ls_Month06 = ls_year + 1 & "-" & "04"
                ls_Month07 = ls_year + 1 & "-" & "05"
                ls_Month08 = ls_year + 1 & "-" & "06"
                ls_Month09 = ls_year + 1 & "-" & "07"
                ls_Month10 = ls_year + 1 & "-" & "08"
                ls_Month11 = ls_year + 1 & "-" & "09"
                ls_Month12 = ls_year + 1 & "-" & "10"
            Case "Dec"
                ls_Month01 = ls_year & "-" & "12"
                ls_Month02 = ls_year + 1 & "-" & "01"
                ls_Month03 = ls_year + 1 & "-" & "02"
                ls_Month04 = ls_year + 1 & "-" & "03"
                ls_Month05 = ls_year + 1 & "-" & "04"
                ls_Month06 = ls_year + 1 & "-" & "05"
                ls_Month07 = ls_year + 1 & "-" & "06"
                ls_Month08 = ls_year + 1 & "-" & "07"
                ls_Month09 = ls_year + 1 & "-" & "08"
                ls_Month10 = ls_year + 1 & "-" & "09"
                ls_Month11 = "-" & ls_year + 1 & "-" & "10"
                ls_Month12 = "-" & ls_year + 1 & "-" & "11"
        End Select

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_sql = uf_SQL(ls_supp, ls_part, ls_Aff,
                            ls_Month01,
                            ls_Month02,
                            ls_Month03,
                            ls_Month04,
                            ls_Month05,
                            ls_Month06,
                            ls_Month07,
                            ls_Month08,
                            ls_Month09,
                            ls_Month10,
                            ls_Month11,
                            ls_Month12
                            )

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
            sqlConn.Close()
            up_GridHeader()

        End Using
    End Sub

    Private Function uf_SQL(ByVal pSupplier As String, _
                                  pPart As String, _
                                  pAffiliate As String, _
                                  pMonth01 As String, _
                                  pMonth02 As String, _
                                  pMonth03 As String, _
                                  pMonth04 As String, _
                                  pMonth05 As String, _
                                  pMonth06 As String, _
                                  pMonth07 As String, _
                                  pMonth08 As String, _
                                  pMonth09 As String, _
                                  pMonth10 As String, _
                                  pMonth11 As String, _
                                  pMonth12 As String
                           )
        Dim ls_Sql As String = ""
        'Dim iMonth As Long
        Dim n As String
        ls_Sql = ls_Sql + " DROP TABLE TEMP#  " & vbCrLf & _
                          " select   *  INTO TEMP# from (  " & vbCrLf & _
                          " 	SELECT DISTINCT PerformanceCls = (SELECT CASE PM.DeliveryByPASICls WHEN 0 THEN RAM.PerformanceCls ELSE RPM.PerformanceCls END), " & vbCrLf & _
                          " 		   PlanActual = 'PLAN', " & vbCrLf & _
                          " 		   SG.SupplierGroupCode,  " & vbCrLf & _
                          " 		   MP.PartNo, " & vbCrLf & _
                          "            ItemNo = '', " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty) " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth01 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month01, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth02 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month02, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth03 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month03, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth04 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month04, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth05 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month05, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth06 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month06, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth07 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month07, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth08 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month08, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth09 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month09, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth10 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month10, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth11 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month11, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (KanbanQty)  " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM Kanban_Detail KD " & vbCrLf & _
                          " 			   LEFT JOIN Kanban_Master KM ON KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                          " 										 AND KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " 										 AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                          " 				   WHERE KD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND KD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND KD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 AND left(KanbanDate,7) = '" & pMonth12 & "' " & vbCrLf & _
                          " 					 GROUP BY KD.PartNo,KD.SupplierID,KD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month12, " & vbCrLf & _
                          " 		   CummQty = 0, " & vbCrLf

        ls_Sql = ls_Sql + " 		   POQty = 0 " & vbCrLf & _
                          "  " & vbCrLf & _
                          " 	FROM dbo.MS_SupplierGroup SG " & vbCrLf & _
                          " 	Left join dbo.MS_Supplier MS ON SG.SupplierGroupCode = MS.SupplierGroupCode " & vbCrLf & _
                          " 	left join dbo.MS_PartMapping MPM ON MS.SupplierID = MPM.SupplierID " & vbCrLf & _
                          " 	INNER join dbo.MS_Parts MP ON MPM.PartNo = MP.PartNo " & vbCrLf

        ls_Sql = ls_Sql + "     LEFT JOIN PO_Detail PD ON PD.PartNo = MP.PartNo  " & vbCrLf & _
                          "  						  AND PD.SupplierID = MS.SupplierID  " & vbCrLf & _
                          "  						  AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          "  	LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo  " & vbCrLf & _
                          "  							AND PD.AffiliateID = PM.AffiliateID  " & vbCrLf & _
                          "  							AND PD.SupplierID = PM.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo  " & vbCrLf & _
                          "  											AND PD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
                          "  											AND PD.SupplierID = RAD.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo  " & vbCrLf & _
                          "  											AND RAD.AffiliateID = RAM.AffiliateID  "

        ls_Sql = ls_Sql + "  											AND RAD.SupplierID = RAM.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo  " & vbCrLf & _
                          "  									AND PD.AffiliateID = RPD.AffiliateID  " & vbCrLf & _
                          "  									AND PD.SupplierID = RPD.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo  " & vbCrLf & _
                          "  									AND RPD.SupplierID = RPM.SupplierID  " & vbCrLf & _
                          "  									AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          "  " & vbCrLf & _
                          " 	WHERE SG.SupplierGroupCode = '" & pSupplier & "' " & vbCrLf & _
                          " 	  AND MP.PartNo = '" & pPart & "' " & vbCrLf & _
                          " 	  AND MPM.AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                          "  " & vbCrLf

        ls_Sql = ls_Sql + " 	UNION ALL " & vbCrLf & _
                          "  " & vbCrLf & _
                          " 	SELECT DISTINCT PerformanceCls = (SELECT CASE PM.DeliveryByPASICls WHEN 0 THEN RAM.PerformanceCls ELSE RPM.PerformanceCls END), " & vbCrLf & _
                          " 		   PlanActual = 'ACTUAL', " & vbCrLf & _
                          " 		   SG.SupplierGroupCode,  " & vbCrLf & _
                          " 		   MP.PartNo, " & vbCrLf & _
                          "            ItemNo = '', " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf

        ls_Sql = ls_Sql + " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth01 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf

        ls_Sql = ls_Sql + " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth01 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month01, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf

        ls_Sql = ls_Sql + " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth02 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf

        ls_Sql = ls_Sql + " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth02 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month02, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf

        ls_Sql = ls_Sql + " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth03 & "' "

        ls_Sql = ls_Sql + " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth03 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf

        ls_Sql = ls_Sql + " 		   ),0) AS Month03, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf

        ls_Sql = ls_Sql + " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth04 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth04 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf

        ls_Sql = ls_Sql + " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month04, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf

        ls_Sql = ls_Sql + " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth05 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth05 & "' " & vbCrLf

        ls_Sql = ls_Sql + " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month05, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf

        ls_Sql = ls_Sql + " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth06 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf

        ls_Sql = ls_Sql + " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth06 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month06, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf

        ls_Sql = ls_Sql + " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth07 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf

        ls_Sql = ls_Sql + " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth07 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month07, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf

        ls_Sql = ls_Sql + " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth08 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf

        ls_Sql = ls_Sql + " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth08 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month08, " & vbCrLf

        ls_Sql = ls_Sql + " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf

        ls_Sql = ls_Sql + " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth09 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth09 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf

        ls_Sql = ls_Sql + " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month09, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf

        ls_Sql = ls_Sql + " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth10 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth10 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf

        ls_Sql = ls_Sql + " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month10, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf

        ls_Sql = ls_Sql + " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth11 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf & _
                          " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf

        ls_Sql = ls_Sql + " 													AND left(RPM.ReceiveDate,7) = '" & pMonth11 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month11, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (CASE PM.DeliveryByPASICls WHEN 0 THEN RAD.RecQty ELSE RPD.GoodRecQty END) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf & _
                          " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf

        ls_Sql = ls_Sql + " 					LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo " & vbCrLf & _
                          " 														 AND PD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 														 AND PD.SupplierID = RAD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 														 AND RAD.AffiliateID = RAM.AffiliateID " & vbCrLf & _
                          " 														 AND RAD.SupplierID = RAM.SupplierID " & vbCrLf & _
                          " 														 AND left(RAM.ReceiveDate,7) = '" & pMonth12 & "' " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo " & vbCrLf & _
                          " 													AND PD.AffiliateID = RPD.AffiliateID " & vbCrLf & _
                          " 													AND PD.SupplierID = RPD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo " & vbCrLf

        ls_Sql = ls_Sql + " 													AND RPD.SupplierID = RPM.SupplierID " & vbCrLf & _
                          " 													AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 													AND left(RPM.ReceiveDate,7) = '" & pMonth12 & "' " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS Month12, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (DSD.DOQty) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 					LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo " & vbCrLf

        ls_Sql = ls_Sql + " 										  AND PD.AffiliateID = PM.AffiliateID " & vbCrLf & _
                          " 										  AND PD.SupplierID = PM.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN DOSupplier_Detail DSD ON DSD.PONo = PD.PONo " & vbCrLf & _
                          " 												   AND DSD.AffiliateID = PD.AffiliateID " & vbCrLf & _
                          " 												   AND DSD.SupplierID = PD.SupplierID " & vbCrLf & _
                          " 					LEFT JOIN DOSupplier_Master DSM ON DSM.SuratJalanNo = DSD.SuratJalanNo " & vbCrLf & _
                          " 												   AND DSM.SupplierID = DSD.SupplierID " & vbCrLf & _
                          " 												   AND DSM.AffiliateID = DSD.AffiliateID " & vbCrLf & _
                          " 												   AND DSM.DeliveryDate = PM.Period " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf

        ls_Sql = ls_Sql + " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS CummQty, " & vbCrLf & _
                          " 		   ISNULL((SELECT SUM (PD.POQty) " & vbCrLf & _
                          " 					FROM PO_Detail PD " & vbCrLf & _
                          " 				   WHERE PD.PartNo = MP.PartNo " & vbCrLf & _
                          " 					 AND PD.SupplierID = MS.SupplierID " & vbCrLf & _
                          " 					 AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          " 					 GROUP BY PD.PartNo,PD.SupplierID,PD.AffiliateID " & vbCrLf & _
                          " 		   ),0) AS POQty " & vbCrLf & _
                          "  " & vbCrLf

        ls_Sql = ls_Sql + " 	FROM dbo.MS_SupplierGroup SG " & vbCrLf & _
                          " 	Left join dbo.MS_Supplier MS ON SG.SupplierGroupCode = MS.SupplierGroupCode " & vbCrLf & _
                          " 	left join dbo.MS_PartMapping MPM ON MS.SupplierID = MPM.SupplierID " & vbCrLf & _
                          " 	INNER join dbo.MS_Parts MP ON MPM.PartNo = MP.PartNo " & vbCrLf

        ls_Sql = ls_Sql + "     LEFT JOIN PO_Detail PD ON PD.PartNo = MP.PartNo  " & vbCrLf & _
                          "  						  AND PD.SupplierID = MS.SupplierID  " & vbCrLf & _
                          "  						  AND PD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                          "  	LEFT JOIN PO_Master PM ON PD.PONo = PM.PONo  " & vbCrLf & _
                          "  							AND PD.AffiliateID = PM.AffiliateID  " & vbCrLf & _
                          "  							AND PD.SupplierID = PM.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceiveAffiliate_Detail RAD ON PD.PONo = RAD.PONo  " & vbCrLf & _
                          "  											AND PD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
                          "  											AND PD.SupplierID = RAD.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceiveAffiliate_Master RAM ON RAD.SuratJalanNo = RAM.SuratJalanNo  " & vbCrLf & _
                          "  											AND RAD.AffiliateID = RAM.AffiliateID  "

        ls_Sql = ls_Sql + "  											AND RAD.SupplierID = RAM.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceivePASI_Detail RPD ON PD.PONo = RPD.PONo  " & vbCrLf & _
                          "  									AND PD.AffiliateID = RPD.AffiliateID  " & vbCrLf & _
                          "  									AND PD.SupplierID = RPD.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceivePASI_Master RPM ON RPD.SuratJalanNo = RPM.SuratJalanNo  " & vbCrLf & _
                          "  									AND RPD.SupplierID = RPM.SupplierID  " & vbCrLf & _
                          "  									AND RPD.AffiliateID = RPM.AffiliateID " & vbCrLf & _
                          " 	WHERE SG.SupplierGroupCode = '" & pSupplier & "' " & vbCrLf & _
                          " 	  AND MP.PartNo = '" & pPart & "' " & vbCrLf & _
                          " 	  AND MPM.AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                          " )A " & vbCrLf & _
                          "" & vbCrLf

        ls_Sql = ls_Sql + " select *, " & vbCrLf & _
                          " 	   AQ = case when PlannedQty = 0 then 0 else (OnTimeQty / PlannedQty) * 100 end, " & vbCrLf & _
                          " 	   AL = case when PlannedLine = 0 then 0 else (OnTimeLine / PlannedLine) * 100 end, " & vbCrLf & _
                          " 	   CQ = case when POQtySub = 0 then 0 else (CummQtySub / POQtySub) * 100 end " & vbCrLf & _
                          " from ( " & vbCrLf & _
                          " 	select  No = '',PerformanceCls, PlanActual, SupplierGroupCode, PartNo, ItemNo,  " & vbCrLf & _
                          " 			Month01,Month02,Month03,Month04,Month05,Month06,Month07,Month08,Month09,Month10,Month11,Month12, " & vbCrLf & _
                          " 			OntimeQty = OTQM01+OTQM02+OTQM03+OTQM04+OTQM05+OTQM06+OTQM07+OTQM08+OTQM09+OTQM10+OTQM11+OTQM12, " & vbCrLf & _
                          " 			PlannedQty = PLQM01+PLQM02+PLQM03+PLQM04+PLQM05+PLQM06+PLQM07+PLQM08+PLQM09+PLQM10+PLQM11+PLQM12, " & vbCrLf & _
                          " 			OntimeLine = OTLM01+OTLM02+OTLM03+OTLM04+OTLM05+OTLM06+OTLM07+OTLM08+OTLM09+OTLM10+OTLM11+OTLM12, " & vbCrLf & _
                          " 			PlannedLine = PLLM01+PLLM02+PLLM03+PLLM04+PLLM05+PLLM06+PLLM07+PLLM08+PLLM09+PLLM10+PLLM11+PLLM12, " & vbCrLf

        ls_Sql = ls_Sql + " 			CummQtySub = CummQtySub, " & vbCrLf & _
                          " 			POQtySub = POQtySub " & vbCrLf & _
                          " 	from ( " & vbCrLf & _
                          " 		SELECT     A.*,		   " & vbCrLf & _
                          " 				  case when a.Month01 > b.Month01 and A.PlanActual='PLAN' then a.Month01- (a.Month01- b.Month01) +(a.Month01- b.Month01) -(a.Month01- b.Month01)   " & vbCrLf & _
                          " 						when B.Month01 > A.Month01 and A.PlanActual='PLAN' THEN a.Month01 ELSE 0 end OTQM01, " & vbCrLf & _
                          " 				  case when a.Month02 > b.Month01 and A.PlanActual='PLAN' then a.Month02- (a.Month02- b.Month02) +(a.Month02- b.Month02) -(a.Month02- b.Month02)   " & vbCrLf & _
                          " 						when B.Month02 > A.Month01 and A.PlanActual='PLAN' THEN a.Month02 ELSE 0 end OTQM02, " & vbCrLf & _
                          " 				  case when a.Month03 > b.Month03 and A.PlanActual='PLAN' then a.Month03- (a.Month03- b.Month03) +(a.Month03- b.Month03) -(a.Month03- b.Month03)   " & vbCrLf & _
                          " 						when B.Month03 > A.Month03 and A.PlanActual='PLAN' THEN a.Month03 ELSE 0 end OTQM03, " & vbCrLf & _
                          " 				  case when a.Month04 > b.Month04 and A.PlanActual='PLAN' then a.Month04- (a.Month04- b.Month04) +(a.Month04- b.Month04) -(a.Month04- b.Month04)   " & vbCrLf

        ls_Sql = ls_Sql + " 						when B.Month04 > A.Month04 and A.PlanActual='PLAN' THEN a.Month04 ELSE 0 end OTQM04, " & vbCrLf & _
                          " 				  case when a.Month05 > b.Month05 and A.PlanActual='PLAN' then a.Month05- (a.Month05- b.Month05) +(a.Month05- b.Month05) -(a.Month05- b.Month05)   " & vbCrLf & _
                          " 						when B.Month05 > A.Month05 and A.PlanActual='PLAN' THEN a.Month05 ELSE 0 end OTQM05, " & vbCrLf & _
                          " 				  case when a.Month06 > b.Month06 and A.PlanActual='PLAN' then a.Month06- (a.Month06- b.Month06) +(a.Month06- b.Month06) -(a.Month06- b.Month06)   " & vbCrLf & _
                          " 						when B.Month06 > A.Month06 and A.PlanActual='PLAN' THEN a.Month06 ELSE 0 end OTQM06, " & vbCrLf & _
                          " 				  case when a.Month07 > b.Month07 and A.PlanActual='PLAN' then a.Month07- (a.Month07- b.Month07) +(a.Month07- b.Month07) -(a.Month07- b.Month07)   " & vbCrLf & _
                          " 						when B.Month07 > A.Month07 and A.PlanActual='PLAN' THEN a.Month07 ELSE 0 end OTQM07, " & vbCrLf & _
                          " 				  case when a.Month08 > b.Month08 and A.PlanActual='PLAN' then a.Month08- (a.Month08- b.Month08) +(a.Month08- b.Month08) -(a.Month08- b.Month08)   " & vbCrLf & _
                          " 						when B.Month08 > A.Month08 and A.PlanActual='PLAN' THEN a.Month08 ELSE 0 end OTQM08, " & vbCrLf & _
                          " 				  case when a.Month09 > b.Month09 and A.PlanActual='PLAN' then a.Month09- (a.Month09- b.Month09) +(a.Month09- b.Month09) -(a.Month09- b.Month09)   " & vbCrLf & _
                          " 						when B.Month09 > A.Month09 and A.PlanActual='PLAN' THEN a.Month09 ELSE 0 end OTQM09, " & vbCrLf

        ls_Sql = ls_Sql + " 				  case when a.Month10 > b.Month10 and A.PlanActual='PLAN' then a.Month10- (a.Month10- b.Month10) +(a.Month10- b.Month10) -(a.Month10- b.Month10)   " & vbCrLf & _
                          " 						when B.Month10 > A.Month10 and A.PlanActual='PLAN' THEN a.Month10 ELSE 0 end OTQM10, " & vbCrLf & _
                          " 				  case when a.Month11 > b.Month11 and A.PlanActual='PLAN' then a.Month11- (a.Month11- b.Month11) +(a.Month11- b.Month11) -(a.Month11- b.Month11)   " & vbCrLf & _
                          " 						when B.Month11 > A.Month11 and A.PlanActual='PLAN' THEN a.Month11 ELSE 0 end OTQM11, " & vbCrLf & _
                          " 				  case when a.Month12 > b.Month12 and A.PlanActual='PLAN' then a.Month12- (a.Month12- b.Month12) +(a.Month12- b.Month12) -(a.Month12- b.Month12)   " & vbCrLf & _
                          " 						when B.Month12 > A.Month12 and A.PlanActual='PLAN' THEN a.Month12 ELSE 0 end OTQM12, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month01 else 0 end PLQM01, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month02 else 0 end PLQM02, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month03 else 0 end PLQM03, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month04 else 0 end PLQM04, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month05 else 0 end PLQM05, " & vbCrLf

        ls_Sql = ls_Sql + " 				  case when A.PlanActual='PLAN' then a.Month06 else 0 end PLQM06, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month07 else 0 end PLQM07, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month08 else 0 end PLQM08, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month09 else 0 end PLQM09, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month10 else 0 end PLQM10, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month11 else 0 end PLQM11, " & vbCrLf & _
                          " 				  case when A.PlanActual='PLAN' then a.Month12 else 0 end PLQM12, " & vbCrLf & _
                          " 				  case when a.Month01 <> b.Month01 and b.Month01 <> 0 AND A.Month01 <> 0 then 1 else 0 end OTLM01, " & vbCrLf & _
                          " 				  case when a.Month02 <> b.Month02 and b.Month02 <> 0 AND A.Month02 <> 0 then 1 else 0 end OTLM02, " & vbCrLf & _
                          " 				  case when a.Month03 <> b.Month03 and b.Month03 <> 0 and a.Month03 <> 0 then 1 else 0 end OTLM03, " & vbCrLf & _
                          " 				  case when a.Month04 <> b.Month04 and b.Month04 <> 0 and a.Month04 <> 0 then 1 else 0 end OTLM04, " & vbCrLf

        ls_Sql = ls_Sql + " 				  case when a.Month05 <> b.Month05 and b.Month05 <> 0 and a.Month05 <> 0 then 1 else 0 end OTLM05, " & vbCrLf & _
                          " 				  case when a.Month06 <> b.Month06 and b.Month06 <> 0 and a.Month06 <> 0 then 1 else 0 end OTLM06, " & vbCrLf & _
                          " 				  case when a.Month07 <> b.Month07 and b.Month07 <> 0 and a.Month07 <> 0 then 1 else 0 end OTLM07, " & vbCrLf & _
                          " 				  case when a.Month08 <> b.Month08 and b.Month08 <> 0 and a.Month08 <> 0 then 1 else 0 end OTLM08, " & vbCrLf & _
                          " 				  case when a.Month09 <> b.Month09 and b.Month09 <> 0 and a.Month09 <> 0 then 1 else 0 end OTLM09, " & vbCrLf & _
                          " 				  case when a.Month10 <> b.Month10 and b.Month10 <> 0 and a.Month10 <> 0 then 1 else 0 end OTLM10, " & vbCrLf & _
                          " 				  case when a.Month11 <> b.Month11 and b.Month11 <> 0 and a.Month11 <> 0 then 1 else 0 end OTLM11, " & vbCrLf & _
                          " 				  case when a.Month12 <> b.Month12 and b.Month12 <> 0 and a.Month12 <> 0 then 1 else 0 end OTLM12, " & vbCrLf & _
                          " 				  case when a.Month01 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM01, " & vbCrLf & _
                          " 				  case when a.Month02 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM02, " & vbCrLf & _
                          " 				  case when a.Month03 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM03, " & vbCrLf

        ls_Sql = ls_Sql + " 				  case when a.Month04 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM04, " & vbCrLf & _
                          " 				  case when a.Month05 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM05, " & vbCrLf & _
                          " 				  case when a.Month06 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM06, " & vbCrLf & _
                          " 				  case when a.Month07 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM07, " & vbCrLf & _
                          " 				  case when a.Month08 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM08, " & vbCrLf & _
                          " 				  case when a.Month09 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM09, " & vbCrLf & _
                          " 				  case when a.Month10 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM10, " & vbCrLf & _
                          " 				  case when a.Month11 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM11, " & vbCrLf & _
                          " 				  case when a.Month12 > 0 and A.PlanActual='PLAN' then 1 else 0 end PLLM12, " & vbCrLf & _
                          "  " & vbCrLf & _
                          " 				  case when A.PlanActual = 'PLAN' then b.CummQty else 0 end CummQtySub, " & vbCrLf

        ls_Sql = ls_Sql + " 				  case when A.PlanActual = 'PLAN' then b.POQty else 0 end POQtySub " & vbCrLf & _
                          " 		 FROM TEMP# A LEFT JOIN  " & vbCrLf & _
                          " 		( " & vbCrLf & _
                          " 			SELECT * FROM TEMP#  " & vbCrLf & _
                          " 			WHERE PlanActual='Actual' " & vbCrLf & _
                          " 		)B on A.SupplierGroupCode=B.SupplierGroupCode AND A.PartNo=B.PartNo AND A.PlanActual='PLAN' " & vbCrLf & _
                          " 	)SUB1 " & vbCrLf & _
                          " )SUB2 " & vbCrLf


        uf_SQL = ls_Sql
    End Function

    Private Sub up_GridEmpty()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ls_SQL + " SELECT TOP 0 No = '', " & vbCrLf & _
                              " 	   PerformanceCls = '', " & vbCrLf & _
                              " 	   PlanActual = '', " & vbCrLf & _
                              " 	   SupplierGroupCode ='', " & vbCrLf & _
                              " 	   PartNo = '', " & vbCrLf & _
                              " 	   ItemNo = '', " & vbCrLf & _
                              " 	   Month01 = '', " & vbCrLf & _
                              " 	   Month02 = '', " & vbCrLf & _
                              " 	   Month03 = '', " & vbCrLf & _
                              " 	   Month04 = '', " & vbCrLf & _
                              " 	   Month05 = '', "

            ls_SQL = ls_SQL + " 	   Month06 = '', " & vbCrLf & _
                              " 	   Month07 = '', " & vbCrLf & _
                              " 	   Month08 = '', " & vbCrLf & _
                              " 	   Month09 = '', " & vbCrLf & _
                              " 	   Month10 = '', " & vbCrLf & _
                              " 	   Month11 = '', " & vbCrLf & _
                              " 	   Month12 = '', " & vbCrLf & _
                              " 	   OntimeQty = '', " & vbCrLf & _
                              " 	   PlannedQty = '', " & vbCrLf & _
                              " 	   OntimeLine = '', " & vbCrLf & _
                              " 	   PlannedLine = '', "

            ls_SQL = ls_SQL + " 	   CummQtySub = '', " & vbCrLf & _
                              " 	   AQ = '', " & vbCrLf & _
                              " 	   AL = '', " & vbCrLf & _
                              " 	   CQ = '' " & vbCrLf & _
                              " "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
            sqlConn.Close()

            up_GridHeader()
        End Using
    End Sub

    Private Sub up_GridHeader()
        Dim ls_month As String = Left(dtPeriod.Text, 3)
        Dim ls_year As Double = Right(dtPeriod.Text, 4)

        Select Case ls_month
            Case "Jan"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(2) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(3) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(4) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(5) & "-" & ls_year
                grid.VisibleColumns(12).Caption = uf_GetMonth(6) & "-" & ls_year
                grid.VisibleColumns(13).Caption = uf_GetMonth(7) & "-" & ls_year
                grid.VisibleColumns(14).Caption = uf_GetMonth(8) & "-" & ls_year
                grid.VisibleColumns(15).Caption = uf_GetMonth(9) & "-" & ls_year
                grid.VisibleColumns(16).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(17).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(18).Caption = uf_GetMonth(12) & "-" & ls_year
            Case "Feb"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(3) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(4) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(5) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(6) & "-" & ls_year
                grid.VisibleColumns(12).Caption = uf_GetMonth(7) & "-" & ls_year
                grid.VisibleColumns(13).Caption = uf_GetMonth(8) & "-" & ls_year
                grid.VisibleColumns(14).Caption = uf_GetMonth(9) & "-" & ls_year
                grid.VisibleColumns(15).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(16).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(17).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(18).Caption = uf_GetMonth(1) & "-" & ls_year + 1
            Case "Mar"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(4) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(5) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(6) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(7) & "-" & ls_year
                grid.VisibleColumns(12).Caption = uf_GetMonth(8) & "-" & ls_year
                grid.VisibleColumns(13).Caption = uf_GetMonth(9) & "-" & ls_year
                grid.VisibleColumns(14).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(15).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(16).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(17).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(2) & "-" & ls_year + 1
            Case "Apr"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(5) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(6) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(7) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(8) & "-" & ls_year
                grid.VisibleColumns(12).Caption = uf_GetMonth(9) & "-" & ls_year
                grid.VisibleColumns(13).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(14).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(15).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(16).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(3) & "-" & ls_year + 1
            Case "May"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(6) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(7) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(8) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(9) & "-" & ls_year
                grid.VisibleColumns(12).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(13).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(14).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(15).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(16).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(3) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(4) & "-" & ls_year + 1
            Case "Jun"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(7) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(8) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(9) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(12).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(13).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(14).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(15).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(16).Caption = uf_GetMonth(3) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(4) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(5) & "-" & ls_year + 1
            Case "Jul"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(8) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(9) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(12).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(13).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(14).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(15).Caption = uf_GetMonth(3) & "-" & ls_year + 1
                grid.VisibleColumns(16).Caption = uf_GetMonth(4) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(5) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(6) & "-" & ls_year + 1
            Case "Aug"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(9) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(12).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(13).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(14).Caption = uf_GetMonth(3) & "-" & ls_year + 1
                grid.VisibleColumns(15).Caption = uf_GetMonth(4) & "-" & ls_year + 1
                grid.VisibleColumns(16).Caption = uf_GetMonth(5) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(6) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(7) & "-" & ls_year + 1
            Case "Sep"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(10) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(11).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(12).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(13).Caption = uf_GetMonth(3) & "-" & ls_year + 1
                grid.VisibleColumns(14).Caption = uf_GetMonth(4) & "-" & ls_year + 1
                grid.VisibleColumns(15).Caption = uf_GetMonth(5) & "-" & ls_year + 1
                grid.VisibleColumns(16).Caption = uf_GetMonth(6) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(7) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(8) & "-" & ls_year + 1
            Case "Oct"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(11) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(10).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(11).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(12).Caption = uf_GetMonth(3) & "-" & ls_year + 1
                grid.VisibleColumns(13).Caption = uf_GetMonth(4) & "-" & ls_year + 1
                grid.VisibleColumns(14).Caption = uf_GetMonth(5) & "-" & ls_year + 1
                grid.VisibleColumns(15).Caption = uf_GetMonth(6) & "-" & ls_year + 1
                grid.VisibleColumns(16).Caption = uf_GetMonth(7) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(8) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(9) & "-" & ls_year + 1
            Case "Nov"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(12) & "-" & ls_year
                grid.VisibleColumns(9).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(10).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(11).Caption = uf_GetMonth(3) & "-" & ls_year + 1
                grid.VisibleColumns(12).Caption = uf_GetMonth(4) & "-" & ls_year + 1
                grid.VisibleColumns(13).Caption = uf_GetMonth(5) & "-" & ls_year + 1
                grid.VisibleColumns(14).Caption = uf_GetMonth(6) & "-" & ls_year + 1
                grid.VisibleColumns(15).Caption = uf_GetMonth(7) & "-" & ls_year + 1
                grid.VisibleColumns(16).Caption = uf_GetMonth(8) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(9) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(10) & "-" & ls_year + 1
            Case "Dec"
                grid.VisibleColumns(7).Caption = ls_month & "-" & ls_year
                grid.VisibleColumns(8).Caption = uf_GetMonth(1) & "-" & ls_year + 1
                grid.VisibleColumns(9).Caption = uf_GetMonth(2) & "-" & ls_year + 1
                grid.VisibleColumns(10).Caption = uf_GetMonth(3) & "-" & ls_year + 1
                grid.VisibleColumns(11).Caption = uf_GetMonth(4) & "-" & ls_year + 1
                grid.VisibleColumns(12).Caption = uf_GetMonth(5) & "-" & ls_year + 1
                grid.VisibleColumns(13).Caption = uf_GetMonth(6) & "-" & ls_year + 1
                grid.VisibleColumns(14).Caption = uf_GetMonth(7) & "-" & ls_year + 1
                grid.VisibleColumns(15).Caption = uf_GetMonth(8) & "-" & ls_year + 1
                grid.VisibleColumns(16).Caption = uf_GetMonth(9) & "-" & ls_year + 1
                grid.VisibleColumns(17).Caption = uf_GetMonth(10) & "-" & ls_year + 1
                grid.VisibleColumns(18).Caption = uf_GetMonth(11) & "-" & ls_year + 1
        End Select
        
    End Sub

    Private Function uf_GetMonth(ByVal pMonth As String)
        Dim ls_month As String = ""
        Select Case pMonth
            Case "1"
                ls_month = "Jan"
            Case "2"
                ls_month = "Feb"
            Case "3"
                ls_month = "Mar"
            Case "4"
                ls_month = "Apr"
            Case "5"
                ls_month = "May"
            Case "6"
                ls_month = "Jun"
            Case "7"
                ls_month = "Jul"
            Case "8"
                ls_month = "Aug"
            Case "9"
                ls_month = "Sep"
            Case "10"
                ls_month = "Oct"
            Case "11"
                ls_month = "Nov"
            Case "12"
                ls_month = "Dec"
        End Select
        uf_GetMonth = ls_month
    End Function

    Private Sub Excel()
        Dim strFileSize As String = ""

        Dim fi As New FileInfo(Server.MapPath("~\Delivery\DELIVERY PERFORM REPORT.xlsx"))
        If fi.Exists Then
            fi.Delete()
            fi = New FileInfo(Server.MapPath("~\Delivery\DELIVERY PERFORM REPORT.xlsx"))
        End If
        Dim exl As New ExcelPackage(fi)
        Dim ws As ExcelWorksheet
        Dim space As Integer = 8
        ws = exl.Workbook.Worksheets.Add("Sheet1")
        ws.Cells(1, 1, 100, 100).Style.Font.Name = "Calibri"
        ws.Cells(1, 1, 100, 100).Style.Font.Size = 9

        With ws
            .Cells(1, 1).Value = "DELIVERY PERFORMANCE REPORT"
            .Cells(3, 1).Value = "PERIOD"
            .Cells(4, 1).Value = "SUPPLIER GROUP"
            .Cells(5, 1).Value = "PART CODE / NAME"

            .Cells(3, 3).Value = ": " & dtPeriod.Text
            .Cells(4, 3).Value = ": " & cboSupplierGroup.Text
            .Cells(5, 3).Value = ": " & cboPartCode.Text

            If Grid.VisibleRowCount > 0 Then
                .Cells(space, 1).Value = "NO."
                .Cells(space, 2).Value = "PERFORMANCE CLS"
                .Cells(space, 3).Value = "PLAN / ACTUAL"
                .Cells(space, 4).Value = "SUPPLIER GROUP"
                .Cells(space, 5).Value = "PART NO"
                .Cells(space, 6).Value = "ITEM NO"
                .Cells(space, 7).Value = grid.VisibleColumns(7).Caption
                .Cells(space, 8).Value = grid.VisibleColumns(8).Caption
                .Cells(space, 9).Value = grid.VisibleColumns(9).Caption
                .Cells(space, 10).Value = grid.VisibleColumns(10).Caption
                .Cells(space, 11).Value = grid.VisibleColumns(11).Caption
                .Cells(space, 12).Value = grid.VisibleColumns(12).Caption
                .Cells(space, 13).Value = grid.VisibleColumns(13).Caption
                .Cells(space, 14).Value = grid.VisibleColumns(14).Caption
                .Cells(space, 15).Value = grid.VisibleColumns(15).Caption
                .Cells(space, 16).Value = grid.VisibleColumns(16).Caption
                .Cells(space, 17).Value = grid.VisibleColumns(17).Caption
                .Cells(space, 18).Value = grid.VisibleColumns(18).Caption
                .Cells(space, 19).Value = "ON TIME QTY"
                .Cells(space, 20).Value = "PLANNED QTY"
                .Cells(space, 21).Value = "ON TIME LINE"
                .Cells(space, 22).Value = "PLANNED LINE"
                .Cells(space, 23).Value = "CUMM QTY"


                For i = 1 To Grid.VisibleRowCount - 1
                    .Cells(i + space, 1).Value = Trim(grid.GetRowValues(i, "No"))
                    .Cells(i + space, 2).Value = Trim(grid.GetRowValues(i, "PerformanceCls"))
                    .Cells(i + space, 3).Value = Trim(grid.GetRowValues(i, "PlanActual"))
                    .Cells(i + space, 4).Value = Trim(grid.GetRowValues(i, "SupplierGroup"))
                    .Cells(i + space, 5).Value = Trim(grid.GetRowValues(i, "PartNo"))
                    .Cells(i + space, 6).Value = Trim(grid.GetRowValues(i, "ItemNo"))
                    .Cells(i + space, 7).Value = Trim(grid.GetRowValues(i, "Month01"))
                    .Cells(i + space, 8).Value = Trim(grid.GetRowValues(i, "Month02"))
                    .Cells(i + space, 9).Value = Trim(grid.GetRowValues(i, "Month03"))
                    .Cells(i + space, 10).Value = Trim(grid.GetRowValues(i, "Month04"))
                    .Cells(i + space, 11).Value = Trim(grid.GetRowValues(i, "Month05"))
                    .Cells(i + space, 12).Value = Trim(grid.GetRowValues(i, "Month06"))
                    .Cells(i + space, 13).Value = Trim(grid.GetRowValues(i, "Month07"))
                    .Cells(i + space, 14).Value = Trim(grid.GetRowValues(i, "Month08"))
                    .Cells(i + space, 15).Value = Trim(grid.GetRowValues(i, "Month09"))
                    .Cells(i + space, 16).Value = Trim(grid.GetRowValues(i, "Month10"))
                    .Cells(i + space, 17).Value = Trim(grid.GetRowValues(i, "Month11"))
                    .Cells(i + space, 18).Value = Trim(grid.GetRowValues(i, "Month12"))
                    .Cells(i + space, 19).Value = Trim(grid.GetRowValues(i, "OntimeQty"))
                    .Cells(i + space, 20).Value = Trim(grid.GetRowValues(i, "PlannedQty"))
                    .Cells(i + space, 21).Value = Trim(grid.GetRowValues(i, "OntimeLine"))
                    .Cells(i + space, 22).Value = Trim(grid.GetRowValues(i, "PlannedLine"))
                    .Cells(i + space, 23).Value = Trim(grid.GetRowValues(i, "CummQty"))

                Next

                Dim rgAll As ExcelRange = ws.Cells(8, 1, grid.VisibleRowCount + 7, 23)
                EpPlusDrawAllBorders(rgAll)

                'save to file
                exl.Save()
            End If
            'redirect to file download
            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)
        End With

        Exit Sub
ErrHandler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_fillcombo()
                lblerrmessage.Text = ""
                'grid.JSProperties("cpdate") = Format(Now, "dd MMM yyyy")

            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())

        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 5, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 5, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)

            Select Case pAction
                Case "gridload"
                    'Call up_GridLoad()
                    Call up_GridLoadParam()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "kosong"
                    Call up_GridEmpty()
            End Select

EndProcedure:
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub btnSubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub btnDownload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDownload.Click
        'Call up_GridLoad()

        'With gridExport
        '    .FileName = Now.ToString("ddMMyyyy") & "_DeliveryPerformanceReport"
        '    .WriteXlsxToResponse()
        'End With
        Call Excel()
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If e.DataColumn.FieldName = "OntimeQty" Or
           e.DataColumn.FieldName = "PlannedQty" Or
           e.DataColumn.FieldName = "OntimeLine" Or
           e.DataColumn.FieldName = "PlannedLine" Or
           e.DataColumn.FieldName = "CummQtySub" Or
           e.DataColumn.FieldName = "AQ" Or
           e.DataColumn.FieldName = "AL" Or
           e.DataColumn.FieldName = "CQ" Then
            If (e.GetValue("PlanActual") = "ACTUAL") Then
                e.Cell.Text = ""
            End If
        End If
    End Sub
End Class