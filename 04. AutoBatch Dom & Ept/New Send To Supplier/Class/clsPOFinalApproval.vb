Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsPOFinalApproval
    Shared Sub up_FinalApprovePODomestic(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")
        Dim ls_sql As String
        Dim i As Integer, j As Integer, k As Integer, n As Integer
        Dim m As Long
        Dim ls_Time As String = "00:00:00"
        Dim jlhHari As Integer = 0
        Dim ls_SeqNo As Integer = 0

        Dim x As Integer

        Dim pPONo As String = ""
        Dim pSupplier As String = ""
        Dim pAffiliate As String = ""
        Dim pAppPeriod As Date
        Dim ls_Prefix As String = ""

        Dim ds As New DataSet
        Dim dsRemaining As New DataSet
        Dim dsMoqQty As New DataSet
        Dim POMoq = 0, POQty As Integer = 0

        Try
            log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Get data PO")

            '01. Get Data PO Final Approval
            ls_sql = " select a.*, RTRIM(ISNULL(b.LabelCode,'C')) LabelPrefix from PO_Master a left join MS_Supplier b on a.SupplierID = b.SupplierID " & vbCrLf & _
                     " WHERE CreateKanbanCls='1' and FinalApproveDate is not null " & vbCrLf & _
                     " and (select sum(x.POQty) from PO_Detail x where x.PONo = a.PONo and x.AffiliateID = a.AffiliateID and x.SupplierID = a.SupplierID) > 0 "

            ds = GB.uf_GetDataSet(ls_sql)

            If ds.Tables(0).Rows.Count > 0 Then
                Using sqlConn As New SqlConnection(cfg.ConnectionString)
                    sqlConn.Open()

                        '02.  Input data remaining
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CreateKanban")
                            Dim sqlCommNew As SqlCommand = sqlConn.CreateCommand
                            sqlCommNew.Connection = sqlConn
                            sqlCommNew.Transaction = sqlTran

                            pPONo = Trim(ds.Tables(0).Rows(i)("PONo"))
                            pSupplier = Trim(ds.Tables(0).Rows(i)("SupplierID"))
                            pAffiliate = Trim(ds.Tables(0).Rows(i)("AffiliateID"))
                            pAppPeriod = ds.Tables(0).Rows(i)("Period")
                            ls_Prefix = Trim(ds.Tables(0).Rows(i)("LabelPrefix"))

                            '03. Create data Non Kanban

                            ls_sql = " select max(convert(numeric,SUBSTRING(KanbanNo,10,2))) + 1 as SeqNo, AffiliateID  from Kanban_Master " & vbCrLf & _
                                      " where AffiliateID = '" & pAffiliate & "' and SupplierID = '" & pSupplier & "' and YEAR(KanbanDate) = " & Year(pAppPeriod) & " and month(KanbanDate) = " & Month(pAppPeriod) & " and kanbanStatus = '1' " & vbCrLf & _
                                      " group by AffiliateID "
                            sqlCommNew.CommandText = ls_sql
                            Dim daSeqNo As New SqlDataAdapter(sqlCommNew)
                            Dim dsSeqNo As New DataSet
                            daSeqNo.Fill(dsSeqNo)
                            'dsSeqNo = GB.uf_GetDataSet(ls_sql, sqlConn, sqlTran)

                            If dsSeqNo.Tables(0).Rows.Count = 0 Then
                                ls_SeqNo = 1
                            Else
                                ls_SeqNo = dsSeqNo.Tables(0).Rows(0)("SeqNo")
                            End If
                            log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Get data SeqNo, PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]. ok")

                            '04. Check Jumlah Hari
                            If Month(pAppPeriod) = "1" Or Month(pAppPeriod) = "3" Or Month(pAppPeriod) = "5" _
                                Or Month(pAppPeriod) = "7" Or Month(pAppPeriod) = "8" Or Month(pAppPeriod) = "10" _
                                Or Month(pAppPeriod) = "12" Then
                                jlhHari = 31
                            End If

                            If Month(pAppPeriod) = "4" Or Month(pAppPeriod) = "6" Or Month(pAppPeriod) = "9" _
                                Or Month(pAppPeriod) = "11" Then
                                jlhHari = 30
                            End If

                            If Month(pAppPeriod) = "2" Then
                                If Year(pAppPeriod) Mod 4 = 0 Then
                                    jlhHari = 29
                                Else
                                    jlhHari = 28
                                End If
                            End If

                            Dim psequence As Long = 0
                            Dim createSukses As Boolean = False
                            '05. Create Kanban Detail

                            For k = 1 To jlhHari
                                Dim ls_KanbanNo As String = Format(pAppPeriod, "yyyyMM") & IIf(k.ToString.Length = 1, "0" & k, k)
                                Dim ls_Date As String = Format(pAppPeriod, "yyyy-MM-") & IIf(k.ToString.Length = 1, "0" & k, k)
                                Dim ls_KanbanDate As String = IIf(k.ToString.Length = 1, "0" & k, k) & Format(pAppPeriod, "-MM-yyyy")

                                'Create Kanban
                                ls_sql = " select  " & vbCrLf & _
                                              " 	a.AffiliateID, a.PONo, a.SupplierID, c.Period,  " & vbCrLf & _
                                              " 	e.DeliveryLocationCode, b.PartNo, d.UnitCls, CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end QtyBox, " & vbCrLf & _
                                              " 	[DeliveryD" & k & "], " & vbCrLf & _
                                              " 	colcycle1 =  CASE WHEN (CASE WHEN (DeliveryD" & k & " - CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)) > 0  " & vbCrLf & _
                                              "                          THEN CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)  " & vbCrLf & _
                                              "                          ELSE 0 END) = 0 THEN DeliveryD" & k & " ELSE  " & vbCrLf & _
                                              "                          (CASE WHEN (DeliveryD" & k & " - CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)) > 0  " & vbCrLf & _
                                              "                          THEN CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)  " & vbCrLf & _
                                              "                          ELSE 0 END) END ,  " & vbCrLf & _
                                              "  	colcycle2 =  CASE WHEN (DeliveryD" & k & " - CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)) > 0  " & vbCrLf & _
                                              "                          THEN CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)  " & vbCrLf & _
                                              "                          ELSE 0 END ,  " & vbCrLf & _
                                              "  	colcycle3 = CASE WHEN (DeliveryD" & k & " - CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)) > 0  " & vbCrLf & _
                                              "                          THEN CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)  "

                                ls_sql = ls_sql + "                          ELSE 0 END ,  " & vbCrLf & _
                                                  "  	colcycle4 = CASE WHEN (DeliveryD" & k & " - (CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)) * 3) > 0  " & vbCrLf & _
                                                  "                          THEN DeliveryD" & k & " - ((CEILING(FLOOR(DeliveryD" & k & "/4) / CASE WHEN isnull(MPM.QtyBox,0) = 0 then MPM.MOQ else MPM.QtyBox end) * isnull(MPM.QtyBox,0)) )*3  " & vbCrLf & _
                                                  "                          ELSE 0 END " & vbCrLf & _
                                                  " , '' LocationID, '' DockID, ETAAffiliate = '" & ls_Date & "', ETAPASI=mp.ETDPASI " & vbCrLf & _
                                                  " from PO_MasterUpload a " & vbCrLf & _
                                                  " inner join PO_DetailUpload b on a.AffiliateID = b.AffiliateID and a.PONo = b.PONo and a.SupplierID = b.SupplierID " & vbCrLf & _
                                                  " inner join PO_Master c on a.AffiliateID = c.AffiliateID and a.PONo = c.PONo and a.SupplierID = c.SupplierID " & vbCrLf & _
                                                  " left join MS_Parts d on b.PartNo = d.PartNo " & vbCrLf & _
                                                  " left join MS_PartMapping MPM on MPM.PartNo = b.PartNo and MPM.AffiliateID = b.AffiliateID and MPM.SupplierID = b.SupplierID " & vbCrLf & _
                                                  " left join MS_DeliveryPlace e on e.AffiliateID = a.AffiliateID and e.DefaultCls = '1' " & vbCrLf & _
                                                  " left join MS_ETD_PASI mp on mp.AffiliateID = a.AffiliateID AND mp.ETAAffiliate = '" & ls_Date & "' " & vbCrLf & _
                                                  " where a.AffiliateID = '" & pAffiliate & "' and a.PONo = '" & pPONo & "' and a.SupplierID = '" & pSupplier & "' and DeliveryD" & k & " > 0 "
                                sqlCommNew.CommandText = ls_sql
                                Dim daCycle As New SqlDataAdapter(sqlCommNew)
                                Dim dsCycle As New DataSet
                                daCycle.Fill(dsCycle)

                                'dsCycle = GB.uf_GetDataSet(ls_sql, sqlConn, sqlTran)
                                log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Get data cycle, PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]. ok")

                                If dsCycle.Tables(0).Rows.Count > 0 Then
                                    createSukses = True
                                    'Insert Master
                                    ls_sql = " INSERT INTO [dbo].[Kanban_Master] " & vbCrLf & _
                                              "            ([KanbanNo] " & vbCrLf & _
                                              "            ,[AffiliateID] " & vbCrLf & _
                                              "            ,[SupplierID] " & vbCrLf & _
                                              "            ,[KanbanCycle] " & vbCrLf & _
                                              "            ,[KanbanDate] " & vbCrLf & _
                                              "            ,[KanbanTime] " & vbCrLf & _
                                              "            ,[KanbanStatus] " & vbCrLf & _
                                              "            ,[AffiliateApproveUser] " & vbCrLf & _
                                              "            ,[AffiliateApproveDate] " & vbCrLf & _
                                              "            ,[SupplierApproveUser] "

                                    ls_sql = ls_sql + "            ,[SupplierApproveDate] " & vbCrLf & _
                                                      "            ,[EntryDate] " & vbCrLf & _
                                                      "            ,[EntryUser] " & vbCrLf & _
                                                      "            ,[DeliveryLocationCode] " & vbCrLf & _
                                                      "            ,[excelcls]) " & vbCrLf & _
                                                      "      VALUES " & vbCrLf & _
                                                      "            ('" & ls_KanbanNo & "-" & ls_SeqNo & "' " & vbCrLf & _
                                                      "            ,'" & pAffiliate & "' "

                                    ls_sql = ls_sql + "            ,'" & dsCycle.Tables(0).Rows(0)("SupplierID") & "'" & vbCrLf & _
                                                      "            ,'1'" & vbCrLf & _
                                                      "            ,'" & ls_Date & "'" & vbCrLf & _
                                                      "            ,'00:00:00' " & vbCrLf & _
                                                      "            ,'1' " & vbCrLf & _
                                                      "            ,'" & pAffiliate & "' " & vbCrLf & _
                                                      "            ,getdate() " & vbCrLf & _
                                                      "            ,'" & pAffiliate & "' " & vbCrLf & _
                                                      "            ,getdate() " & vbCrLf & _
                                                      "            ,getdate() " & vbCrLf & _
                                                      "            ,'" & pAffiliate & "'"

                                    ls_sql = ls_sql + "            ,'" & dsCycle.Tables(0).Rows(0)("DeliveryLocationCode") & "'" & vbCrLf & _
                                                      "            ,'1')"
                                    'Dim sqlCommHeader As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    'sqlCommHeader.ExecuteNonQuery()
                                    'sqlCommHeader.Dispose()
                                    sqlCommNew.CommandText = ls_sql
                                    x = sqlCommNew.ExecuteNonQuery

                                    log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Insert data Master Kanban, PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]. ok")

                                    'Insert Detail
                                    For n = 0 To dsCycle.Tables(0).Rows.Count - 1
                                        Dim ls_QtyBox As Integer = dsCycle.Tables(0).Rows(n)("QtyBox")
                                        Dim ls_Cycle As Integer = dsCycle.Tables(0).Rows(n)("DeliveryD" & k)
                                        Dim ls_Ulang As Long = ls_Cycle / ls_QtyBox

                                        If ls_Ulang < 1 Then
                                            createSukses = False
                                            GoTo step1
                                        End If

                                        ls_sql = "SELECT ISNULL(MOQ,0) MOQ, ISNULL(QtyBox,0) Qty FROM dbo.MS_PartMapping WHERE PartNo='" + dsCycle.Tables(0).Rows(n)("PartNo") + "' AND SupplierID='" + dsCycle.Tables(0).Rows(n)("SupplierID") + "' AND AffiliateID='" + pAffiliate + "' "
                                        dsMoqQty = GB.uf_GetDataSet(ls_sql, sqlConn, sqlTran)

                                        If dsMoqQty.Tables(0).Rows.Count > 0 Then
                                            POMoq = dsRemaining.Tables(0).Rows(0)("MOQ")
                                            POQty = dsRemaining.Tables(0).Rows(0)("Qty")
                                        End If

                                        ls_sql = " INSERT INTO [dbo].[Kanban_Detail] " & vbCrLf & _
                                                  "            ([KanbanNo] " & vbCrLf & _
                                                  "            ,[AffiliateID] " & vbCrLf & _
                                                  "            ,[SupplierID] " & vbCrLf & _
                                                  "            ,[PartNo] " & vbCrLf & _
                                                  "            ,[PONo] " & vbCrLf & _
                                                  "            ,[DeliveryLocationCode] " & vbCrLf & _
                                                  "            ,[UnitCls] " & vbCrLf & _
                                                  "            ,[KanbanQty] " & vbCrLf & _
                                                  "            ,[POMOQ] " & vbCrLf & _
                                                  "            ,[POQtyBox]) " & vbCrLf & _
                                                  "      VALUES " & vbCrLf & _
                                                  "            ('" & ls_KanbanNo & "-" & ls_SeqNo & "' "

                                        ls_sql = ls_sql + "            ,'" & pAffiliate & "' " & vbCrLf & _
                                                          "            ,'" & dsCycle.Tables(0).Rows(n)("SupplierID") & "' " & vbCrLf & _
                                                          "            ,'" & dsCycle.Tables(0).Rows(n)("PartNo") & "' " & vbCrLf & _
                                                          "            ,'" & dsCycle.Tables(0).Rows(n)("PONo") & "' " & vbCrLf & _
                                                          "            ,'" & dsCycle.Tables(0).Rows(n)("DeliveryLocationCode") & "' " & vbCrLf & _
                                                          "            ,'" & dsCycle.Tables(0).Rows(n)("UnitCls") & "' " & vbCrLf & _
                                                          "            ,'" & ls_Cycle & "' " & vbCrLf & _
                                                          "            ,'" & POMoq & "' " & vbCrLf & _
                                                          "            ,'" & POQty & "') "
                                        'Dim sqlCommDetail As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        'sqlCommDetail.ExecuteNonQuery()
                                        'sqlCommDetail.Dispose()
                                        sqlCommNew.CommandText = ls_sql
                                        x = sqlCommNew.ExecuteNonQuery

                                        For m = 0 To ls_Ulang - 1
                                            Dim temp As String = "32G8," & Trim(dsCycle.Tables(0).Rows(n)("PONo")) & "," & ls_KanbanNo.Trim & "-" & ls_SeqNo & "," & IIf(k.ToString.Length = 1, "0" & k, k) & Format(pAppPeriod, "/MM/yyyy") & "," & Trim(dsCycle.Tables(0).Rows(n)("PartNo")) & "," & Replace(ls_QtyBox, ".00", "") & "," & ls_Prefix & Format(psequence + 1, "00000") '"20160216-2,20160216-2,7184-8880-50,C001"
                                            Dim ls_Barcode As String = "http://zxing.org/w/chart?cht=qr&chs=120x120&chld=L&choe=ISO-8859-1&chl=" & temp

                                            ls_sql = " INSERT INTO [dbo].[Kanban_Barcode] " & vbCrLf & _
                                                      "            ([AffiliateID] " & vbCrLf & _
                                                      "            ,[SupplierID] " & vbCrLf & _
                                                      "            ,[DockID] " & vbCrLf & _
                                                      "            ,[LocationID] " & vbCrLf & _
                                                      "            ,[ETAAffiliate] " & vbCrLf & _
                                                      "            ,[ETAPASI] " & vbCrLf & _
                                                      "            ,[PONo] " & vbCrLf & _
                                                      "            ,[KanbanNo] " & vbCrLf & _
                                                      "            ,[Cycle] " & vbCrLf & _
                                                      "            ,[Partno] " & vbCrLf & _
                                                      "            ,[BoxNo] " & vbCrLf & _
                                                      "            ,[SeqNoStart] " & vbCrLf & _
                                                      "            ,[SeqNoEnd] " & vbCrLf & _
                                                      "            ,[Qty] " & vbCrLf & _
                                                      "            ,[Barcode] " & vbCrLf & _
                                                      "            ,[DeliveryLocationCode] " & vbCrLf & _
                                                      "            ,[barcode2]) " & vbCrLf & _
                                                      "      VALUES "

                                            ls_sql = ls_sql + "            ('" & pAffiliate & "' " & vbCrLf & _
                                                              "            ,'" & dsCycle.Tables(0).Rows(n)("SupplierID") & "' " & vbCrLf & _
                                                              "            ,'" & dsCycle.Tables(0).Rows(n)("DockID") & "' " & vbCrLf & _
                                                              "            ,'" & dsCycle.Tables(0).Rows(n)("LocationID") & "' " & vbCrLf & _
                                                              "            ,'" & dsCycle.Tables(0).Rows(n)("ETAAffiliate") & "' " & vbCrLf & _
                                                              "            ,'" & dsCycle.Tables(0).Rows(n)("ETAPASI") & "' " & vbCrLf & _
                                                              "            ,'" & dsCycle.Tables(0).Rows(n)("PONo") & "' " & vbCrLf & _
                                                              "            ,'" & ls_KanbanNo & "-" & ls_SeqNo & "' " & vbCrLf & _
                                                              "            ,'1' " & vbCrLf & _
                                                              "            ,'" & dsCycle.Tables(0).Rows(n)("PartNo") & "' " & vbCrLf & _
                                                              "            ,'" & ls_Prefix & Format(psequence + 1, "00000") & "' " & vbCrLf & _
                                                              "            ,'" & m + 1 & "' " & vbCrLf & _
                                                              "            ,'" & ls_Ulang & "' " & vbCrLf & _
                                                              "            ,'" & ls_QtyBox & "' " & vbCrLf & _
                                                              "            ,'" & ls_Barcode & "' " & vbCrLf & _
                                                              "            ,'" & dsCycle.Tables(0).Rows(n)("DeliveryLocationCode") & "' " & vbCrLf & _
                                                              "            ,'" & temp & "') "

                                            'Dim sqlCommBarcode As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                            'sqlCommBarcode.ExecuteNonQuery()
                                            'sqlCommBarcode.Dispose()
                                            sqlCommNew.CommandText = ls_sql
                                            x = sqlCommNew.ExecuteNonQuery
                                            psequence = psequence + 1
                                            If psequence > 99999 Then
                                                psequence = 0
                                            End If
                                        Next
                                    Next

                                    log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Insert data Detail Kanban, PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]. ok")

                                End If
                            Next
step1:
                            If createSukses = True Then
                                '02. Update data remaining 
                                ls_sql = "select a.Period, b.PONo, b.AffiliateID, b.SupplierID, b.PartNo, ISNULL(b.POQty,0)POQty from PO_Master a" & vbCrLf & _
                                         "inner join PO_DetailUpload b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.PONo" & vbCrLf & _
                                         "where a.PONo = '" & pPONo & "' and a.SupplierID = '" & pSupplier & "' and a.AffiliateID = '" & pAffiliate & "' and CreateKanbanCls = '1'"
                                dsRemaining = GB.uf_GetDataSet(ls_sql, sqlConn, sqlTran)
                                log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Get data remaining, PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]. ok")

                                If dsRemaining.Tables(0).Rows.Count > 0 Then
                                    For j = 0 To dsRemaining.Tables(0).Rows.Count - 1
                                        ls_sql = " Update RemainingCapacity set QtyRemaining = QtyRemaining - " & dsRemaining.Tables(0).Rows(j)("POQty") & " " & vbCrLf & _
                                                 " WHERE Period = '" & Format(dsRemaining.Tables(0).Rows(j)("Period"), "yyyyMM") & "' and SupplierID = '" & pSupplier & "' and PartNo = '" & dsRemaining.Tables(0).Rows(j)("PartNo") & "'" & vbCrLf
                                        'Dim sqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        'sqlComm.ExecuteNonQuery()
                                        'sqlComm.Dispose()
                                        sqlCommNew.CommandText = ls_sql
                                        x = sqlCommNew.ExecuteNonQuery
                                    Next
                                    log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Update data remaining, PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]. ok")
                                End If


                                ls_sql = " Update PO_Master set CreateKanbanCls = 2 " & vbCrLf & _
                                     " WHERE SupplierID = '" & pSupplier & "' and AffiliateID = '" & pAffiliate & "' and PONo = '" & pPONo & "'" & vbCrLf

                                'Dim sqlCommMaster As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                'sqlCommMaster.ExecuteNonQuery()
                                'sqlCommMaster.Dispose()
                                sqlCommNew.CommandText = ls_sql
                                x = sqlCommNew.ExecuteNonQuery
                                log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Update data Master PO, PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]. ok")

                                sqlTran.Commit()

                                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Final Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok", LogName)
                                LogName.Refresh()
                            Else
                                log.WriteToProcessLog(Date.Now, "FinalApprovePO", "Update data Master PO, PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]. NG")

                                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Final Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] NG", LogName)
                                LogName.Refresh()
                            End If
                        End Using
                    Next                        
                End Using
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] " & ex.Message
            ErrSummary = "PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] " & ex.Message
        Finally
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
            'If Not dsCycle Is Nothing Then
            '    dsCycle.Dispose()
            'End If
            If Not dsRemaining Is Nothing Then
                dsRemaining.Dispose()
            End If
            'If Not dsSeqNo Is Nothing Then
            '    dsSeqNo.Dispose()
            'End If
        End Try
    End Sub

End Class
