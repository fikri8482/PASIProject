Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Imports System.Net
Imports System.IO

Imports System.Windows.Forms
Imports System.Reflection

Public Class clsTmpDB
    Public Shared Function BatchProcessStatus() As String
        Dim name As String = ""

        sql = "SELECT * FROM dbo.BatchProcessStatus "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            BatchProcessStatus = ds.Tables(0).Rows(0)("BatchProcessStatus")
        Else
            BatchProcessStatus = ""
        End If
    End Function

    Public Shared Sub BatchProcessStatusUpdate()
        sql = "UPDATE dbo.BatchProcessStatus SET BatchProcessStatus = '2'"
        uf_ExecuteSql(sql)
    End Sub

    Public Shared Function UnitCls(ByVal untCls As String) As String
        Dim name As String = ""

        sql = "SELECT * FROM dbo.MS_UnitCls WITH(NOLOCK) WHERE [Description] = '" & Trim(untCls) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            UnitCls = ds.Tables(0).Rows(0)("UnitCls")
        Else
            UnitCls = ""
        End If
    End Function

    Public Shared Function UnitClsPart(ByVal PartNo As String) As String
        Dim name As String = ""

        sql = "SELECT TOP 1 * FROM dbo.MS_Parts WITH(NOLOCK) WHERE [PartNo] = '" & Trim(PartNo) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            UnitClsPart = ds.Tables(0).Rows(0)("UnitCls")
        Else
            UnitClsPart = ""
        End If
    End Function

    'Public Shared Function QtyBox(ByVal pSupplier As String, ByVal pAffiliate As String, ByVal pPartNo As String) As String
    '    Dim name As String = ""

    '    sql = "SELECT QtyBox FROM dbo.MS_PartMapping WITH(NOLOCK) WHERE PartNo = '" & Trim(pPartNo) & "' and AffiliateID = '" & Trim(pAffiliate) & "' and SupplierID = '" & Trim(pSupplier) & "' "
    '    ds = uf_GetDataSet(sql)

    '    If ds.Tables(0).Rows.Count > 0 Then
    '        QtyBox = ds.Tables(0).Rows(0)("QtyBox")
    '    Else
    '        QtyBox = "0"
    '    End If
    'End Function

    Public Shared Function POSeqNo(ByVal SeqNo As String) As Integer
        Dim name As String = ""

        sql = "SELECT * FROM dbo.PORev_Master WITH(NOLOCK) WHERE PORevNo = '" & Trim(SeqNo) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            POSeqNo = ds.Tables(0).Rows(0)("SeqNo")
        Else
            POSeqNo = "1"
        End If
    End Function

    Public Shared Function CurrCls(ByVal CurCls As String) As String
        Dim name As String = ""

        sql = "SELECT * FROM MS_CurrCls WITH(NOLOCK) WHERE [Description] = '" & Trim(CurCls) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CurrCls = ds.Tables(0).Rows(0)("CurrCls")
        Else
            CurrCls = ""
        End If
    End Function

    Public Shared Function GetPrice(ByVal pAffiliate As String, ByVal pPartNo As String, ByVal pCurrCls As String) As String
        Dim name As String = ""

        sql = "SELECT * FROM MS_Price WITH(NOLOCK) WHERE AffiliateID = '" & Trim(pAffiliate) & "' AND PartNo = '" & Trim(pPartNo) & "' AND CurrCls = '" & Trim(pCurrCls) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            GetPrice = ds.Tables(0).Rows(0)("Price")
        Else
            GetPrice = 0
        End If
    End Function

    Public Shared Function Attachment(ByVal pTem As clsTmp) As DataSet
        Dim name As String = ""

        sql = "SELECT ISNULL(AttachmentFolder,'')AttachmentFolder,ISNULL(AttachmentBackupFolder,'')AttachmentBackupFolder, " & vbCrLf & _
              "ISNULL(Interval,0)Interval FROM dbo.MS_EmailSetting "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If

    End Function

    Public Shared Function KanbanNo(ByVal PONo As String) As String
        Dim name As String = ""

        sql = "SELECT TOP 1 * FROM dbo.Kanban_Detail WHERE PONo = '" & Trim(PONo) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            KanbanNo = ds.Tables(0).Rows(0)("KanbanNo")
        Else
            KanbanNo = ""
        End If
    End Function

    Public Shared Function POKanbanCls(ByVal pPONo As String, ByVal pPartNo As String, ByVal pAffiliateID As String, ByVal pSupplierID As String) As String
        Dim name As String = ""

        sql = "SELECT KanbanCls FROM dbo.PO_Detail WITH(NOLOCK) WHERE PONo = '" & Trim(pPONo) & "' AND PartNo = '" & Trim(pPartNo) & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            POKanbanCls = ds.Tables(0).Rows(0)("KanbanCls")
        Else
            POKanbanCls = ""
        End If
    End Function

    Public Shared Function POReceive(ByVal pPONo As String, ByVal pPartNo As String, ByVal pAffiliateID As String, ByVal pSupplierID As String, ByVal pQty As Double) As Boolean
        Dim name As String = ""

        sql = "  Select Remaining = POD.POQty - ISNULL(RD.Qty,0)  " & vbCrLf & _
              "  From PO_Detail POD WITH(NOLOCK) " & vbCrLf & _
              "  LEFT JOIN (Select SupplierID,AffiliateID,PONo,PartNo,Qty=SUM(GoodRecQty)   " & vbCrLf & _
              "  		   From ReceivePASI_Detail WITH(NOLOCK) " & vbCrLf & _
              "  		   Group By SupplierID,AffiliateID,PONo,PartNo) RD  " & vbCrLf & _
              "  	   ON POD.PONo = RD.PONo  " & vbCrLf & _
              "  	   AND POD.AffiliateID = RD.AffiliateID  " & vbCrLf & _
              "  	   AND POD.SupplierID = RD.SupplierID  " & vbCrLf & _
              "  	   AND POD.PartNo = RD.PartNo " & vbCrLf & _
              "  Where POD.PONo = '" & Trim(pPONo) & "' " & vbCrLf & _
              "  AND POD.AffiliateID = '" & Trim(pAffiliateID) & "'  "

        sql = sql + "  AND POD.SupplierID = '" & Trim(pSupplierID) & "'  " & vbCrLf & _
                    "  AND POD.PartNo = '" & Trim(pPartNo) & "'  "

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            If (ds.Tables(0).Rows(0)("Remaining") - CDbl(pQty)) < 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Shared Function StatusDelivery(ByVal pPONo As String) As DataSet
        Dim name As String = "" '0 = Aff, 1 = PASI

        sql = "SELECT *, ISNULL(Amount,0)AmountT FROM Po_Master WITH(NOLOCK) WHERE ISNULL(DeliveryByPASICls,'0') = '1' AND PONo = '" & Trim(pPONo) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetRecValue(ByVal pPONo As String, ByVal tblName As String, ByVal pSuratJalan As String, ByVal pSupplierID As String _
                                       , ByVal pAffiliate As String, ByVal pKanbanNo As String, ByVal pPartNo As String) As DataSet
        Dim name As String = "" '0 = Aff, 1 = PASI

        'sql = " SELECT * FROM " & tblName & " WHERE SuratJalanNo = '" & Trim(pSuratJalan) & "' AND SupplierID = '" & Trim(pSupplierID) & "' " & vbCrLf & _
        '      " AND AffiliateID = '" & Trim(pAffiliate) & "' AND PoNo = '" & Trim(pPONo) & "' AND KanbanNo = '" & Trim(pKanbanNo) & "' AND PartNo = '" & Trim(pPartNo) & "' "

        sql = " SELECT * FROM " & tblName & " WITH(NOLOCK) WHERE SupplierID = '" & Trim(pSupplierID) & "' " & vbCrLf & _
              " AND AffiliateID = '" & Trim(pAffiliate) & "' AND PoNo = '" & Trim(pPONo) & "' AND KanbanNo = '" & Trim(pKanbanNo) & "' AND PartNo = '" & Trim(pPartNo) & "' "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetReceive(ByVal pStatus As String) As DataSet
        Dim name As String = ""

        sql = "SELECT ISNULL(AttachmentFolder,'')AttachmentFolder,ISNULL(AttachmentBackupFolder,'')AttachmentBackupFolder, " & vbCrLf & _
              "ISNULL(Interval,0)Interval FROM dbo.MS_EmailSetting "
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If

    End Function

    Public Shared Function CekKanbanAutoApprover(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer
        sql = "SELECT * FROM dbo.Kanban_Master" & vbCrLf & _
            " WHERE KanbanDate = '" & Trim(pTmp.KanbanDate) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
            " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierApproveUser = 'AUTO APPROVED' "

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekKanbanAutoApprover = 1
        Else
            CekKanbanAutoApprover = 0
        End If

    End Function

    Public Shared Function CekPOAutoApprover(ByVal tblName As String, ByVal pPONo As String, ByVal pSupplierID As String, ByVal pStatus As Integer) As Integer
        Dim name As String = ""

        sql = "SELECT * FROM " & tblName & " WITH(NOLOCK) WHERE " & vbCrLf

        If pStatus = "0" Then
            sql = sql + " SupplierApproveUser = 'AUTO APPROVED' " & vbCrLf
        ElseIf pStatus = "1" Then
            sql = sql + " SupplierApprovePendingUser = 'AUTO APPROVED'" & vbCrLf
        ElseIf pStatus = "2" Then
            sql = sql + " SupplierUnApproveUser = 'AUTO APPROVED'" & vbCrLf
        End If

        If tblName = "dbo.PORev_Master" Then
            sql = sql + " AND PORevNo = '" & Trim(pPONo) & "' and SupplierID ='" & Trim(pSupplierID) & "'"
        Else
            sql = sql + " AND PONo = '" & Trim(pPONo) & "' and SupplierID ='" & Trim(pSupplierID) & "'"
        End If

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOAutoApprover = 1
        Else
            CekPOAutoApprover = 0
        End If
    End Function

    Public Shared Function UpdateKanban(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer

        sql = " UPDATE dbo.Kanban_Master " & vbCrLf & _
                "   SET SupplierApproveDate = GETDATE(), " & vbCrLf & _
                " 	    SupplierApproveUser = '" & Trim(pTmp.PIC) & "' " & vbCrLf & _
                " WHERE KanbanDate = '" & Trim(pTmp.KanbanDate) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
                " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        Return i
    End Function

    Public Shared Function UpdatePOMaster(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal pStatus As String, ByVal tblName As String) As Integer
        '--
        sql = " UPDATE " & tblName & " " & vbCrLf & _
            "   SET " & vbCrLf
        If pStatus = "0" Then
            sql = sql + " SupplierApproveDate = GETDATE(), " & vbCrLf & _
                        " SupplierApproveUser = '" & Trim(pTmp.SupplierApproveUser) & "' " & vbCrLf
        ElseIf pStatus = "1" Then
            sql = sql + " SupplierApprovePendingDate = GETDATE(), " & vbCrLf & _
                        " SupplierApprovePendingUser = '" & Trim(pTmp.SupplierApproveUser) & "' " & vbCrLf
        ElseIf pStatus = "2" Then
            sql = sql + " SupplierUnApproveDate = GETDATE(), " & vbCrLf & _
                        " SupplierUnApproveUser = '" & Trim(pTmp.SupplierApproveUser) & "' " & vbCrLf
        End If

        If tblName = "dbo.PORev_Master" Then
            sql = sql + " WHERE PORevNo = '" & Trim(pTmp.PORevNo) & "' "
        Else
            sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' "
        End If

        sql = sql + " and SupplierID = '" & pTmp.SupplierID & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        Return i
    End Function

    Public Shared Function UpdatePOMasterUpload(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal tblName As String, ByVal templateCode As String) As Integer

        sql = " UPDATE " & tblName & " " & vbCrLf & _
            "   SET Remarks = '" & Trim(pTmp.Remarks) & "' " & vbCrLf & _
            " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' " & vbCrLf
        If templateCode = "POR" Then
            sql = sql + " AND PORevNo = '" & Trim(pTmp.PORevNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND SeqNo = " & pTmp.POSeqNo & " "
        Else
            sql = sql + " AND SupplierID = '" & Trim(pTmp.SupplierID) & "' "
        End If

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        Return i
    End Function

    Public Shared Function insertMasterKanban(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal pKanbanNo As String, ByVal pKanbanCycle As Integer) As Integer

        sql = " UPDATE dbo.Kanban_Master " & vbCrLf & _
            " SET KanbanDate = '" & pTmp.KanbanDate & "', " & vbCrLf & _
            " 	KanbanTime = '" & pTmp.KanbanTime & "', " & vbCrLf & _
            " 	KanbanStatus = '" & Trim(pTmp.KanbanStatus) & "', " & vbCrLf & _
            " 	AffiliateApproveUser = '" & Trim(pTmp.AffiliateApproveUser) & "', " & vbCrLf & _
            " 	AffiliateApproveDate = '" & Trim(pTmp.AffiliateApproveDate) & "', " & vbCrLf & _
            " 	SupplierApproveUser = '" & Trim(pTmp.SupplierApproveUser) & "', " & vbCrLf & _
            " 	SupplierApproveDate = '" & Trim(pTmp.SupplierApproveDate) & "', " & vbCrLf & _
            " 	EntryDate = '" & Trim(pTmp.EntryDate) & "', " & vbCrLf & _
            " 	EntryUser = '" & Trim(pTmp.EntryUser) & "', " & vbCrLf & _
            " 	UpdateDate = '" & Trim(pTmp.UpdateDate) & "', " & vbCrLf

        sql = sql + " 	UpdateUser = '" & Trim(pTmp.UpdateUser) & "', " & vbCrLf & _
                    " 	DeliveryLocationCode = '" & Trim(pTmp.DeliveryLocation) & "' " & vbCrLf & _
                    " WHERE KanbanNo = '" & Trim(pKanbanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
                    " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.Kanban_Master " & vbCrLf & _
                        " VALUES  ( '" & Trim(pKanbanNo) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierID) & "' ," & vbCrLf & _
                        "           '" & Trim(pKanbanCycle) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.KanbanDate) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.KanbanTime) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.KanbanStatus) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateApproveUser) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateApproveDate) & "' ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierApproveUser) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierApproveDate) & "' , " & vbCrLf

            sql = sql + "           GETDATE() ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.EntryUser) & "' ,  " & vbCrLf & _
                        "           GETDATE() , " & vbCrLf & _
                        "           '" & Trim(pTmp.UpdateUser) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.DeliveryLocation) & "'   " & vbCrLf & _
                        "         ) "

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()
        End If
        Return i
    End Function

    Public Shared Function insertDetailKanban(ByVal pTmp As clsTmp, ByVal pKanbanNo As String) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " UPDATE dbo.Kanban_Detail " & vbCrLf & _
                        " SET KanbanQty = " & Trim(pTmp.KanbanQty) & " " & vbCrLf & _
                        " WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                        " AND KanbanNo = '" & Trim(pKanbanNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' "

            Dim SQLCom As New SqlCommand(sql, SQLCon)

            Dim i As Integer = SQLCom.ExecuteNonQuery

            If i = 0 Then
                sql = " INSERT INTO dbo.Kanban_Detail " & vbCrLf & _
                            " VALUES  ( '" & Trim(pKanbanNo) & "' , -- SuratJalanNo - char(20) " & vbCrLf & _
                            "           '" & Trim(pTmp.AffiliateID) & "' , -- SupplierID - char(20) " & vbCrLf & _
                            "           '" & Trim(pTmp.SupplierID) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                            "           '" & Trim(pTmp.PartNo) & "' , -- PONo - char(20) " & vbCrLf & _
                            "           '" & Trim(pTmp.PONo) & "' , -- POKanbanCls - char(1) " & vbCrLf & _
                            "           '" & Trim(pTmp.UnitCls) & "' , -- KanbanNo - char(20) " & vbCrLf & _
                            "           '" & Trim(pTmp.KanbanQty) & "',  -- DefectRecQty - numeric " & vbCrLf & _
                            "           '" & uf_GetMOQ(0, Trim(pTmp.PONo), Trim(pTmp.PartNo), Trim(pTmp.SupplierID), Trim(pTmp.AffiliateID)) & "',  -- POMOQ - numeric " & vbCrLf & _
                            "           '" & uf_GetQtybox(0, Trim(pTmp.PONo), Trim(pTmp.PartNo), Trim(pTmp.SupplierID), Trim(pTmp.AffiliateID)) & "'  -- POQtyBox - numeric " & vbCrLf & _
                            "       ) "

                SQLCom.CommandText = sql
                i = SQLCom.ExecuteNonQuery()
            End If

            Return i
        End Using

    End Function

    Public Shared Function insertMasterDO(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer

        sql = " UPDATE dbo.DOSupplier_Master " & vbCrLf & _
            " SET DeliveryDate = '" & pTmp.DeliveryDate & "', " & vbCrLf & _
            " 	PIC = '" & Trim(pTmp.PIC) & "', " & vbCrLf & _
            " 	JenisArmada = '" & Trim(pTmp.JenisArmada) & "', " & vbCrLf & _
            " 	DriverName = '" & Trim(pTmp.DriverName) & "', " & vbCrLf & _
            " 	DriverContact = '" & Trim(pTmp.DriverCont) & "', " & vbCrLf & _
            " 	NoPol = '" & Trim(pTmp.NoPol) & "', " & vbCrLf & _
            " 	TotalBox = '" & Trim(pTmp.TotalBox) & "', " & vbCrLf & _
            " 	EntryDate = '" & Trim(pTmp.EntryDate) & "', " & vbCrLf & _
            " 	EntryUser = '" & Trim(pTmp.EntryUser) & "', " & vbCrLf & _
            " 	UpdateDate = '" & Trim(pTmp.UpdateDate) & "', " & vbCrLf

        sql = sql + " 	UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
                    " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
                    " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.DOSupplier_Master " & vbCrLf & _
                        " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.DeliveryDate) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.PIC) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.JenisArmada) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.DriverName) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.DriverCont) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.NoPol) & "' ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.TotalBox) & "' ," & vbCrLf

            sql = sql + "           GETDATE() ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.EntryUser) & "' ,  " & vbCrLf & _
                        "           GETDATE() , " & vbCrLf & _
                        "           '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
                        "         ) "

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()
        End If
        Return i
    End Function

    Public Shared Function insertDetailDO(ByVal pTmp As clsTmp, ByVal pMoq As Integer, ByVal pQtyBox As Integer, ByVal SQLCom As SqlCommand) As Integer

        'Using SQLCon As New SqlConnection(uf_GetConString)
        'SQLCon.Open()

        sql = " UPDATE dbo.DOSupplier_Detail " & vbCrLf & _
              " SET DOQty = " & Trim(pTmp.DOQty) & " " & vbCrLf & _
              " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
              " AND KanbanNo = '" & Trim(pTmp.KanbanNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' "

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.DOSupplier_Detail " & vbCrLf & _
                        " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.PONo) & "' ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.POKanbanCls) & "' ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.KanbanNo) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.PartNo) & "',  " & vbCrLf & _
                        "           '" & Trim(pTmp.UnitCls) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.DOQty) & "',  " & vbCrLf & _
                        "           '" & pMoq & "',  " & vbCrLf & _
                        "           '" & pQtyBox & "'  " & vbCrLf & _
                        "       ) "

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()
        End If

        Return i
        'End Using
    End Function

    'Public Shared Function insertDetailDO(ByVal pTmp As clsTmp) As Integer

    '    Using SQLCon As New SqlConnection(uf_GetConString)
    '        SQLCon.Open()

    '        sql = " UPDATE dbo.DOSupplier_Detail " & vbCrLf & _
    '                    " SET DOQty = " & Trim(pTmp.DOQty) & " " & vbCrLf & _
    '                    " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
    '                    " AND KanbanNo = '" & Trim(pTmp.KanbanNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' "

    '        Dim SQLCom As New SqlCommand(sql, SQLCon)

    '        Dim i As Integer = SQLCom.ExecuteNonQuery

    '        If i = 0 Then
    '            sql = " INSERT INTO dbo.DOSupplier_Detail " & vbCrLf & _
    '                        " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' , " & vbCrLf & _
    '                        "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
    '                        "           '" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
    '                        "           '" & Trim(pTmp.PONo) & "' ,  " & vbCrLf & _
    '                        "           '" & Trim(pTmp.POKanbanCls) & "' ,  " & vbCrLf & _
    '                        "           '" & Trim(pTmp.KanbanNo) & "' , " & vbCrLf & _
    '                        "           '" & Trim(pTmp.PartNo) & "',  " & vbCrLf & _
    '                        "           '" & Trim(pTmp.UnitCls) & "' , " & vbCrLf & _
    '                        "           '" & Trim(pTmp.DOQty) & "'  " & vbCrLf & _
    '                        "       ) "

    '            SQLCom.CommandText = sql
    '            i = SQLCom.ExecuteNonQuery()
    '        End If

    '        Return i
    '    End Using

    'End Function

    Public Shared Function CekValidasiQtyDN(ByVal pTmp As clsTmp) As Boolean
        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            sql = "Select * From Kanban_Detail Where PartNo = '" & Trim(pTmp.PartNo) & "' And PONo = '" & Trim(pTmp.PONo) & "' And UnitCls = '" & Trim(pTmp.UnitCls) & "' And KanbanNo = '" & Trim(pTmp.POKanbanCls) & "'"
            ds = uf_GetDataSet(sql)

            If ds.Tables(0).Rows.Count > 0 Then
                If pTmp.DOQty > Trim(ds.Tables(0).Rows(0)("KanbanQty")) Then
                    Return True
                End If
            End If

        End Using
    End Function

    Public Shared Function insertMasterPOUpload(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal tblName As String, ByVal templateCode As String) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)

            SQLCon.Open()

            sql = " UPDATE " & tblName & " " & vbCrLf & _
                  " SET Remarks = '" & Trim(pTmp.Remarks) & "' " & vbCrLf

            If tblName = "dbo.PORev_MasterUpload" Then
                sql = sql + " WHERE PORevNo = '" & Trim(pTmp.PORevNo) & "' AND PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND SeqNo = " & pTmp.POSeqNo & " "
            Else
                sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' "
            End If

            SQLCom.CommandText = sql
            Dim i As Integer = SQLCom.ExecuteNonQuery

            If i = 0 Then
                sql = " INSERT INTO " & tblName & " " & vbCrLf
                If templateCode = "POR" Then
                    sql = sql + " VALUES  ( '" & Trim(pTmp.PORevNo) & "' ,'" & Trim(pTmp.PONo) & "' , -- PONo - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.AffiliateID) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "' , " & Trim(pTmp.POSeqNo) & " , -- SupplierID - char(20) " & vbCrLf
                Else
                    sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "' , -- PONo - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.AffiliateID) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "',  -- SupplierID - char(20) " & vbCrLf
                End If

                sql = sql + "           '" & (pTmp.Remarks) & "' , -- Remarks - date " & vbCrLf & _
                            "           GETDATE() , -- EntryDate - datetime " & vbCrLf & _
                            "           '" & Trim(pTmp.EntryUser) & "' , -- EntryUser - char(15) " & vbCrLf & _
                            "           GETDATE() , -- UpdateDate - datetime " & vbCrLf & _
                            "           '" & Trim(pTmp.UpdateUser) & "'  -- UpdateUser - char(15) " & vbCrLf & _
                            "         ) "


                SQLCom.CommandText = sql
                i = SQLCom.ExecuteNonQuery()
            End If

            Return i
        End Using

    End Function

    Public Shared Function insertDetailPOUpload(ByVal pTmp As clsTmp, _
            ByVal DelD1 As Integer, ByVal DelD1Old As Integer, _
            ByVal DelD2 As Integer, ByVal DelD2Old As Integer, _
            ByVal DelD3 As Integer, ByVal DelD3Old As Integer, _
            ByVal DelD4 As Integer, ByVal DelD4Old As Integer, _
            ByVal DelD5 As Integer, ByVal DelD5Old As Integer, _
            ByVal DelD6 As Integer, ByVal DelD6Old As Integer, _
            ByVal DelD7 As Integer, ByVal DelD7Old As Integer, _
            ByVal DelD8 As Integer, ByVal DelD8Old As Integer, _
            ByVal DelD9 As Integer, ByVal DelD9Old As Integer, _
            ByVal DelD10 As Integer, ByVal DelD10Old As Integer, _
            ByVal DelD11 As Integer, ByVal DelD11Old As Integer, _
            ByVal DelD12 As Integer, ByVal DelD12Old As Integer, _
            ByVal DelD13 As Integer, ByVal DelD13Old As Integer, _
            ByVal DelD14 As Integer, ByVal DelD14Old As Integer, _
            ByVal DelD15 As Integer, ByVal DelD15Old As Integer, _
            ByVal DelD16 As Integer, ByVal DelD16Old As Integer, _
            ByVal DelD17 As Integer, ByVal DelD17Old As Integer, _
            ByVal DelD18 As Integer, ByVal DelD18Old As Integer, _
            ByVal DelD19 As Integer, ByVal DelD19Old As Integer, _
            ByVal DelD20 As Integer, ByVal DelD20Old As Integer, _
            ByVal DelD21 As Integer, ByVal DelD21Old As Integer, _
            ByVal DelD22 As Integer, ByVal DelD22Old As Integer, _
            ByVal DelD23 As Integer, ByVal DelD23Old As Integer, _
            ByVal DelD24 As Integer, ByVal DelD24Old As Integer, _
            ByVal DelD25 As Integer, ByVal DelD25Old As Integer, _
            ByVal DelD26 As Integer, ByVal DelD26Old As Integer, _
            ByVal DelD27 As Integer, ByVal DelD27Old As Integer, _
            ByVal DelD28 As Integer, ByVal DelD28Old As Integer, _
            ByVal DelD29 As Integer, ByVal DelD29Old As Integer, _
            ByVal DelD30 As Integer, ByVal DelD30Old As Integer, _
            ByVal DelD31 As Integer, ByVal DelD31Old As Integer, ByVal tblName As String, ByVal templateCode As String) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " UPDATE " & tblName & " " & vbCrLf & _
                  " SET EntryDate = GETDATE() " & vbCrLf

            If tblName = "dbo.PORev_DetailUpload" Then
                sql = sql + " WHERE PORevNo = '" & Trim(pTmp.PORevNo) & "' AND PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND SeqNo = " & pTmp.POSeqNo & " AND PartNo = '" & Trim(pTmp.PartNo) & "' "
            Else
                sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' "
            End If

            Dim SQLCom As New SqlCommand(sql, SQLCon)

            Dim i As Integer = SQLCom.ExecuteNonQuery

            If i = 0 Then

                sql = " INSERT INTO " & tblName & "  " & vbCrLf

                If templateCode = "POR" Then
                    sql = sql + " VALUES  ( '" & Trim(pTmp.PORevNo) & "','" & Trim(pTmp.PONo) & "' , -- PONo - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.AffiliateID) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "' , -- SupplierID - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.PartNo) & "' , " & Trim(pTmp.POSeqNo) & ", -- PartNo - char(25) " & vbCrLf

                Else
                    sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "' , -- PONo - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.AffiliateID) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "' , -- SupplierID - char(20) " & vbCrLf & _
                                "           '" & Trim(pTmp.PartNo) & "' , -- PartNo - char(25) " & vbCrLf
                End If

                sql = sql + "           '" & Trim(pTmp.DifferentCls) & "', " & vbCrLf & _
                            "           '" & Trim(pTmp.POKanbanCls) & "', " & vbCrLf & _
                            "           NULL, " & vbCrLf
                sql = sql + "           '" & Trim(pTmp.POQty) & "' , -- POQty - numeric " & vbCrLf & _
                            "           '" & Trim(pTmp.POQtyOld) & "' , -- POQtyOld - numeric " & vbCrLf & _
                            "           '" & Trim(pTmp.CurrCls) & "', --CurrCls " & vbCrLf & _
                            "           " & pTmp.Price & " ," & vbCrLf & _
                            "           " & pTmp.Amount & ", " & vbCrLf & _
                            "           " & DelD1 & " , -- DeliveryD1 - numeric " & vbCrLf & _
                            "           " & DelD1Old & " , -- DeliveryD1Old - numeric " & vbCrLf & _
                            "           " & DelD2 & " , -- DeliveryD2 - numeric " & vbCrLf & _
                            "           " & DelD2Old & " , -- DeliveryD2Old - numeric " & vbCrLf & _
                            "           " & DelD3 & " , -- DeliveryD3 - numeric " & vbCrLf & _
                            "           " & DelD3Old & " , -- DeliveryD3Old - numeric " & vbCrLf & _
                            "           " & DelD4 & " , -- DeliveryD4 - numeric " & vbCrLf & _
                            "           " & DelD4Old & " , -- DeliveryD4Old - numeric " & vbCrLf & _
                            "           " & DelD5 & " , -- DeliveryD5 - numeric " & vbCrLf

                sql = sql + "           " & DelD5Old & " , -- DeliveryD5Old - numeric " & vbCrLf & _
                            "           " & DelD6 & " , -- DeliveryD6 - numeric " & vbCrLf & _
                            "           " & DelD6Old & " , -- DeliveryD6Old - numeric " & vbCrLf & _
                            "           " & DelD7 & " , -- DeliveryD7 - numeric " & vbCrLf & _
                            "           " & DelD7Old & " , -- DeliveryD7Old - numeric " & vbCrLf & _
                            "           " & DelD8 & " , -- DeliveryD8 - numeric " & vbCrLf & _
                            "           " & DelD8Old & " , -- DeliveryD8Old - numeric " & vbCrLf & _
                            "           " & DelD9 & " , -- DeliveryD9 - numeric " & vbCrLf & _
                            "           " & DelD9Old & " , -- DeliveryD9Old - numeric " & vbCrLf & _
                            "           " & DelD10 & " , -- DeliveryD10 - numeric " & vbCrLf & _
                            "           " & DelD10Old & " , -- DeliveryD10Old - numeric " & vbCrLf

                sql = sql + "           " & DelD11 & " , -- DeliveryD11 - numeric " & vbCrLf & _
                            "           " & DelD11Old & " , -- DeliveryD11Old - numeric " & vbCrLf & _
                            "           " & DelD12 & " , -- DeliveryD12 - numeric " & vbCrLf & _
                            "           " & DelD12Old & " , -- DeliveryD12Old - numeric " & vbCrLf & _
                            "           " & DelD13 & " , -- DeliveryD13 - numeric " & vbCrLf & _
                            "           " & DelD13Old & " , -- DeliveryD13Old - numeric " & vbCrLf & _
                            "           " & DelD14 & " , -- DeliveryD14 - numeric " & vbCrLf & _
                            "           " & DelD14Old & " , -- DeliveryD14Old - numeric " & vbCrLf & _
                            "           " & DelD15 & " , -- DeliveryD15 - numeric " & vbCrLf & _
                            "           " & DelD15Old & " , -- DeliveryD15Old - numeric " & vbCrLf & _
                            "           " & DelD16 & " , -- DeliveryD16 - numeric " & vbCrLf

                sql = sql + "           " & DelD16Old & " , -- DeliveryD16Old - numeric " & vbCrLf & _
                            "           " & DelD17 & " , -- DeliveryD17 - numeric " & vbCrLf & _
                            "           " & DelD17Old & " , -- DeliveryD17Old - numeric " & vbCrLf & _
                            "           " & DelD18 & " , -- DeliveryD18 - numeric " & vbCrLf & _
                            "           " & DelD18Old & " , -- DeliveryD18Old - numeric " & vbCrLf & _
                            "           " & DelD19 & " , -- DeliveryD19 - numeric " & vbCrLf & _
                            "           " & DelD19Old & " , -- DeliveryD19Old - numeric " & vbCrLf & _
                            "           " & DelD20 & " , -- DeliveryD20 - numeric " & vbCrLf & _
                            "           " & DelD20Old & " , -- DeliveryD20Old - numeric " & vbCrLf & _
                            "           " & DelD21 & " , -- DeliveryD21 - numeric " & vbCrLf & _
                            "           " & DelD21Old & " , -- DeliveryD21Old - numeric " & vbCrLf

                sql = sql + "           " & DelD22 & " , -- DeliveryD22 - numeric " & vbCrLf & _
                            "           " & DelD22Old & " , -- DeliveryD22Old - numeric " & vbCrLf & _
                            "           " & DelD23 & " , -- DeliveryD23 - numeric " & vbCrLf & _
                            "           " & DelD23Old & " , -- DeliveryD23Old - numeric " & vbCrLf & _
                            "           " & DelD24 & " , -- DeliveryD24 - numeric " & vbCrLf & _
                            "           " & DelD24Old & " , -- DeliveryD24Old - numeric " & vbCrLf & _
                            "           " & DelD25 & " , -- DeliveryD25 - numeric " & vbCrLf & _
                            "           " & DelD25Old & " , -- DeliveryD25Old - numeric " & vbCrLf & _
                            "           " & DelD26 & " , -- DeliveryD26 - numeric " & vbCrLf & _
                            "           " & DelD26Old & " , -- DeliveryD26Old - numeric " & vbCrLf & _
                            "           " & DelD27 & " , -- DeliveryD27 - numeric " & vbCrLf

                sql = sql + "           " & DelD27Old & " , -- DeliveryD27Old - numeric " & vbCrLf & _
                            "           " & DelD28 & " , -- DeliveryD28 - numeric " & vbCrLf & _
                            "           " & DelD28Old & " , -- DeliveryD28Old - numeric " & vbCrLf & _
                            "           " & DelD29 & " , -- DeliveryD29 - numeric " & vbCrLf & _
                            "           " & DelD29Old & " , -- DeliveryD29Old - numeric " & vbCrLf & _
                            "           " & DelD30 & " , -- DeliveryD30 - numeric " & vbCrLf & _
                            "           " & DelD30Old & " , -- DeliveryD30Old - numeric " & vbCrLf & _
                            "           " & DelD31 & " , -- DeliveryD31 - numeric " & vbCrLf & _
                            "           " & DelD31Old & " , -- DeliveryD31Old - numeric " & vbCrLf & _
                            "           GETDATE() , -- EntryDate - datetime " & vbCrLf & _
                            "           '' , -- EntryUser - char(15) " & vbCrLf

                sql = sql + "           GETDATE() , -- UpdateDate - datetime " & vbCrLf & _
                            "           ''  -- UpdateUser - char(15) " & vbCrLf & _
                            "         ) "

                SQLCom.CommandText = sql
                i = SQLCom.ExecuteNonQuery()
            End If

            Return i
        End Using

    End Function

    Public Shared Function insertMasterInvoice(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer

        sql = " UPDATE dbo.InvoiceSupplier_Master " & vbCrLf & _
            " SET InvoiceDate = '" & pTmp.InvoiceDate & "', " & vbCrLf & _
            " 	DueDate = '" & pTmp.DueDate & "', " & vbCrLf & _
            " 	PaymentTerm = '" & Trim(pTmp.PaymentItem) & "', " & vbCrLf & _
            " 	TotalAmount = '" & Trim(pTmp.TotalAmount) & "', " & vbCrLf & _
            " 	EntryDate = '" & Trim(pTmp.EntryDate) & "', " & vbCrLf & _
            " 	EntryUser = '" & Trim(pTmp.EntryUser) & "', " & vbCrLf & _
            " 	UpdateDate = '" & Trim(pTmp.UpdateDate) & "', " & vbCrLf

        sql = sql + " 	UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
                    " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
                    " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND InvoiceNo = '" & Trim(pTmp.InvoiceNo) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.InvoiceSupplier_Master " & vbCrLf & _
                        " VALUES  ( '" & Trim(pTmp.InvoiceNo) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.SuratJalanNo) & "' ," & vbCrLf & _
                        "           '" & pTmp.InvoiceDate & "' ," & vbCrLf & _
                        "           '" & pTmp.DueDate & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.PaymentItem) & "' ," & vbCrLf & _
                        "           '" & pTmp.TotalAmount & "' , " & vbCrLf

            sql = sql + "           GETDATE() ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.EntryUser) & "' ,  " & vbCrLf & _
                        "           GETDATE() , " & vbCrLf & _
                        "           '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
                        "         ) "

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()
        End If
        Return i
    End Function

    Public Shared Function insertDetailInvoice(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer

        'Using SQLCon As New SqlConnection(uf_GetConString)
        '    SQLCon.Open()

        sql = " UPDATE dbo.InvoiceSupplier_Detail " & vbCrLf & _
                    " SET ReceiveQty = " & Trim(pTmp.ReceiveQty) & ", " & vbCrLf & _
                    "     ReceiveCurrCls = '" & Trim(pTmp.ReceiveCurrCls) & "', " & vbCrLf & _
                    "     ReceivePrice = '" & Trim(pTmp.ReceivePrice) & "', " & vbCrLf & _
                    "     ReceiveAmount = '" & Trim(pTmp.ReceiveAmount) & "', " & vbCrLf & _
                    "     InvQty = " & Trim(pTmp.InvoiceQty) & ", " & vbCrLf & _
                    "     InvCurrCls = '" & Trim(pTmp.InvCurrCls) & "', " & vbCrLf & _
                    "     InvPrice = '" & Trim(pTmp.InvPrice) & "', " & vbCrLf & _
                    "     InvAmount = '" & Trim(pTmp.InvAmount) & "' " & vbCrLf & _
                    " WHERE InvoiceNo = '" & Trim(pTmp.InvoiceNo) & "' AND SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                    " AND KanbanNo = '" & Trim(pTmp.KanbanNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' "

        'Dim SQLCom As New SqlCommand(sql, SQLCon)
        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.InvoiceSupplier_Detail " & vbCrLf & _
                        " VALUES  ( '" & Trim(pTmp.InvoiceNo) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateID) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.SuratJalanNo) & "' ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.PONo) & "' ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.POKanbanCls) & "' ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.KanbanNo) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.PartNo) & "',  " & vbCrLf & _
                        "            " & Trim(pTmp.ReceiveQty) & " , " & vbCrLf & _
                        "           '" & Trim(pTmp.ReceiveCurrCls) & "',  " & vbCrLf & _
                        "           '" & Trim(pTmp.ReceivePrice) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.ReceiveAmount) & "',  " & vbCrLf & _
                        "            " & Trim(pTmp.InvoiceQty) & " , " & vbCrLf & _
                        "           '" & Trim(pTmp.InvCurrCls) & "',  " & vbCrLf & _
                        "           '" & Trim(pTmp.InvPrice) & "' , " & vbCrLf & _
                        "           '" & Trim(pTmp.InvAmount) & "'  " & vbCrLf & _
                        "       ) "

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()
        End If

        Return i
        'End Using

    End Function

    Shared Function GetServerDate(ByVal pConStr As String) As DateTime

        sql = "select CAST(GETDATE() AS DATETIME) as tanggal "

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            GetServerDate = ds.Tables(0).Rows(0)("tanggal")
        End If
    End Function

    Public Shared Function CekPOMonthlyAutoApprover(ByVal tblName As String, ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pStatus As Integer) As Integer
        Dim name As String = ""

        sql = "SELECT * FROM " & tblName & " WITH(NOLOCK) WHERE " & vbCrLf

        If pStatus = "0" Then
            sql = sql + " SupplierApproveUser = 'AUTO APPROVED' " & vbCrLf
        Else
            sql = sql + " SupplierApproveUser = 'AUTO APPROVED' " & vbCrLf
        End If

        If tblName = "dbo.PO_Master_Export" Then
            sql = sql + " AND PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo1 = '" & Trim(pOrderNo) & "'"
        Else
            sql = sql + " AND PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo1 = '" & Trim(pOrderNo) & "'"
        End If

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOMonthlyAutoApprover = 1
        Else
            CekPOMonthlyAutoApprover = 0
        End If
    End Function

    Public Shared Function CekDOBoxDetail(ByVal tblName As String, ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pPartNo As String) As String
        Dim name As String = ""
        sql = "select isnull((Right(max(boxno),6) + 1),'') as boxno From DOSupplier_DetailBox_Export where pono = '" & Trim(pPONo) & "' and AffiliateID = '" & Trim(pAffiliateID) & "' and OrderNo = '" & Trim(pOrderNo) & "' and PartNo = '" & Trim(pPartNo) & "' and SupplierID = '" & Trim(pSupplierID) & "'"

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekDOBoxDetail = ds.Tables(0).Rows(0)("boxno")
        Else
            CekDOBoxDetail = ""
        End If
    End Function

    Public Shared Function CekRecEX(ByVal Tmp As clsTmp) As Boolean
        Dim name As String = ""
        sql = "select isnull(suratjalanno,'') as SJ From ReceiveForwarder_Master " & vbCrLf & _
              " where suratjalanno = '" & Trim(Tmp.SuratJalanNo) & "' " & vbCrLf & _
              " and pono = '" & Trim(Tmp.PONo) & "' and AffiliateID = '" & Trim(Tmp.AffiliateID) & "' " & vbCrLf & _
              " and SupplierID = '" & Trim(Tmp.SupplierID) & "' " & vbCrLf & _
              " and ForwarderID = '" & Trim(Tmp.ForwarderID) & "'" & vbCrLf & _
              " and OrderNo = '" & Trim(Tmp.OrderNo) & "' "

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            If ds.Tables(0).Rows(0)("SJ") = "" Then CekRecEX = False Else CekRecEX = True
        Else
            CekRecEX = False
        End If
    End Function

    Public Shared Function CekPOEX(ByVal Tmp As clsTmp) As Boolean
        Dim name As String = ""
        sql = "select * From PO_Master_Export " & vbCrLf & _
              " where OrderNo1 = '" & Trim(Tmp.OrderNo) & "' " & vbCrLf & _
              " and AffiliateID = '" & Trim(Tmp.AffiliateID) & "' " & vbCrLf & _
              " and SupplierID = '" & Trim(Tmp.SupplierID) & "' " & vbCrLf

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOEX = True
        Else
            CekPOEX = False
        End If
    End Function

    Public Shared Function CekPOMonthlyStatus(ByVal tblName As String, ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pStatus As Integer) As Integer
        Dim name As String = ""

        sql = " SELECT * FROM " & tblName & " WITH(NOLOCK) WHERE " & vbCrLf & _
              " (FinalApprovalCls is not NULL or FinalApprovalCls = 0) AND " & vbCrLf & _
              " PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo1 = '" & Trim(pOrderNo) & "'"

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOMonthlyStatus = 1
        Else
            CekPOMonthlyStatus = 0
        End If
    End Function

    Public Shared Function CekExistsPart(ByVal tblName As String, ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pPartno As String) As Boolean
        Dim name As String = ""
        Dim i As Integer = 0

        sql = " SELECT * FROM " & tblName & " WITH(NOLOCK) WHERE " & vbCrLf & _
              " PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "'" & vbCrLf & _
              " and OrderNo1 = '" & Trim(pOrderNo) & "' and PartNo = '" & Trim(pPartno) & "'"

        ds = uf_GetDataSet(sql)

        For i = 0 To ds.Tables(0).Rows.Count
            If ds.Tables(0).Rows.Count > 0 Then
                CekExistsPart = True
                Exit For
            Else
                CekExistsPart = False
                Exit For
            End If
        Next
    End Function

    Public Shared Function CekPOMonthlyEXIST(ByVal tblName As String, ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pStatus As Integer) As Integer
        Dim name As String = ""


        '    'cek split
        'sql = " select jml = count(OrderNo1) from " & tblName & " with(nolock) " & vbCrLf
        '    If tblName = "dbo.PO_DetailUpload_Export" Then
        '        sql = sql + " WHERE PONo = '" & Trim(pPONo) & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND OrderNo1 = '" & Trim(pOrderNo) & "'"
        '    Else
        '        sql = sql + " WHERE PONo = '" & Trim(pPONo) & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND OrderNo1 = '" & Trim(pOrderNo) & "'"
        '    End If
        '    ds = uf_GetDataSet(sql)
        '    'cek split

        'If ds.Tables(0).Rows.Count > 0 Then
        '    Dim jmlOrder As Integer = 0
        '    jmlOrder = ds.Tables(0).Rows(0)("Jml")
        '    If jmlOrder = 0 Then
        '        pOrderNo = pOrderNo
        '    Else
        '        pOrderNo = pOrderNo + "-" + jmlOrder + 1
        '    End If
        'End If

        sql = "SELECT * FROM " & tblName & " WITH(NOLOCK) WHERE " & vbCrLf

        If tblName = "dbo.PO_Master_Export" Then
            sql = sql + " PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo1 = '" & Trim(pOrderNo) & "'"
        Else
            sql = sql + " PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo1 = '" & Trim(pOrderNo) & "'"
        End If

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOMonthlyEXIST = 1
        Else
            CekPOMonthlyEXIST = 0
        End If
    End Function

    Public Shared Function CekSplitPO(ByVal tblName As String, ByVal pPONo As String, ByVal pOrderNo As String, ByVal pAffiliateID As String) As Integer
        Dim name As String = ""

        sql = "Select x = Max(x) from ("
        sql = sql + "select x = isnull(right(Rtrim(Max(OrderNo1)),1),0) from PO_Master_Export with(Nolock) where pono = '" & Trim(pPONo) & "' and AffiliateID = '" & Trim(pAffiliateID) & "' and PONo <> OrderNo1" & vbCrLf
        sql = sql + "UNION ALL " & vbCrLf
        sql = sql + "select x = isnull(right(Rtrim(Max(OrderNo1)),1),0) from PO_Master_ExportRecoverySplit with(Nolock) where pono = '" & Trim(pPONo) & "' and AffiliateID = '" & Trim(pAffiliateID) & "' and PONo <> OrderNo1" & vbCrLf
        sql = sql + ")c"

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekSplitPO = ds.Tables(0).Rows(0)("x")
        Else
            CekSplitPO = 0
        End If
    End Function

    Public Shared Function CekPOMontlyApprover(ByVal pPONo As String, ByVal pPartNo As String, ByVal pAffiliateID As String) As Integer
        Dim name As String = ""

        sql = "SELECT PartNo FROM PO_Detail_Export WITH(NOLOCK) WHERE " & vbCrLf & _
              " OrderNo1 = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and PartNo in (" & Trim(pPartNo) & ")"

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOMontlyApprover = 1
        Else
            CekPOMontlyApprover = 0
        End If
    End Function

    Public Shared Function CekPOMontlyQty(ByVal pPONo As String, ByVal pPartNo As String, ByVal pAffiliateID As String) As Integer
        Dim name As String = ""

        sql = "SELECT Week1 FROM PO_Detail_Export WITH(NOLOCK) WHERE " & vbCrLf & _
              " OrderNo1 = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and PartNo = '" & Trim(pPartNo) & "'"

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOMontlyQty = ds.Tables(0).Rows(0)("Week1")
        Else
            CekPOMontlyQty = 0
        End If

    End Function

    Public Shared Function CekPOMontlyMOQ(ByVal pPartNo As String) As Integer
        Dim name As String = ""

        sql = "SELECT MOQ FROM MS_Parts WITH(NOLOCK) WHERE " & vbCrLf & _
              " PartNo = '" & Trim(pPartNo) & "'"

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOMontlyMOQ = ds.Tables(0).Rows(0)("MOQ")
        Else
            CekPOMontlyMOQ = 0
        End If
    End Function

    Public Shared Function CekDataDNSupplier(ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal PartNo As String, ByVal Box As String) As Integer
        Dim name As String = ""

        sql = "SELECT * FROM DOSupplier_DetailBox_Export ds WITH(NOLOCK) WHERE " & vbCrLf

        sql = sql + " PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo = '" & Trim(pOrderNo) & "'" & vbCrLf & _
                    " And PartNo = '" & Trim(PartNo) & "' and BoxNo = '" & Trim(Box) & "' and not exists " & vbCrLf & _
                    "(" & vbCrLf & _
                    "   select * from ReceiveForwarder_DetailBox a WITH(NOLOCK)" & vbCrLf & _
                    "   WHERE a.PartNo = ds.PartNo and (DS.BoxNo between a.Label1 and a.Label2) and StatusDefect = '1'" & vbCrLf & _
                    ")" & vbCrLf

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekDataDNSupplier = 1
        Else
            CekDataDNSupplier = 0
        End If
    End Function

    Public Shared Function CekDataDNremaining(ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal PartNo As String, ByVal Box As String) As Integer
        Dim name As String = ""

        sql = "SELECT * FROM DOSupplier_DetailBox_Export ds WITH(NOLOCK) WHERE " & vbCrLf

        sql = sql + " PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo = '" & Trim(pOrderNo) & "'" & vbCrLf & _
                    " And PartNo = '" & Trim(PartNo) & "' and BoxNo = '" & Trim(Box) & "' and exists " & vbCrLf & _
                    "(" & vbCrLf & _
                    "   select * from ReceiveForwarder_DetailBox a WITH(NOLOCK)" & vbCrLf & _
                    "   WHERE a.PartNo = ds.PartNo and (DS.BoxNo between a.Label1 and a.Label2) and StatusDefect = '1'" & vbCrLf & _
                    ")" & vbCrLf

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekDataDNremaining = 1
        Else
            CekDataDNremaining = 0
        End If
    End Function

    Public Shared Function CekBoxNoDNSupplier(ByVal PartNo As String, ByVal Box As String, ByVal OrderNo As String) As Integer
        Dim name As String = ""

        sql = "SELECT * FROM PrintLabelExport ds WITH(NOLOCK) WHERE " & vbCrLf

        sql = sql + " PartNo = '" & PartNo & "' AND LabelNo = '" & Box & "' AND OrderNo = '" & Trim(OrderNo) & "' /*And SuratJalanNo_fwd = ''*/" & vbCrLf

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekBoxNoDNSupplier = 1
        Else
            CekBoxNoDNSupplier = 0
        End If
    End Function

    Public Shared Function CekReceiveBox(ByVal SuratJalan As String, ByVal PartNo As String, ByVal PONo As String, ByVal Box As String) As Integer
        Dim name As String = ""

        sql = "Exec sp_SelectBoxNo '" & SuratJalan & "','" & PartNo & "','" & PONo & "','" & Box & "'" & vbCrLf

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekReceiveBox = 1
        Else
            CekReceiveBox = 0
        End If
    End Function

    Public Shared Function CekBoxNoDouble(ByVal Box As String) As Integer
        Dim name As String = ""

        sql = "SELECT * FROM DOSupplier_DetailBox_Export WITH(NOLOCK) WHERE " & vbCrLf

        sql = sql + " BoxNo = '" & Box & "' " & vbCrLf

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekBoxNoDouble = 1
        Else
            CekBoxNoDouble = 0
        End If
    End Function

    Public Shared Function CekDataDetailDNSupplier(ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal PartNo As String, ByVal Box As String, ByVal pSuratJalanNo As String) As Integer
        Dim name As String = ""

        sql = "SELECT AffiliateID, MAX(SeqNo) + 1 SeqNo FROM DOSupplier_Detail_Export ds WITH(NOLOCK) WHERE " & vbCrLf

        sql = sql + " PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo = '" & Trim(pOrderNo) & "'" & vbCrLf & _
                    " And PartNo = '" & Trim(PartNo) & "' and SuratJalanNo = '" & Trim(pSuratJalanNo) & "'" & vbCrLf & _
                    " GROUP BY AffiliateID"
        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekDataDetailDNSupplier = ds.Tables(0).Rows(0)("SeqNo")
        Else
            CekDataDetailDNSupplier = 1
        End If
    End Function

    Public Shared Function CekPOMontlyQtyBox(ByVal pPartNo As String) As Integer
        Dim name As String = ""

        sql = "SELECT QtyBox FROM MS_Parts WITH(NOLOCK) WHERE " & vbCrLf & _
              " PartNo = '" & Trim(pPartNo) & "'"

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOMontlyQtyBox = ds.Tables(0).Rows(0)("QtyBox")
        Else
            CekPOMontlyQtyBox = 0
        End If
    End Function

    Public Shared Function CekPOMontly(ByVal pPONo As String, ByVal pPartNo As String, ByVal pAffiliateID As String) As Integer
        Dim name As String = ""

        sql = "SELECT COUNT(PartNo) PartNo FROM PO_Detail_Export WITH(NOLOCK) WHERE " & vbCrLf & _
              " OrderNo1 = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and PartNo in (" & Trim(pPartNo) & ")"

        ds = uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            CekPOMontly = ds.Tables(0).Rows(0)("PartNo")
        Else
            CekPOMontly = 0
        End If
    End Function

    Public Shared Function insertMasterPOMonthlyUpload(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal tblName As String, ByVal templateCode As String) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            Dim tempOrderNo As String '= pTmp.OrderNo + "-1"
            SQLCon.Open()

            tempOrderNo = Trim(pTmp.OrderNo)

            sql = " UPDATE " & tblName & " " & vbCrLf & _
                  " SET ETDVendor1 = '" & Format(pTmp.ETDSplit, "yyyy-MM-dd") & "', " & vbCrLf & _
                  " Remarks = '" & Trim(pTmp.Remarks) & "', " & vbCrLf & _
                  " UpdateDate = GETDATE(), " & vbCrLf & _
                  " UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf

            If tblName = "dbo.PO_DetailUpload_Export" Then
                sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND OrderNo1 = '" & Trim(pTmp.OrderNo) & "'"
            Else
                sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND OrderNo1 = '" & Trim(pTmp.OrderNo) & "'"
            End If

            'If pTmp.QtySplit > 0 Then
            '    sql = sql + " AND ETDVendor1 = '" & pTmp.ETDSplit & "'"
            'End If

            SQLCom.CommandText = sql
            Dim i As Integer = SQLCom.ExecuteNonQuery

            If i = 0 Then

                'tempOrderNo = pTmp.OrderNo

                sql = " INSERT INTO " & tblName & " " & vbCrLf
                If templateCode = "POEM" Then
                    sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "' ,'" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                                "           '" & Trim(pTmp.ForwarderID) & "' , '" & tempOrderNo & "', '" & Format(pTmp.ETDSplit, "yyyy-MM-dd") & "' ,  " & vbCrLf
                Else
                    sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "' ,'" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                                "           '" & Trim(pTmp.ForwarderID) & "' , '" & tempOrderNo & "', '" & Format(pTmp.ETDSplit, "yyyy-MM-dd") & "' ,  " & vbCrLf
                End If

                sql = sql + "           '" & (pTmp.Remarks) & "' , " & vbCrLf & _
                            "           GETDATE() ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.EntryUser) & "' ," & vbCrLf & _
                            "           GETDATE() , " & vbCrLf & _
                            "           '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
                            "         ) "


                SQLCom.CommandText = sql
                i = SQLCom.ExecuteNonQuery()

            End If

            Return i
        End Using

    End Function

    Public Shared Function insertMasterPOMonthlyAfterUpload(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal tblName As String, ByVal templateCode As String, ByVal pweek As Integer) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)

            SQLCon.Open()
            Dim tempOrderNo As String

            tempOrderNo = pTmp.OrderNo

            sql = " INSERT INTO " & tblName & " " & vbCrLf

            If templateCode = "POEM" Then
                sql = sql + " (POno,AffiliateID,SupplierID, ForwarderID, Period,  " & vbCrLf & _
                      " CommercialCls, EmergencyCls, Shipcls, errorStatus,OrderNo1, ETDVendor1, ETDPort1, ETAPort1, " & vbCrLf & _
                      " ETAFactory1, UploadDate, UploadUser, EntryDate, EntryUser, Excelcls, PasiSendToSupplierDate, PasiSendToSupplierUser, updatedate, updateUser) " & vbCrLf & _
                      "  "
                sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "' ,'" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.ForwarderID) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.Period) & "', " & vbCrLf & _
                            "           '1','M', (select isnull(ShipCls,'') from PO_Master_Export where PoNo = '" & Trim(pTmp.PONo) & "' and AffiliateID = '" & Trim(pTmp.AffiliateID) & "' and supplierID = '" & Trim(pTmp.SupplierID) & "' and OrderNo1 = '" & Trim(pTmp.PONo) & "'), 'OK', " & vbCrLf & _
                            "           '" & tempOrderNo & "', '" & Format(pTmp.ETDSplit, "yyyy-MM-dd") & "' ,  " & vbCrLf & _
                            "           (SELECT ETDPort FROM ms_ETD_Export WHERE period = '" & Trim(pTmp.Period) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND supplierID = '" & Trim(pTmp.SupplierID) & "' AND Week = '" & pweek & "'), " & vbCrLf & _
                            "           (SELECT ETAPort FROM ms_ETD_Export WHERE period = '" & Trim(pTmp.Period) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND supplierID = '" & Trim(pTmp.SupplierID) & "' AND Week = '" & pweek & "'), " & vbCrLf & _
                            "           (SELECT EtaFactory FROM ms_ETD_Export WHERE period = '" & Trim(pTmp.Period) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND supplierID = '" & Trim(pTmp.SupplierID) & "' AND Week = '" & pweek & "'), " & vbCrLf & _
                            "           getdate(), " & vbCrLf & _
                            "           '" & Trim(pTmp.UpdateUser) & "', " & vbCrLf & _
                            "           GETDATE() , " & vbCrLf & _
                            "           '" & Trim(pTmp.UpdateUser) & "', '2', " & vbCrLf & _
                            "           getdate(), 'export', getdate(), 'export' " & vbCrLf & _
                            "         ) " & vbCrLf
            Else
                sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "' ,'" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.ForwarderID) & "' , '" & tempOrderNo & "', '" & Format(pTmp.ETDVendor1, "yyyy-MM-dd") & "' ,  " & vbCrLf
                sql = sql + "           '" & (pTmp.Remarks) & "' , " & vbCrLf & _
                        "           GETDATE() ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.EntryUser) & "' ," & vbCrLf & _
                        "           GETDATE() , " & vbCrLf & _
                        "           '" & Trim(pTmp.UpdateUser) & "', " & vbCrLf & _
                        "           getdate(), 'export', getdate(), 'export' " & vbCrLf & _
                        "         ) "
            End If

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()

            Return i
        End Using

    End Function

    Public Shared Function UpdatePOMonthlyMasterUpload(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal tblName As String, ByVal templateCode As String) As Integer

        sql = " UPDATE " & tblName & " " & vbCrLf & _
            "   SET Remarks = '" & Trim(pTmp.Remarks) & "', " & vbCrLf & _
            "   ETDVendor1 = '" & Format(pTmp.ETDVendor1, "yyyy-MM-dd") & "', " & vbCrLf & _
            "   UpdateDate = GETDATE(), UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf

        If templateCode = "POEM" Then
            sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND OrderNo1 = '" & Trim(pTmp.OrderNo) & "'"
        Else
            sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND OrderNo1 = '" & Trim(pTmp.OrderNo) & "'"
        End If

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        Return i
    End Function

    Public Shared Function UpdatePOMonthlyMaster(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal pStatus As String, ByVal tblName As String) As Integer
        '--
        sql = " UPDATE " & tblName & " " & vbCrLf & _
            "   SET " & vbCrLf
        'If pStatus = "0" Then
        sql = sql + " SupplierApproveDate = GETDATE(), " & vbCrLf & _
                    " SupplierApproveUser = '" & Trim(pTmp.SupplierApproveUser) & "' " & vbCrLf
        'End If

        If tblName = "dbo.PO_Master_Export" Then
            sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND OrderNo1 = '" & Trim(pTmp.OrderNo) & "'"
        Else
            sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND OrderNo1 = '" & Trim(pTmp.OrderNo) & "'"
        End If

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        Return i
    End Function

    Public Shared Function DeletePODetail0(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal pStatus As String, ByVal tblName As String) As Integer
        sql = "Delete PO_Detail_Export "
        sql = sql + " WHERE week1 = 0"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        Return i
    End Function

    Public Shared Function insertDetailPOMonthlyUpload(ByVal pTmp As clsTmp, _
            ByVal PONo As String, ByVal AffiliateID As String, _
            ByVal SupplierID As String, ByVal ForwarderID As String, _
            ByVal PartNo As String, ByVal Week1 As String, ByVal Week1Old As String, _
            ByVal Week2 As String, ByVal Week2Old As String, _
            ByVal Week3 As String, ByVal Week3Old As String, _
            ByVal Week4 As String, ByVal Week4Old As String, _
            ByVal Week5 As String, ByVal Week5Old As String, _
            ByVal TotalPOQty As String, ByVal TotalPOQtyOld As String, _
            ByVal PreviousForecast As String, ByVal Forecast1 As String, _
            ByVal Forecast2 As String, ByVal Forecast3 As String, _
            ByVal tblName As String, ByVal templateCode As String, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            If templateCode = "POEM" Then

                sql = " UPDATE PO_DetailUpload_Export " & vbCrLf & _
                      " SET TotalPOQty = '" & Trim(pTmp.POQty) & "', " & vbCrLf & _
                      "     TotalPOQtyOld = '" & Trim(pTmp.POQtyOld) & "', " & vbCrLf & _
                      "     UpdateDate = GETDATE(), " & vbCrLf & _
                      "     UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf

            Else

                sql = " UPDATE " & tblName & " " & vbCrLf & _
                     " SET TotalPOQty = '" & Trim(pTmp.POQty) & "', " & vbCrLf & _
                     "     TotalPOQtyOld = '" & Trim(pTmp.POQtyOld) & "', " & vbCrLf & _
                     "     UpdateDate = GETDATE(), " & vbCrLf & _
                     "     UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf
            End If

            Dim ls_tempPONO As String = ""

            ls_tempPONO = Trim(pTmp.OrderNo)


            If tblName = "dbo.PO_DetailUpload_Export" Then
                sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND OrderNo1 = '" & Trim(ls_tempPONO) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "'"
            Else
                sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND OrderNo1 = '" & Trim(ls_tempPONO) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "'"
            End If

            ''update PO DETAIL DIKURANGIN QTY SPLIT
            'If statusSplit = True Then
            '    sql = sql + " UPDATE PO_Detail_Export " & vbCrLf & _
            '          " SET Week1 = Week1 - " & Trim(pTmp.POQty) & ", " & vbCrLf & _
            '          "     UpdateDate = GETDATE(), " & vbCrLf & _
            '          "     UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf
            '    sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND OrderNo1 = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "'"
            'End If
            ''update PO DETAIL DIKURANGIN QTY SPLIT


            sqlCom.CommandText = sql

            Dim i As Integer = sqlCom.ExecuteNonQuery

            If i = 0 Then
                sql = " INSERT INTO " & tblName & "  " & vbCrLf

                If templateCode = "POEM" Then
                    sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "','" & Trim(pTmp.AffiliateID) & "' , " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                                "           '" & Trim(pTmp.ForwarderID) & "' , '" & Trim(ls_tempPONO) & "' , '" & Trim(pTmp.PartNo) & "', " & vbCrLf & _
                                "            '" & Trim(pTmp.POQty) & "', '" & Trim(pTmp.POQtyOld) & "'," & vbCrLf & _
                                "            '" & Trim(pTmp.POQty) & "', '" & Trim(pTmp.POQtyOld) & "'," & vbCrLf & _
                                "            GETDATE(), '" & Trim(pTmp.EntryUser) & "', GETDATE(), '" & Trim(pTmp.UpdateUser) & "')" & vbCrLf

                Else
                    sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "','" & Trim(pTmp.AffiliateID) & "' , " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                                 "           '" & Trim(pTmp.ForwarderID) & "' , '" & Trim(ls_tempPONO) & "' , '" & Trim(pTmp.PartNo) & "', " & vbCrLf & _
                                "            '" & Trim(pTmp.POQty) & "', '" & Trim(pTmp.POQtyOld) & "'," & vbCrLf & _
                                "            '" & Trim(pTmp.POQty) & "', '" & Trim(pTmp.POQtyOld) & "'," & vbCrLf & _
                                "            GETDATE(), '" & Trim(pTmp.EntryUser) & "', GETDATE(), '" & Trim(pTmp.UpdateUser) & "')" & vbCrLf

                End If

                sqlCom.CommandText = sql
                i = sqlCom.ExecuteNonQuery()

            End If

            Return i
        End Using

    End Function

    Public Shared Function insertDetailPOMonthlyAfterUpload(ByVal pTmp As clsTmp, _
            ByVal PONo As String, ByVal AffiliateID As String, _
            ByVal SupplierID As String, ByVal ForwarderID As String, _
            ByVal PartNo As String, ByVal Week1 As String, ByVal Week1Old As String, _
            ByVal Week2 As String, ByVal Week2Old As String, _
            ByVal Week3 As String, ByVal Week3Old As String, _
            ByVal Week4 As String, ByVal Week4Old As String, _
            ByVal Week5 As String, ByVal Week5Old As String, _
            ByVal TotalPOQty As String, ByVal TotalPOQtyOld As String, _
            ByVal PreviousForecast As String, ByVal Forecast1 As String, _
            ByVal Forecast2 As String, ByVal Forecast3 As String, _
            ByVal tblName As String, ByVal templateCode As String, ByVal sqlCom As SqlCommand, ByVal statusSplit As Boolean) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            If statusSplit = True Then
                If templateCode = "POEM" Then

                    sql = " UPDATE " & tblName & " " & vbCrLf & _
                          " SET --TotalPOQty = '" & Trim(pTmp.POQty) & "', " & vbCrLf & _
                          "     --TotalPOQtyOld = '" & Trim(pTmp.POQtyOld) & "', " & vbCrLf & _
                          "     UpdateDate = GETDATE(), " & vbCrLf & _
                          "     UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf

                Else

                    sql = " UPDATE " & tblName & " " & vbCrLf & _
                         " SET --TotalPOQty = '" & Trim(pTmp.POQty) & "', " & vbCrLf & _
                         "     UpdateDate = GETDATE(), " & vbCrLf & _
                         "     UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf
                End If

                Dim ls_tempPONO As String = ""
                ls_tempPONO = Trim(pTmp.OrderNo)

                If tblName = "dbo.PO_Detail_Export" Then
                    sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND OrderNo1 = '" & Trim(ls_tempPONO) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "'" & vbCrLf
                Else
                    sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND OrderNo1 = '" & Trim(ls_tempPONO) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "'" & vbCrLf
                End If

                sqlCom.CommandText = sql

                Dim i As Integer = sqlCom.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO " & tblName & "  " & vbCrLf

                    If templateCode = "POEM" Then
                        sql = sql + " (POno,AffiliateID,SupplierID, ForwarderID,OrderNo1,PartNo, " & vbCrLf & _
                              " Week1, TotalPOQty, PreviousForecast, Forecast1, Forecast2, Forecast3, Variance, VariancePercentage, EntryDate, EntryUser, POMOQ, POQtyBox) " & vbCrLf & _
                              "  "
                        sql = sql + " VALUES  ( '" & Trim(pTmp.PONo) & "','" & Trim(pTmp.AffiliateID) & "' , " & vbCrLf & _
                                    "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                                    "           '" & Trim(pTmp.ForwarderID) & "' , '" & Trim(ls_tempPONO) & "' , '" & Trim(pTmp.PartNo) & "', " & vbCrLf & _
                                    "            '" & Trim(pTmp.POQty) & "', '" & Trim(pTmp.POQty) & "'," & vbCrLf & _
                                    "            '" & Trim(pTmp.POQty) & "', '" & Trim(pTmp.POQty) & "', '" & Trim(pTmp.POQty) & "', '" & Trim(pTmp.POQty) & "', " & vbCrLf & _
                                    "            0,0, GETDATE(), '" & Trim(pTmp.EntryUser) & "', " & vbCrLf & _
                                    "            '" & uf_GetMOQ(0, "", pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID) & "', '" & uf_GetQtybox(0, "", pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID) & "' ) " & vbCrLf

                    Else

                    End If

                    If templateCode = "POEM" Then 'update jika ada split
                        If Trim(pTmp.PONo) <> Trim(ls_tempPONO) Then

                            sql = sql + " UPDATE " & tblName & " " & vbCrLf & _
                                        " SET TotalPOQty = isnull((select Sum(Week1) from PO_DetailUpload_Export with(Nolock) " & vbCrLf & _
                                        "                       where SupplierID = '" & Trim(pTmp.SupplierID) & "' and pono = '" & Trim(pTmp.PONo) & "' and OrderNo1 = '" & Trim(pTmp.PONo) & "' and PartNo = '" & Trim(pTmp.PartNo) & "' " & vbCrLf & _
                                        "                       Group By PONO,AffiliateID,SupplierID,ForwarDerID,PartNo),0) " & vbCrLf
                            sql = sql + " ,Week1 = isnull((select Sum(Week1) from PO_DetailUpload_Export with(Nolock) " & vbCrLf & _
                                        "                       where SupplierID = '" & Trim(pTmp.SupplierID) & "' and pono = '" & Trim(pTmp.PONo) & "' and OrderNo1 = '" & Trim(pTmp.PONo) & "' and PartNo = '" & Trim(pTmp.PartNo) & "' " & vbCrLf & _
                                        "                       Group By PONO,AffiliateID,SupplierID,ForwarDerID,PartNo),0) " & vbCrLf
                            sql = sql + " WHERE PONo = '" & Trim(pTmp.PONo) & "' AND OrderNo1 = '" & Trim(pTmp.PONo) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "'"


                        End If
                    End If

                    sqlCom.CommandText = sql
                    i = sqlCom.ExecuteNonQuery()

                End If

                Return i
            End If
        End Using

    End Function

    Public Shared Function UpdateRECEX(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand, ByVal Box As String) As Integer

        sql = " Update ReceiveForwarder_Master  " & vbCrLf & _
              " SET SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' " & vbCrLf & _
              " Where SuratJalanNo In " & vbCrLf & _
              " (SELECT Top 1 SuratJalanNo FROM ReceiveForwarder_DetailBox wHERE PONo = '" & Trim(pTmp.PONo) & "' And PartNo = '" & Trim(pTmp.PartNo) & "' And '" & Box & "' Between Label1 And Label2) " & vbCrLf & _
              " Update ReceiveForwarder_Detail  " & vbCrLf & _
              " SET SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' " & vbCrLf & _
              " Where SuratJalanNo In " & vbCrLf & _
              " (SELECT Top 1 SuratJalanNo FROM ReceiveForwarder_DetailBox wHERE PONo = '" & Trim(pTmp.PONo) & "' And PartNo = '" & Trim(pTmp.PartNo) & "' And '" & Box & "' Between Label1 And Label2) " & vbCrLf & _
              " Update ReceiveForwarder_DetailBox  " & vbCrLf & _
              " SET SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' " & vbCrLf & _
              " Where SuratJalanNo In " & vbCrLf & _
              " (SELECT Top 1 SuratJalanNo FROM ReceiveForwarder_DetailBox wHERE PONo = '" & Trim(pTmp.PONo) & "' And PartNo = '" & Trim(pTmp.PartNo) & "' And '" & Box & "' Between Label1 And Label2) " & vbCrLf & _
              " UPDATE PrintLabelExport SET SuratJalanNo_FWD = '" & Trim(pTmp.SuratJalanNo) & "' Where PONo = '" & Trim(pTmp.PONo) & "' And PartNo = '" & Trim(pTmp.PartNo) & "' And LabelNo = '" & Box & "'" & vbCrLf

        sql = sql + " Update ShippingInstruction_Detail Set SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' Where AffiliateID = '" & Trim(pTmp.AffiliateID) & "' And SupplierID = '" & Trim(pTmp.SupplierID) & "' And PartNo = '" & Trim(pTmp.PartNo) & "' And OrderNo = '" & Trim(pTmp.PONo) & "' And LEFT(BoxNo,9) = '" & Trim(Box) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery
        Return i
    End Function

    Public Shared Function insertMasterDOEX(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer
        sql = " UPDATE dbo.DOSUpplier_Master_Export " & vbCrLf & _
            " SET UpdateDate = '" & Trim(pTmp.UpdateDate) & "', " & vbCrLf

        sql = sql + " 	UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
                    " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
                    " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND OrderNo = '" & Trim(pTmp.OrderNo) & "'" & vbCrLf & _
                    " AND PONo = '" & Trim(pTmp.PONo) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.DOSUpplier_Master_Export " & vbCrLf & _
                  " (suratjalanno, supplierID, AffiliateID, PONo, OrderNo, EntryDate, EntryUser, UpdateDate, UpdateUser, ExcelCls, DeliveryDate, CommercialCls) " & vbCrLf & _
                        " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.PONo) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.OrderNo) & "' ," & vbCrLf

            sql = sql + "           GETDATE() ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.EntryUser) & "' ,  " & vbCrLf & _
                        "           GETDATE() , " & vbCrLf & _
                        "           '" & Trim(pTmp.UpdateUser) & "', '1', Getdate(), '" & pTmp.CommercialCls & "' " & vbCrLf & _
                        "         ) "

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()
        End If
        Return i
    End Function

    Public Shared Function updateExcelDOEX(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer
        sql = " UPDATE dbo.DOSUpplier_Master_Export " & vbCrLf & _
                    " SET UpdateDate = '" & Trim(pTmp.UpdateDate) & "', " & vbCrLf

        sql = sql + " 	UpdateUser = '" & Trim(pTmp.UpdateUser) & "', " & vbCrLf & _
                    "   ExcelCls = '0' " & vbCrLf & _
                    " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
                    " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND OrderNo = '" & Trim(pTmp.OrderNo) & "'" & vbCrLf & _
                    " AND PONo = '" & Trim(pTmp.PONo) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery
        Return i
    End Function

    Public Shared Function insertDetailDOEX(ByVal pTmp As clsTmp, ByVal pSeqNo As Integer, ByVal SQLCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " UPDATE dbo.DOSUpplier_Detail_Export " & vbCrLf & _
                        " SET DOQty = " & Trim(pTmp.DOQty) & " " & vbCrLf & _
                        " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                        " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' " & vbCrLf & _
                        " AND PONo = '" & Trim(pTmp.PONo) & "' and SeqNo = '" & pSeqNo & "'"

            'Dim SQLCom As New SqlCommand(sql, SQLCon)

            'Dim i As Integer = SQLCom.ExecuteNonQuery
            SQLCom.CommandText = sql
            Dim i As Integer = SQLCom.ExecuteNonQuery

            If i = 0 Then
                sql = " INSERT INTO dbo.DOSUpplier_Detail_Export " & vbCrLf & _
                            " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.PONo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.PartNo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.OrderNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pSeqNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.DOQty) & "' ," & vbCrLf & _
                            "           '" & uf_GetMOQ(2, pTmp.PONo, pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID) & "' ," & vbCrLf & _
                            "           '" & uf_GetQtybox(2, pTmp.PONo, pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID) & "', " & vbCrLf & _
                            "           '" & uf_GetPriceSupplier(pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID) & "' " & vbCrLf & _
                            "       ) "

                SQLCom.CommandText = sql
                i = SQLCom.ExecuteNonQuery()
            End If

            Return i
        End Using

    End Function

    Public Shared Function insertDetailDOBoxEX(ByVal pTmp As clsTmp, ByVal pBoxNo As String, ByVal pSeqNo As Integer, ByVal SqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " UPDATE dbo.DOSUpplier_DetailBox_Export " & vbCrLf & _
                        " SET BoxNo = '" & Trim(pBoxNo) & "' " & vbCrLf & _
                        " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                        " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' " & vbCrLf & _
                        " AND PONo = '" & Trim(pTmp.PONo) & "'" & vbCrLf & _
                        " AND BoxNo = '" & Trim(pBoxNo) & "' and SeqNo = '" & pSeqNo & "'"

            SqlCom.CommandText = sql
            Dim i As Integer = SqlCom.ExecuteNonQuery

            If i = 0 Then
                sql = " INSERT INTO dbo.DOSUpplier_DetailBox_Export " & vbCrLf & _
                            " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.PONo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.PartNo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.OrderNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pBoxNo) & "'," & vbCrLf & _
                            "           '" & Trim(pSeqNo) & "'" & vbCrLf & _
                            "       ) "

                SqlCom.CommandText = sql
                i = SqlCom.ExecuteNonQuery()
            End If

            Return i
        End Using

    End Function

    Public Shared Function insertMasterRecevingEX(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer

        sql = " UPDATE dbo.receiveForwarder_master " & vbCrLf & _
            " SET UpdateDate = '" & Trim(pTmp.UpdateDate) & "', " & vbCrLf

        sql = sql + " 	UpdateUser = '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
                    " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
                    " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND OrderNo = '" & Trim(pTmp.OrderNo) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.receiveForwarder_master " & vbCrLf & _
                  " (suratjalanno, supplierID, AffiliateID, PONo, ForwarderID, OrderNo, ExcelCls, ReceiveDate, ReceiveBy,EntryDate, EntryUser, UpdateDate, UpdateUser) " & vbCrLf & _
                        " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.SupplierID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.AffiliateID) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.PONo) & "' ," & vbCrLf & _
                        "           '" & Trim(pTmp.ForwarderID) & "', " & vbCrLf & _
                        "           '" & Trim(pTmp.OrderNo) & "', " & vbCrLf & _
                        "           '0', " & vbCrLf
            sql = sql + "           GETDATE() ,  " & vbCrLf & _
                        "           '" & Trim(pTmp.EntryUser) & "' ,  " & vbCrLf & _
                        "           GETDATE() , " & vbCrLf & _
                        "           '" & Trim(pTmp.UpdateUser) & "', " & vbCrLf & _
                        "           GETDATE() , " & vbCrLf & _
                        "           '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
                        "         ) "

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()
        End If
        Return i
    End Function

    Public Shared Function updateExcelMasterRecevingEX(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer

        sql = " UPDATE dbo.receiveForwarder_master " & vbCrLf & _
            " SET ExcelCls = '1' " & vbCrLf & _
            " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
            " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND OrderNo = '" & Trim(pTmp.OrderNo) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery
        Return i
    End Function

    Public Shared Function insertDetailReceivingEX(ByVal pTmp As clsTmp, ByVal xstatus As String, ByVal SqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " UPDATE dbo.receiveForwarder_Detail " & vbCrLf & _
                  " SET GoodRecQty = GoodRecQty + " & Trim(pTmp.GoodRecQty) & ", " & vbCrLf & _
                  " DefectRecQty = DefectRecQty + " & Trim(pTmp.DefectRecQty) & " " & vbCrLf & _
                  " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' "
            'sql = sql + " UPDATE dbo.receiveForwarder_Detail " & vbCrLf & _
            '      " SET OrderNo = '" & Trim(pTmp.OrderNo) & "'" & vbCrLf & _
            '      " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
            '      " AND OrderNo = '" & Trim(pTmp.PONo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' "
            SqlCom.CommandText = sql

            Dim i As Integer = SqlCom.ExecuteNonQuery

            If i = 0 Then
                sql = " INSERT INTO dbo.receiveForwarder_Detail " & vbCrLf & _
                            " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.PONo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.PartNo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pTmp.OrderNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.GoodRecQty) & "',  " & vbCrLf & _
                            "           '" & Trim(pTmp.DefectRecQty) & "'  " & vbCrLf & _
                            "       ) "

                SqlCom.CommandText = sql
                i = SqlCom.ExecuteNonQuery()
            End If

            Return i
        End Using

    End Function

    Public Shared Function insertDetailReceivingEX_BOX(ByVal pTmp As clsTmp, ByVal L1 As String, ByVal L2 As String, ByVal xstatus As String, ByVal totBox As Long, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            Dim Label1 As String = ""
            Dim Label2 As String = ""


            sql = "SELECT Label1,Label2 FROM dbo.receiveForwarder_DetailBox WITH(NOLOCK) WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' and (Label1 = '" & Trim(L1) & "' OR Label2 = '" & Trim(L2) & "') and StatusDefect <>'1'"
            ds = uf_GetDataSet(sql)

            If ds.Tables(0).Rows.Count > 0 Then
                Label1 = ds.Tables(0).Rows(0)("Label1")
                Label2 = ds.Tables(0).Rows(0)("Label2")
                insertDetailReceivingEX_BOX = 0
            Else
                Label1 = ""
                Label2 = ""

                If xstatus = "G" Then xstatus = "0" Else xstatus = "1"
                sql = " UPDATE dbo.receiveForwarder_DetailBox " & vbCrLf & _
                      " SET SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "', label1 = '" & Trim(L1) & "', " & vbCrLf & _
                      " Label2 = '" & Trim(L2) & "', " & vbCrLf & _
                      " StatusDefect = '" & Trim(xstatus) & "', " & vbCrLf & _
                      " Box = " & totBox & " " & vbCrLf & _
                      " WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                      " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' --and statusDefect = '" & Trim(xstatus) & "' " & vbCrLf & _
                      " AND Label1 = '" & Trim(L1) & "'" & vbCrLf & _
                      " AND Label2 = '" & Trim(L2) & "'" & vbCrLf
                'sql = sql + " UPDATE dbo.receiveForwarder_DetailBox " & vbCrLf & _
                '      " SET SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "', label1 = '" & Trim(L1) & "', " & vbCrLf & _
                '      " Label2 = '" & Trim(L2) & "', " & vbCrLf & _
                '      " StatusDefect = '" & Trim(xstatus) & "', " & vbCrLf & _
                '      " Box = " & totBox & ", " & vbCrLf & _
                '      " OrderNo = '" & Trim(pTmp.OrderNo) & "'" & vbCrLf & _
                '      " WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                '      " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' --and statusDefect = '" & Trim(xstatus) & "' " & vbCrLf & _
                '      " AND Label1 = '" & Trim(L1) & "'" & vbCrLf & _
                '      " AND Label2 = '" & Trim(L2) & "'" & vbCrLf
                sqlCom.CommandText = sql

                Dim i As Integer = sqlCom.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO dbo.receiveForwarder_DetailBox " & vbCrLf & _
                                " VALUES  ( '" & Trim(pTmp.SuratJalanNo) & "' , " & vbCrLf & _
                                "           '" & Trim(pTmp.SupplierID) & "' , " & vbCrLf & _
                                "           '" & Trim(pTmp.AffiliateID) & "' ,  " & vbCrLf & _
                                "           '" & Trim(pTmp.PONo) & "' ,  " & vbCrLf & _
                                "           '" & Trim(pTmp.OrderNo) & "' ,  " & vbCrLf & _
                                "           '" & Trim(pTmp.PartNo) & "' , " & vbCrLf & _
                                "           '" & Trim(L1) & "',  " & vbCrLf & _
                                "           '" & Trim(L2) & "',  " & vbCrLf & _
                                "           " & CDbl(totBox) & ",  " & vbCrLf & _
                                "           '" & Trim(xstatus) & "',  " & vbCrLf & _
                                "           NULL  " & vbCrLf & _
                                "       ) "

                    sqlCom.CommandText = sql
                    i = sqlCom.ExecuteNonQuery()
                End If

                Return i
            End If
        End Using

    End Function

    Public Shared Function CekBox(ByVal pTmp As clsTmp, ByVal L1 As String, ByVal L2 As String, ByVal xstatus As String, ByVal totBox As Long, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            Dim Label1 As String = ""
            Dim Label2 As String = ""


            sql = "select distinct BoxNo as box from DOSupplier_DetailBox_Export WITH(NOLOCK) WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' and BoxNo = '" & Trim(L1) & "' "
            sql = sql + "UNION ALL select distinct BoxNo as box from DOSupplier_DetailBox_Export WITH(NOLOCK) WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' and BoxNo = '" & Trim(L2) & "' "
            ds = uf_GetDataSet(sql)

            If ds.Tables(0).Rows.Count = 2 Then
                CekBox = 2
            Else
                CekBox = 0
            End If
        End Using

    End Function

    Public Shared Function CekBox2(ByVal pTmp As clsTmp, ByVal L1 As String, ByVal L2 As String, ByVal xstatus As String, ByVal totBox As Long, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            Dim Label1 As String = ""
            Dim Label2 As String = ""


            'sql = "select distinct BoxNo as box from DOSupplier_DetailBox_Export WITH(NOLOCK) WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
            '      " AND OrderNo = '" & Trim(pTmp.OrderNo1) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' and BoxNo = '" & Trim(L1) & "' "
            'sql = sql + "UNION ALL select distinct BoxNo as box from DOSupplier_DetailBox_Export WITH(NOLOCK) WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
            '      " AND OrderNo = '" & Trim(pTmp.OrderNo1) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' and BoxNo = '" & Trim(L2) & "' "
            sql = "select * from PrintLabelExport WITH(NOLOCK)" & vbCrLf & _
                  "where PONo = '" & Trim(pTmp.PONo) & "' and PartNo = '" & Trim(pTmp.PartNo) & "' and LabelNo = '" & Trim(L1) & "' and SuratJalanNo_FWD <> '' and statusDefect <> '1'"
            ds = uf_GetDataSet(sql)

            If ds.Tables(0).Rows.Count = 1 Then
                CekBox2 = 1
            Else
                CekBox2 = 0
            End If
        End Using

    End Function

    Public Shared Function CekBoxDNSupplier(ByVal pTmp As clsTmp, ByVal L1 As String, ByVal L2 As String, ByVal xstatus As String, ByVal totBox As Long, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            Dim Label1 As String = ""
            Dim Label2 As String = ""


            sql = "SELECT * FROM DOSupplier_DetailBox_Export ds WITH(NOLOCK) WHERE " & vbCrLf

            sql = sql + " PONo = '" & Trim(pTmp.PONo) & "' and AffiliateID ='" & Trim(pTmp.AffiliateID) & "' and SupplierID ='" & Trim(pTmp.SupplierID) & "' and OrderNo = '" & Trim(pTmp.OrderNo) & "'" & vbCrLf & _
                        " And PartNo = '" & Trim(pTmp.PartNo) & "' and BoxNo = '" & Trim(L1) & "' " & vbCrLf
            ds = uf_GetDataSet(sql)

            If ds.Tables(0).Rows.Count > 0 Then
                CekBoxDNSupplier = 1
            Else
                CekBoxDNSupplier = 0
            End If
        End Using

    End Function

    Public Shared Function getAffiliateID(ByVal pConsgineeCode As String) As String

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            Dim Label1 As String = ""
            Dim Label2 As String = ""

            sql = "select AffiliateID from MS_Affiliate WITH(NOLOCK)" & vbCrLf & _
                  "where ConsigneeCode = '" & pConsgineeCode & "'"
            ds = uf_GetDataSet(sql)

            If ds.Tables(0).Rows.Count > 0 Then
                getAffiliateID = ds.Tables(0).Rows(0)("AffiliateID")
            Else
                getAffiliateID = ""
            End If
        End Using

    End Function
    Public Shared Function CekData(ByVal pTmp As clsTmp, ByVal pTable As String) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            Dim Label1 As String = ""
            Dim Label2 As String = ""


            sql = "select * from " & Trim(pTable) & " WITH(NOLOCK) WHERE SuratjalanNo = '" & Trim(pTmp.SuratJalanNo.Replace(vbNewLine, "")) & "' and SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' "
            ds = uf_GetDataSet(sql)

            If ds.Tables(0).Rows.Count > 0 Then
                CekData = 1
            Else
                CekData = 0
            End If
        End Using

    End Function
    Public Shared Function CekData2(ByVal pTmp As clsTmp, ByVal pTable As String) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()
            Dim Label1 As String = ""
            Dim Label2 As String = ""


            sql = "select * from " & Trim(pTable) & " WITH(NOLOCK) WHERE SuratjalanNo = '" & Trim(pTmp.SuratJalanNo) & "' and SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' "
            ds = uf_GetDataSet(sql)

            If ds.Tables(0).Rows.Count > 0 Then
                CekData2 = 1
            Else
                CekData2 = 0
            End If
        End Using

    End Function

    Public Shared Function DeleteRecEX(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer

        sql = " Delete dbo.receiveForwarder_master " & vbCrLf & _
              " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
              " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND OrderNo = '" & Trim(pTmp.OrderNo) & "'"
        sql = sql + " Delete dbo.receiveForwarder_master " & vbCrLf & _
                    " WHERE SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' AND SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
                    " AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND OrderNo = '" & Trim(pTmp.OrderNo) & "'"

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery
        Return i
    End Function

    Public Shared Function UpdateLabelPrint_RecEX(ByVal pTmp As clsTmp, ByVal pLabelNo As String, ByVal xstatus As String, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            'sql = " Select * from PrintLabelExport " & vbCrLf & _
            '            " WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
            '            " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' and LabelNo = '" & Trim(pLabelNo) & "' "



            'If i = 0 Then
            If xstatus = "G" Then xstatus = "0" Else xstatus = "1"

            sql = " Update PrintLabelExport SET SuratJalanNo_FWD = '" & Trim(pTmp.SuratJalanNo) & "', StatusDefect = '" & xstatus & "' " & vbCrLf & _
                  " WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' and LabelNo = '" & Trim(pLabelNo) & "' "
            sql = sql + " Update PrintLabelExport SET SuratJalanNo_FWD = '" & Trim(pTmp.SuratJalanNo) & "', StatusDefect = '" & xstatus & "' " & vbCrLf & _
                  " WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.PONo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' and LabelNo = '" & Trim(pLabelNo) & "' "
            sqlCom.CommandText = sql

            Dim i As Integer = sqlCom.ExecuteNonQuery
            sqlCom.CommandText = sql
            i = sqlCom.ExecuteNonQuery()
            'End If

            Return i
        End Using

    End Function

    Public Shared Function RemainingReceiveExport(ByVal pTmp As clsTmp, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " Update remainingreceive_Export SET Status = '1' " & vbCrLf & _
                  " WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(pTmp.OrderNo) & "'-- AND PartNo = '" & Trim(pTmp.PartNo) & "' "
            sqlCom.CommandText = sql

            Dim i As Integer = sqlCom.ExecuteNonQuery
            sqlCom.CommandText = sql
            i = sqlCom.ExecuteNonQuery()
            'End If

            Return i
        End Using

    End Function

    Public Shared Function updateRemainingLblEpt(ByVal pTmp As clsTmp, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " Update PrintLabelExport SET statusRemaining = '1' " & vbCrLf & _
                  " WHERE SupplierID = '" & Trim(pTmp.SupplierID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
                  " AND SuratJalanNo_FWD = '' --AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' "
            sqlCom.CommandText = sql

            Dim i As Integer = sqlCom.ExecuteNonQuery
            sqlCom.CommandText = sql
            i = sqlCom.ExecuteNonQuery()
            'End If

            Return i
        End Using

    End Function

    Public Shared Function insertMasterInvoiceEx(ByVal pTmp As clsTmp, ByVal SQLCom As SqlCommand) As Integer
        sql = " UPDATE dbo.InvoiceSupplier_Master_Export " & vbCrLf & _
            " SET InvoiceDate = '" & pTmp.InvoiceDate & "', " & vbCrLf & _
            " 	PIC = '" & pTmp.PIC & "', " & vbCrLf & _
            " 	Remarks = '" & pTmp.Remarks & "', " & vbCrLf & _
            " 	DueDate = '" & pTmp.DueDate & "', " & vbCrLf & _
            " 	PaymentTerm = '" & Trim(pTmp.PaymentItem) & "', " & vbCrLf & _
            " 	UpdateDate = GETDATE(), " & vbCrLf & _
            " 	UpdateUser = '" & Trim(pTmp.EntryUser) & "' " & vbCrLf & _
            " Where InvoiceNo = '" & Trim(pTmp.InvoiceNo) & "' " & vbCrLf & _
            "   And SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' " & vbCrLf & _
            "   And AffiliateID = '" & Trim(pTmp.AffiliateID) & "' " & vbCrLf & _
            "   And SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
            "   And ForwarderID = '" & Trim(pTmp.ForwarderID) & "' " & vbCrLf & _
            "   And PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
            "   And OrderNo = '" & Trim(pTmp.OrderNo) & "' "

        SQLCom.CommandText = sql
        Dim i As Integer = SQLCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.InvoiceSupplier_Master_Export " & vbCrLf & _
              " VALUES  ( '" & Trim(pTmp.InvoiceNo) & "' ," & vbCrLf & _
              "           '" & Trim(pTmp.InvoiceDate) & "' ," & vbCrLf & _
              "           '" & Trim(pTmp.SuratJalanNo) & "' ," & vbCrLf & _
              "           '" & Trim(pTmp.AffiliateID) & "' ," & vbCrLf & _
              "           '" & Trim(pTmp.SupplierID) & "' ," & vbCrLf & _
              "           '" & Trim(pTmp.ForwarderID) & "' ," & vbCrLf & _
              "           '" & Trim(pTmp.PONo) & "' ," & vbCrLf & _
              "           '" & Trim(pTmp.OrderNo) & "' ," & vbCrLf & _
              "           '" & Trim(pTmp.PaymentItem) & "' ," & vbCrLf & _
              "           '" & pTmp.DueDate & "' , " & vbCrLf & _
              "           0 , " & vbCrLf & _
              "           '" & Trim(pTmp.PIC) & "' ,  " & vbCrLf & _
              "           '" & Trim(pTmp.Remarks) & "' ,  " & vbCrLf & _
              "           GETDATE() ,  " & vbCrLf & _
              "           '" & Trim(pTmp.EntryUser) & "' ,  " & vbCrLf & _
              "           GETDATE() , " & vbCrLf & _
              "           '" & Trim(pTmp.UpdateUser) & "' " & vbCrLf & _
              "         ) "

            SQLCom.CommandText = sql
            i = SQLCom.ExecuteNonQuery()
        End If

        Return i
    End Function

    Public Shared Function insertDetailInvoiceEx(ByVal pTmp As clsTmp, ByVal sqlCom As SqlCommand) As Integer
        sql = " UPDATE dbo.InvoiceSupplier_Detail_Export " & vbCrLf & _
            " SET DOQty = '" & pTmp.DOQty & "', " & vbCrLf & _
            " 	InvQty = '" & pTmp.InvoiceQty & "', " & vbCrLf & _
            " 	Price = '" & pTmp.InvPrice & "', " & vbCrLf & _
            " 	InvAmount = '" & pTmp.InvAmount & "', " & vbCrLf & _
            " 	InvCurrCls = '" & pTmp.InvCurrCls & "' " & vbCrLf & _
            " Where InvoiceNo = '" & Trim(pTmp.InvoiceNo) & "' " & vbCrLf & _
            "   And SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' " & vbCrLf & _
            "   And AffiliateID = '" & Trim(pTmp.AffiliateID) & "' " & vbCrLf & _
            "   And SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
            "   And ForwarderID = '" & Trim(pTmp.ForwarderID) & "' " & vbCrLf & _
            "   And PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
            "   And OrderNo = '" & Trim(pTmp.OrderNo) & "' " & vbCrLf & _
            "   And PartNo = '" & Trim(pTmp.PartNo) & "' "

        sqlCom.CommandText = sql
        Dim i As Integer = sqlCom.ExecuteNonQuery

        If i = 0 Then
            sql = " INSERT INTO dbo.InvoiceSupplier_Detail_Export " & vbCrLf & _
                      " VALUES  ( '" & Trim(pTmp.InvoiceNo) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.SuratJalanNo) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.AffiliateID) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.SupplierID) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.ForwarderID) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.PONo) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.OrderNo) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.PartNo) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.DOQty) & "' ," & vbCrLf & _
                      "           '" & Trim(pTmp.InvoiceQty) & "' , " & vbCrLf & _
                      "           '" & Trim(pTmp.InvPrice) & "' , " & vbCrLf & _
                      "           '" & Trim(pTmp.InvAmount) & "' , " & vbCrLf & _
                      "           '" & Trim(pTmp.InvCurrCls) & "'  " & vbCrLf & _
                      "         ) "

            sqlCom.CommandText = sql
            i = sqlCom.ExecuteNonQuery()
        End If

        sql = " UPDATE InvoiceSupplier_Master_Export " & vbCrLf & _
              " SET TotalAmount = (Select SUM(InvAmount) Amount From InvoiceSupplier_Detail_Export " & vbCrLf & _
              "                   Where InvoiceNo = '" & Trim(pTmp.InvoiceNo) & "' " & vbCrLf & _
              "                   And   SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' " & vbCrLf & _
              "                   And   AffiliateID = '" & Trim(pTmp.AffiliateID) & "' " & vbCrLf & _
              "                   And   SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
              "                   And   ForwarderID = '" & Trim(pTmp.ForwarderID) & "' " & vbCrLf & _
              "                   And   PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
              "                   And   OrderNo = '" & Trim(pTmp.OrderNo) & "' )" & vbCrLf & _
              " Where InvoiceNo = '" & Trim(pTmp.InvoiceNo) & "' " & vbCrLf & _
              " And   SuratJalanNo = '" & Trim(pTmp.SuratJalanNo) & "' " & vbCrLf & _
              " And   AffiliateID = '" & Trim(pTmp.AffiliateID) & "' " & vbCrLf & _
              " And   SupplierID = '" & Trim(pTmp.SupplierID) & "' " & vbCrLf & _
              " And   ForwarderID = '" & Trim(pTmp.ForwarderID) & "' " & vbCrLf & _
              " And   PONo = '" & Trim(pTmp.PONo) & "' " & vbCrLf & _
              " And   OrderNo = '" & Trim(pTmp.OrderNo) & "' "

        sqlCom.CommandText = sql
        sqlCom.ExecuteNonQuery()

        Return i
    End Function

    Public Shared Function UpdateShippingInstruction(ByVal pTmp As clsTmp, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " UPDATE ShippingInstruction_Master" & vbCrLf & _
                      " SET ETDPort = '" & Trim(pTmp.ETDJakarta) & "', " & vbCrLf & _
                      "     GrossWeight = '" & Trim(pTmp.Gross) & "', " & vbCrLf & _
                      "     ShippingLineS = '" & Trim(pTmp.ShippingLine) & "', " & vbCrLf & _
                      "     NamaKapalS = '" & Trim(pTmp.NamaKapal) & "', " & vbCrLf & _
                      "     VesselS = '" & Trim(pTmp.Vassel) & "', " & vbCrLf & _
                      "     UpdateUser = 'Batch', " & vbCrLf & _
                      "     UpdateDate = Getdate() " & vbCrLf

            sql = sql + " WHERE ShippingInstructionNo = '" & Trim(pTmp.InvoiceNo) & "' " & vbCrLf
            sqlCom.CommandText = sql

            Dim i As Integer = sqlCom.ExecuteNonQuery

            Return i
        End Using

    End Function

    Public Shared Function insertMasterTaily(ByVal pTmp As clsTmp, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " UPDATE Tally_Master" & vbCrLf & _
                      " SET SealNo = '" & Trim(pTmp.SealNo) & "', " & vbCrLf & _
                      "     Tare = '" & Trim(pTmp.Tare) & "', " & vbCrLf & _
                      "     Gross = '" & Trim(pTmp.Gross) & "', " & vbCrLf & _
                      "     TotalCarton = '" & Trim(pTmp.TotalCarton) & "', " & vbCrLf & _
                      "     Vessel = '" & Trim(pTmp.Vassel) & "', " & vbCrLf & _
                      "     DNNo = '" & Trim(pTmp.DONo) & "', " & vbCrLf & _
                      "     ContainerSize = '" & Trim(pTmp.SizeContainer) & "', " & vbCrLf & _
                      "     ETD = '" & Trim(pTmp.ETDJakarta) & "', " & vbCrLf & _
                      "     ShippingLine = '" & Trim(pTmp.ShippingLine) & "', " & vbCrLf & _
                      "     DestinationPort = '" & Trim(pTmp.DestinationPort) & "', " & vbCrLf & _
                      "     StuffingDate = " & IIf(pTmp.StuffingDate = "", "NULL", "'" & pTmp.StuffingDate & "'") & "," & vbCrLf & _
                      "     NamaKapal = '" & Trim(pTmp.NamaKapal) & "' " & vbCrLf

            sql = sql + " WHERE ShippingInstructionNo = '" & Trim(pTmp.InvoiceNo) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND ContainerNo = '" & Trim(pTmp.ContainerNo) & "' " & vbCrLf
            sqlCom.CommandText = sql

            Dim i As Integer = sqlCom.ExecuteNonQuery

            If i = 0 Then

                sql = " INSERT INTO Tally_Master " & vbCrLf & _
                      " (ShippingInstructionNo,ForwarderID,AffiliateID, ContainerNo,SealNo,Tare, " & vbCrLf & _
                      " Gross, TotalCarton, Vessel, DNNo, ContainerSize, ETD, ShippingLine, DestinationPort, NamaKapal, StuffingDate, TallyCls2) " & vbCrLf & _
                      "  "
                sql = sql + " VALUES  ( '" & Trim(pTmp.InvoiceNo) & "','" & Trim(pTmp.ForwarderID) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.AffiliateID) & "' , '" & Trim(pTmp.ContainerNo) & "', " & vbCrLf & _
                            "           '" & Trim(pTmp.SealNo) & "' , '" & Trim(pTmp.Tare) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.Gross) & "', '" & Trim(pTmp.TotalCarton) & "'," & vbCrLf & _
                            "           '" & Trim(pTmp.Vassel) & "', '" & Trim(pTmp.DONo) & "'," & vbCrLf & _
                            "           '" & Trim(pTmp.SizeContainer) & "', '" & Trim(pTmp.ETDJakarta) & "', " & vbCrLf & _
                            "           '" & Trim(pTmp.ShippingLine) & "', '" & Trim(pTmp.DestinationPort) & "'," & vbCrLf & _
                            "           '" & Trim(pTmp.NamaKapal) & "', " & vbCrLf & _
                            "           " & IIf(pTmp.StuffingDate = "", "NULL", "'" & pTmp.StuffingDate & "'") & ",'1')" & vbCrLf

                sqlCom.CommandText = sql
                i = sqlCom.ExecuteNonQuery()

            End If

            Return i
        End Using

    End Function

    Public Shared Function insertDetailTaily(ByVal pTmp As clsTmp, ByVal sqlCom As SqlCommand) As Integer

        Using SQLCon As New SqlConnection(uf_GetConString)
            SQLCon.Open()

            sql = " UPDATE Tally_Detail" & vbCrLf & _
                      " SET Length = '" & Trim(pTmp.Length) & "', Width = '" & Trim(pTmp.Width) & "', Height = '" & Trim(pTmp.Height) & "', " & vbCrLf & _
                      "     M3 = '" & Trim(pTmp.M3) & "', WeightPallet = '" & Trim(pTmp.WeightPallet) & "', caseNo2 = '" & Trim(pTmp.BoxNo2) & "', " & vbCrLf & _
                      "     TotalBox = '" & Trim(pTmp.TotBoxEx) & "', ContainerNo = '" & Trim(pTmp.ContainerNo) & "', " & vbCrLf & _
                      "     POMOQ = '" & uf_GetMOQ(4, pTmp.OrderNo, pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID, pTmp.InvoiceNo) & "', " & vbCrLf & _
                      "     POQtyBox = '" & uf_GetQtybox(4, pTmp.OrderNo, pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID, pTmp.InvoiceNo) & "' " & vbCrLf

            sql = sql + " WHERE ShippingInstructionNo = '" & Trim(pTmp.InvoiceNo) & "' AND ForwarderID = '" & Trim(pTmp.ForwarderID) & "' " & vbCrLf & _
                        "       AND AffiliateID = '" & Trim(pTmp.AffiliateID) & "' AND PalletNo = '" & Trim(pTmp.PalletNo) & "' " & vbCrLf & _
                        "       AND OrderNo = '" & Trim(pTmp.OrderNo) & "' AND PartNo = '" & Trim(pTmp.PartNo) & "' AND CaseNo = '" & Trim(pTmp.BoxNo) & "' --and CaseNo2 = '" & Trim(pTmp.BoxNo2) & "'" & vbCrLf 'Ga perlu CaseNo2 29112021

            sqlCom.CommandText = sql

            Dim i As Integer = sqlCom.ExecuteNonQuery

            If i = 0 Then

                sql = " INSERT INTO Tally_Detail " & vbCrLf & _
                      " (ShippingInstructionNo,ForwarderID,AffiliateID, PalletNo, OrderNo, PartNo, " & vbCrLf & _
                      " CaseNo, Length, Width, Height, M3, WeightPallet, caseNo2, TotalBox, ContainerNo, POMOQ, POQtyBox ) " & vbCrLf & _
                      "  "
                sql = sql + " VALUES  ( '" & Trim(pTmp.InvoiceNo) & "','" & Trim(pTmp.ForwarderID) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.AffiliateID) & "' , '" & Trim(pTmp.PalletNo) & "', " & vbCrLf & _
                            "           '" & Trim(pTmp.OrderNo) & "' , '" & Trim(pTmp.PartNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pTmp.BoxNo) & "', '" & Trim(pTmp.Length) & "'," & vbCrLf & _
                            "           '" & Trim(pTmp.Width) & "', '" & Trim(pTmp.Height) & "'," & vbCrLf & _
                            "           '" & Trim(pTmp.M3) & "', '" & Trim(pTmp.WeightPallet) & "', " & vbCrLf & _
                            "           '" & Trim(pTmp.BoxNo2) & "'," & vbCrLf & _
                            "           " & Trim(pTmp.TotBoxEx) & ", " & vbCrLf & _
                            "           '" & Trim(pTmp.ContainerNo) & "', " & vbCrLf & _
                            "           '" & uf_GetMOQ(4, pTmp.OrderNo, pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID, pTmp.InvoiceNo) & "', " & vbCrLf & _
                            "           '" & uf_GetQtybox(4, pTmp.OrderNo, pTmp.PartNo, pTmp.SupplierID, pTmp.AffiliateID, pTmp.InvoiceNo) & "' " & vbCrLf & _
                            "           )" & vbCrLf

                sqlCom.CommandText = sql
                i = sqlCom.ExecuteNonQuery()
            End If

            Return i
        End Using

    End Function

    Public Shared Function uf_GetMOQ(ByVal pType As Integer, ByVal pPoNo As String, ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, Optional ByVal pShipNo As String = "") As Integer
        Dim MOQ As Integer = 0

        Using Cn As New SqlConnection(uf_GetConString)
            Dim ls_SQL As String = ""

            If pType = 0 Then ' dari partMapping
                ls_SQL = "SELECT ISNULL(MOQ,0) MOQ FROM dbo.MS_PartMapping WHERE PartNo='" + pPartNo + "' AND SupplierID='" + pSupplierID + "' AND AffiliateID='" + pAffiliateID + "'"
            ElseIf pType = 1 Then ' dari podetail
                ls_SQL = "SELECT ISNULL(a.POMOQ,b.MOQ) MOQ FROM dbo.PO_Detail a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE PONo='" + pPoNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "'"
            ElseIf pType = 2 Then ' dari podetail Export
                ls_SQL = "SELECT ISNULL(a.POMOQ,b.MOQ) MOQ FROM dbo.PO_Detail_Export a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE PONo='" + pPoNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "'"
            ElseIf pType = 3 Then ' dari podetail Export using OrderNo
                ls_SQL = "SELECT ISNULL(a.POMOQ,b.MOQ) MOQ FROM dbo.PO_Detail_Export a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE OrderNo1 ='" + pPoNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "'"
            ElseIf pType = 4 Then
                ls_SQL = "select ISNULL(b.POMOQ,ISNULL(a.POMOQ,c.MOQ)) MOQ from PO_Detail_Export a left join " & vbCrLf & _
                    "(select AffiliateID, SupplierID, OrderNo, PartNo, POMOQ, POQtyBox from ShippingInstruction_Detail where ShippingInstructionNo = '" + pShipNo + "' and AffiliateID = '" + pAffiliateID + "' and PartNo = '" + pPartNo + "' and OrderNo = '" + pPoNo + "') b " & vbCrLf & _
                    "on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.OrderNo and a.PartNo = b.PartNo left join MS_PartMapping c on a.PartNo = c.PartNo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID " & vbCrLf & _
                    "where a.PONo = '" + pPoNo + "' and a.AffiliateID = '" + pAffiliateID + "' and a.PartNo = '" + pPartNo + "' "
            End If

            ds = uf_GetDataSet(ls_SQL)
            If ds.Tables(0).Rows.Count > 0 Then
                MOQ = ds.Tables(0).Rows(0)("MOQ")
            End If
        End Using
        Return MOQ
    End Function

    Public Shared Function uf_GetQtybox(ByVal pType As Integer, ByVal pPoNo As String, ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, Optional ByVal pShipNo As String = "") As Integer
        Dim Qty As Integer = 0
        Using Cn As New SqlConnection(uf_GetConString)
            Dim ls_SQL As String = ""

            If pType = 0 Then ' dari partMapping
                ls_SQL = "SELECT ISNULL(QtyBox,0) Qty dbo.MS_PartMapping WHERE PartNo='" + pPartNo + "' AND SupplierID='" + pSupplierID + "' AND AffiliateID='" + pAffiliateID + "'"
            ElseIf pType = 1 Then ' dari podetail
                ls_SQL = "SELECT ISNULL(a.POQtyBox,b.QtyBox) Qty FROM dbo.PO_Detail a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE PONo='" + pPoNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "'"
            ElseIf pType = 2 Then ' dari podetail Export
                ls_SQL = "SELECT ISNULL(a.POQtyBox,b.QtyBox) Qty FROM dbo.PO_Detail_Export a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE PONo ='" + pPoNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "'"
            ElseIf pType = 3 Then ' dari podetail Export using OrderNo
                ls_SQL = "SELECT ISNULL(a.POQtyBox,b.QtyBox) Qty FROM dbo.PO_Detail_Export a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE OrderNo1 ='" + pPoNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "'"
            ElseIf pType = 4 Then
                ls_SQL = "select ISNULL(b.POQtyBox,ISNULL(a.POQtyBox,c.QtyBox)) Qty from PO_Detail_Export a left join " & vbCrLf & _
                    "(select AffiliateID, SupplierID, OrderNo, PartNo, POMOQ, POQtyBox from ShippingInstruction_Detail where ShippingInstructionNo = '" + pShipNo + "' and AffiliateID = '" + pAffiliateID + "' and PartNo = '" + pPartNo + "' and OrderNo = '" + pPoNo + "') b " & vbCrLf & _
                    "on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.OrderNo and a.PartNo = b.PartNo left join MS_PartMapping c on a.PartNo = c.PartNo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID " & vbCrLf & _
                    "where a.PONo = '" + pPoNo + "' and a.AffiliateID = '" + pAffiliateID + "' and a.PartNo = '" + pPartNo + "' "
            End If

            ds = uf_GetDataSet(ls_SQL)
            If ds.Tables(0).Rows.Count > 0 Then
                Qty = ds.Tables(0).Rows(0)("Qty")
            End If
        End Using
        Return Qty
    End Function

    Public Shared Function uf_GetPriceSupplier(ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, Optional ByVal pShipNo As String = "") As Double
        Dim price As Double = 0
        Using Cn As New SqlConnection(uf_GetConString)
            Dim ls_SQL As String = ""

            ls_SQL = "select ISNULL(b.Price,0) Price " & vbCrLf & _
                     "from MS_PartMapping a join MS_Price b on a.PartNo = b.PartNo and a.AffiliateID = b.DeliveryLocationID and a.SupplierID = b.AffiliateID " & vbCrLf & _
                     "where a.AffiliateID = '" & pAffiliateID & "' and a.PartNo = '" & pPartNo & "' and '" & Now.ToString("yyyy-MM-dd") & "' between b.StartDate and b.EndDate and a.SupplierID = '" & pSupplierID & "' " & vbCrLf
            ds = uf_GetDataSet(ls_SQL)
            If ds.Tables(0).Rows.Count > 0 Then
                price = ds.Tables(0).Rows(0)("Price")
            End If
        End Using
        Return price
    End Function

End Class
