Imports System.Data
Imports System.Data.SqlClient

Public Class clsSynchronizeData
    Shared Sub up_SynchronizeData(ByVal cfg As GlobalSetting.clsConfig,
                             ByVal log As GlobalSetting.clsLog,
                             ByVal GB As GlobalSetting.clsGlobal,
                             ByVal LogName As RichTextBox,
                             ByVal pAtttacment As String,
                             ByVal pResult As String,
                             ByVal pScreenName As String,
                             Optional ByRef errMsg As String = "",
                             Optional ByRef ErrSummary As String = "")

        Dim ls_SQL As String = ""

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE ReceivePASI_Detail SET  " & vbCrLf & _
                          " 		GoodRecQty = (SELECT COUNT(BoxNo) * b.QtyBox FROM ReceivePASISeq_Detail a  " & vbCrLf & _
                          " 					  left join MS_PartMapping b on a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                          " 					  and a.SupplierID = b.SupplierID and a.PartNo = b.PartNo " & vbCrLf & _
                          " 					  WHERE a.SuratJalanNo = ReceivePASI_Detail.SuratJalanNo and a.PartNo = ReceivePASI_Detail.PartNo " & vbCrLf & _
                          " 						and a.KanbanNo = ReceivePASI_Detail.KanbanNo and a.AffiliateID = ReceivePASI_Detail.AffiliateID " & vbCrLf & _
                          " 						and a.SupplierID = ReceivePASI_Detail.SupplierID " & vbCrLf & _
                          " 					  GROUP BY QtyBox)  " & vbCrLf & _
                          " where exists  " & vbCrLf & _
                          " ( " & vbCrLf & _
                          " 	select * from ReceivePASI_Master a "

                ls_SQL = ls_SQL + " 	where ReceiveDate >= '2017-01-01' and ReceivePASI_Detail.SuratJalanNo = a.SuratJalanNo  " & vbCrLf & _
                                  " 		and a.AffiliateID = ReceivePASI_Detail.AffiliateID and a.SupplierID = ReceivePASI_Detail.SupplierID " & vbCrLf & _
                                  "  " & vbCrLf & _
                                  " )  and exists " & vbCrLf & _
                                  " ( " & vbCrLf & _
                                  " 	SELECT * FROM ReceivePASISeq_Detail a  " & vbCrLf & _
                                  " 					  left join MS_PartMapping b on a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                                  " 					  and a.SupplierID = b.SupplierID and a.PartNo = b.PartNo " & vbCrLf & _
                                  " 					  WHERE a.SuratJalanNo = ReceivePASI_Detail.SuratJalanNo and a.PartNo = ReceivePASI_Detail.PartNo " & vbCrLf & _
                                  " 						and a.KanbanNo = ReceivePASI_Detail.KanbanNo and a.AffiliateID = ReceivePASI_Detail.AffiliateID " & vbCrLf & _
                                  " 						and a.SupplierID = ReceivePASI_Detail.SupplierID  "

                ls_SQL = ls_SQL + " 					  GROUP BY QtyBox " & vbCrLf & _
                                  " 					  HAVING (COUNT(BoxNo) * b.QtyBox) <> ReceivePASI_Detail.GoodRecQty " & vbCrLf & _
                                  " ) " & vbCrLf & _
                                  " and ISNULL(SuratJalanNo,'') <> '' " & vbCrLf & _
                                  "  "

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = ex.Message
            ErrSummary = ex.Message
        End Try

    End Sub

End Class
