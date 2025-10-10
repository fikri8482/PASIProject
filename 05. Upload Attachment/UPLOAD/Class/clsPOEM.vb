Public Class clsPOEM
    Shared Sub up_Upload(ByVal EF As Excel.Worksheet, ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pBackup As String,
                              ByVal pBackupError As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Try
            Dim i As Integer
            Dim UploadDetailList As New List(Of clsUploadExcel)
            Dim UploadHeader As New clsUploadExcel


            UploadHeader.AffiliateID = EF.Range("H3").Value.ToString.Trim
            UploadHeader.ForwarderID = EF.Range("H4").Value.ToString.Trim
            UploadHeader.SupplierID = EF.Range("H5").Value.ToString.Trim
            UploadHeader.PONo = EF.Range("I9").Value.ToString.Trim
            If IsDBNull(EF.Range("P9").Value) Or IsNothing(EF.Range("P9").Value) Then
                UploadHeader.OrderNo = ""
            Else
                UploadHeader.OrderNo = EF.Range("P9").Value.ToString.Trim
            End If


            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim UploadDetail As New clsUploadExcel
                UploadDetail.PartNo = Trim(ds.Tables(0).Rows(i)("SMTP"))
                UploadDetail.PartNo = Trim(ds.Tables(0).Rows(i)("PORTSMTP"))
                UploadDetail.PartNo = If(IsDBNull(ds.Tables(0).Rows(i)("usernameSMTP")), "", ds.Tables(0).Rows(i)("usernameSMTP"))
                UploadDetail.PartNo = If(IsDBNull(ds.Tables(0).Rows(i)("passwordSMTP")), "", ds.Tables(0).Rows(i)("passwordSMTP"))
                UploadDetailList.Add(UploadDetail)
            Next

        Catch ex As Exception
            errMsg = "PONo [" & "" & "], Supplier [" & "" & "], Affiliate [" & "" & "] " & ex.Message
            ErrSummary = "PONo [" & "" & "], Supplier [" & "" & "], Affiliate [" & "" & "] " & ex.Message
        End Try

    End Sub
End Class
