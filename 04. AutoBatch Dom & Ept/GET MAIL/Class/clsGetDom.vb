Imports OpenPop.Pop3
Imports OpenPop.Mime
Imports System.Data.SqlClient
Imports System.IO

Public Class clsGetDom
    Shared Sub up_GetEmailDom(ByVal cfg As GlobalSetting.clsConfig,
                                  ByVal log As GlobalSetting.clsLog,
                                  ByVal GB As GlobalSetting.clsGlobal,
                                  ByVal LogName As RichTextBox,
                                  ByVal pHost As String,
                                  ByVal pPort As String,
                                  ByVal pUserName As String,
                                  ByVal pPassword As String,
                                  ByVal pBackupFolder As String,
                                  ByVal pScreenName As String,
                                  Optional ByRef errMsg As String = "",
                                  Optional ByRef ErrSummary As String = "")

        Dim pop3Client As Pop3Client
        Dim message As Message

        Dim countEmail As Integer

        Dim eMailFrom As String
        Dim eMailSendDate As String
        Dim eMailSubject As String()

        Dim ds As DataSet

        Dim pInsert As Boolean = False

        Dim pPONO As String = ""
        Dim pPartNo As String = ""
        Dim pSupplier As String = ""

        Try
            Application.DoEvents()
            log.WriteToProcessLog(Date.Now, pScreenName, "Connect to pop e-mail: " & pHost)
            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Connect to pop e-mail: " & pHost, LogName)
            LogName.Refresh()

            pop3Client = New Pop3Client()
            pop3Client.Connect(pHost, Val(pPort), False)

            If pop3Client.Connected = False Then
                Application.DoEvents()
                log.WriteToProcessLog(Date.Now, pScreenName, "Disconnect to pop e-mail: " & pHost)
                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Disconnect to e-mail: " & pHost, LogName)
                LogName.Refresh()
            Else
                Application.DoEvents()
                log.WriteToProcessLog(Date.Now, pScreenName, "Connect to e-mail address: " & pUserName)
                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Connect to e-mail address: " & pUserName, LogName)
                LogName.Refresh()
                pop3Client.Authenticate(pUserName, pPassword, AuthenticationMethod.UsernameAndPassword)

                countEmail = pop3Client.GetMessageCount()

                For i As Integer = countEmail To 1 Step -1
                    pInsert = False

                    message = pop3Client.GetMessage(i)

                    Try
                        eMailFrom = message.Headers.From.Address.ToString
                    Catch ex As Exception
                        eMailFrom = ""
                    End Try

                    Try
                        eMailSubject = Split(message.Headers.Subject.ToString, ";")
                    Catch ex As Exception
                        eMailSubject = Nothing
                    End Try

                    Try
                        eMailSendDate = message.Headers.DateSent.ToString
                    Catch ex As Exception
                        eMailSendDate = ""
                    End Try

                    Try
                        pSupplier = eMailSubject(1)
                    Catch ex As Exception
                        pSupplier = ""
                    End Try

                    Try
                        pPONO = eMailSubject(2)
                    Catch ex As Exception
                        pPONO = ""
                    End Try

                    Try
                        pPartNo = eMailSubject(3)
                    Catch ex As Exception
                        pPartNo = ""
                    End Try

                    '1. Check Email, Jika dari PEMI dan TOS bypass pInsert = True
                    If Split(eMailFrom, "@")(1) = "pemi.co.id" Or Split(eMailFrom, "@")(1) = "tos.co.id" Then
                        pInsert = True
                    End If

                    '2. Check Email Supplier, Jika tidak terdarftar di Master pInsert = False
                    If pInsert = False Then
                        ds = cekEmailSupplier(GB, pSupplier.Trim, eMailFrom.Trim)
                        If ds Is Nothing Then
                            pInsert = False
                        Else
                            pInsert = True
                        End If

                    End If

                    If Trim(pSupplier) <> "" Then
                        '3. Jika pInsert = True insert ke database
                        If pInsert = True Then
                            Using SQLCon As New SqlConnection(cfg.ConnectionString)
                                SQLCon.Open()

                                Dim SQLCom As SqlCommand = SQLCon.CreateCommand
                                Dim SQLTrans As SqlTransaction

                                SQLTrans = SQLCon.BeginTransaction
                                SQLCom.Connection = SQLCon
                                SQLCom.Transaction = SQLTrans

                                If UCase(eMailSubject(0)) = "CODE_SO" Then
                                    Application.DoEvents()
                                    log.WriteToProcessLog(Date.Now, pScreenName, "Update Request Summary Outstanding start: " & eMailFrom)
                                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Update Request Summary Outstanding start: " & eMailFrom, LogName)
                                    LogName.Refresh()

                                    insertOutstandingEmail(eMailFrom, pSupplier, pPONO, pPartNo, SQLCom)

                                    Application.DoEvents()
                                    log.WriteToProcessLog(Date.Now, pScreenName, "Update Request Summary Outstanding end: " & eMailFrom)
                                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Update Request Summary Outstanding end: " & eMailFrom, LogName)
                                    LogName.Refresh()
                                ElseIf UCase(eMailSubject(0)) = "CODE_SF" Then
                                    Application.DoEvents()
                                    log.WriteToProcessLog(Date.Now, pScreenName, "Update Request Summary Forecast start: " & eMailFrom)
                                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Update Request Summary Forecast start: " & eMailFrom, LogName)
                                    LogName.Refresh()

                                    insertForecastEmail(eMailFrom, pSupplier, pPONO, pPartNo, SQLCom)

                                    Application.DoEvents()
                                    log.WriteToProcessLog(Date.Now, pScreenName, "Update Request Summary Forecast end: " & eMailFrom)
                                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Update Request Summary Forecast end: " & eMailFrom, LogName)
                                    LogName.Refresh()
                                ElseIf UCase(eMailSubject(0)) = "CODE_STOCK" Then
                                    Application.DoEvents()
                                    log.WriteToProcessLog(Date.Now, pScreenName, "Update Request Stock Opname start: " & eMailFrom)
                                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Update Request Stock Opname start: " & eMailFrom, LogName)
                                    LogName.Refresh()

                                    insertFowarderEmail(eMailFrom, pSupplier, pPONO, pPartNo, SQLCom)

                                    Application.DoEvents()
                                    log.WriteToProcessLog(Date.Now, pScreenName, "Update Request Stock Opname end: " & eMailFrom)
                                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Update Request Stock Opname end: " & eMailFrom, LogName)
                                    LogName.Refresh()
                                End If

                                SQLTrans.Commit()
                            End Using
                        Else
                            Application.DoEvents()
                            log.WriteToProcessLog(Date.Now, pScreenName, "The Sender's email address: " & eMailFrom & " not exists in database")
                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "The Sender's email address: " & eMailFrom & " not exists in database", LogName)
                            LogName.Refresh()
                        End If
                    End If

                    'DOWNLOAD ATTACHMENT
                    Try
                        Dim eAttachments As List(Of MessagePart) = message.FindAllAttachments()
                        Dim iEmail As Integer = 0
                        Dim pAttachment As String = ""

                        Application.DoEvents()
                        'msgInfo = "download attachment email process..."
                        'gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

                        For Each attachment As MessagePart In eAttachments
                            pAttachment = attachment.FileName

                            Try
                                If Right(pAttachment.Trim, 3) = "xls" Or Right(pAttachment.Trim, 4) = "xlsx" Or Right(pAttachment.Trim, 4) = "xlsm" Then
                                    File.WriteAllBytes(Trim(pBackupFolder) & "\" & attachment.FileName.ToString, attachment.Body)
                                    iEmail = iEmail + 1
                                End If
                            Catch ex As Exception                                
                                log.WriteToProcessLog(Date.Now, pScreenName, "Download attachement error: " & eMailFrom & "===" & ex.Message)
                            End Try                            
                        Next

                        If iEmail > 0 Then
                            Application.DoEvents()
                            log.WriteToProcessLog(Date.Now, pScreenName, "Download attachement success: " & eMailFrom)
                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Download attachement success: " & eMailFrom, LogName)
                            LogName.Refresh()
                        End If                        
                    Catch ex As Exception
                        Application.DoEvents()
                        log.WriteToProcessLog(Date.Now, pScreenName, "Download attachement error: " & eMailFrom & "===" & ex.Message)
                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Download attachement error: " & eMailFrom & "===" & ex.Message, LogName)
                        LogName.Refresh()
                    End Try

                    pop3Client.DeleteMessage(i)
                Next

                pop3Client.Disconnect()
                Application.DoEvents()
                log.WriteToProcessLog(Date.Now, pScreenName, "Disconnect to pop e-mail: " & pHost)
                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Disconnect to pop e-mail: " & pHost, LogName)
                LogName.Refresh()

            End If
        Catch ex As Exception
            errMsg = ex.Message
            ErrSummary = ex.Message
        Finally
            Application.DoEvents()
            log.WriteToProcessLog(Date.Now, pScreenName, "Disconnect to e-mail: " & pUserName)
            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Disconnect to e-mail: " & pUserName, LogName)
            LogName.Refresh()
        End Try
    End Sub

    Shared Function cekEmailSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal SupplierID As String, ByVal EmailAddress As String) As DataSet
        Dim ls_Sql As String

        ls_Sql = "SELECT SupplierID FROM dbo.MS_EmailSupplier WHERE SupplierID = '" & Trim(SupplierID) & "' " & vbCrLf & _
                 "and (SummaryOutstandingTO LIKE '%" & EmailAddress & "%' OR SummaryOutstandingCC LIKE '%" & EmailAddress & "%')" & vbCrLf
        ls_Sql = ls_Sql + "UNION ALL" & vbCrLf & _
                 "SELECT ForwarderID FROM dbo.MS_EmailForwarder WHERE ForwarderID = '" & Trim(SupplierID) & "' " & vbCrLf & _
                 "and (ForwarderReceivingTo LIKE '%" & EmailAddress & "%' OR ForwarderReceivingCC LIKE '%" & EmailAddress & "%')"

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_Sql)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Function insertOutstandingEmail(ByVal pEmailFrom As String, ByVal pSupplierID As String, _
                                           ByVal pPONo As String, ByVal pPartNo As String, ByVal SQLCom As SqlCommand) As Integer
        Dim ls_Sql As String

        Try
            ls_Sql = " INSERT INTO dbo.SupplierSumOutstanding_Request " & vbCrLf & _
                        " VALUES  ( '" & Trim(pEmailFrom) & "' ," & vbCrLf & _
                        "           GETDATE() ," & vbCrLf & _
                        "           '" & Trim(pSupplierID) & "' ," & vbCrLf & _
                        "           '" & Trim(pPONo) & "' ," & vbCrLf & _
                        "           '" & Trim(pPartNo) & "' ," & vbCrLf & _
                        "           '1' ,  " & vbCrLf & _
                        "           GETDATE() ) "

            SQLCom.CommandText = ls_Sql
            Dim i As Integer = SQLCom.ExecuteNonQuery
            Return i
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Shared Function insertForecastEmail(ByVal pEmailFrom As String, ByVal pSupplierID As String, _
                                          ByVal pPONo As String, ByVal pPartNo As String, ByVal SQLCom As SqlCommand) As Integer
        Dim ls_Sql As String

        Try
            ls_Sql = " INSERT INTO dbo.SupplierSumForecast_Request " & vbCrLf & _
                        " VALUES  ( '" & Trim(pEmailFrom) & "' ," & vbCrLf & _
                        "           GETDATE() ," & vbCrLf & _
                        "           '" & Trim(pSupplierID) & "' ," & vbCrLf & _
                        "           '" & Trim(pPONo) & "' ," & vbCrLf & _
                        "           '" & Trim(pPartNo) & "' ," & vbCrLf & _
                        "           '1' ,  " & vbCrLf & _
                        "           GETDATE() ) "

            SQLCom.CommandText = ls_Sql
            Dim i As Integer = SQLCom.ExecuteNonQuery
            Return i
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Shared Function insertFowarderEmail(ByVal pEmailFrom As String, ByVal pSupplierID As String, _
                                           ByVal pPONo As String, ByVal pPartNo As String, ByVal SQLCom As SqlCommand) As Integer
        Dim ls_Sql As String

        Try
            ls_Sql = " INSERT INTO dbo.ForwarderStockOpname_Request " & vbCrLf & _
                        " VALUES  ( '" & Trim(pEmailFrom) & "' ," & vbCrLf & _
                        "           GETDATE() ," & vbCrLf & _
                        "           '" & Trim(pSupplierID) & "' ," & vbCrLf & _
                        "           '" & Trim(pPONo) & "' ," & vbCrLf & _
                        "           '" & Trim(pPartNo) & "' ," & vbCrLf & _
                        "           '1' ,  " & vbCrLf & _
                        "           GETDATE() ) "

            SQLCom.CommandText = ls_Sql
            Dim i As Integer = SQLCom.ExecuteNonQuery
            Return i
        Catch ex As Exception
            Return 0
        End Try
    End Function
End Class
