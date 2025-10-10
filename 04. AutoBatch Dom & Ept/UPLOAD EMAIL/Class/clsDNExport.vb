Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Transactions

Public Class clsDNExport
    Shared Sub up_DNExport(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResultDom As String,
                              ByVal pResultExp As String,
                              ByVal pScreenName As String,
                              ByVal pFileName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Dim cls As New clsReceivingExportProperty
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application

        Dim sheetNumber As Integer = 1
        Dim checkHeader As Boolean = True

        Dim tempPONo As String
        Dim tempOrderNo As String
        Dim tempTotalBox As Double

        Try
            Dim ls_file As String = pAtttacment & "\" & pFileName
            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            log.WriteToProcessLog(Date.Now, pScreenName, "Read Sheet [" & sheetNumber & "], FileName [" & pFileName & "]")

            If ExcelSheet.Range("H3").Value Is Nothing Then
                checkHeader = False
                pub_ErrorMessage = "Header Template tidak sesuai (Consignee Code kosong). Silahkan dicek kembali Template yang disubmit!"
                GoTo step001
            Else
                cls.AffiliateID = Trim(ExcelSheet.Range("H3").Value.ToString & "")
                'cls.AffiliateID = clsGeneral.AffiliateConsignee(GB, cls.AffiliateID)
            End If

            If ExcelSheet.Range("H4").Value Is Nothing Then
                checkHeader = False
                pub_ErrorMessage = "Header Template tidak sesuai (Delivery Location kosong). Silahkan dicek kembali Template yang disubmit!"
                GoTo step001
            Else
                cls.DeliveryLocation = Trim(ExcelSheet.Range("H4").Value.ToString & "")
            End If

            If ExcelSheet.Range("H5").Value Is Nothing Then
                checkHeader = False
                pub_ErrorMessage = "Header Template tidak sesuai (Supplier Code kosong). Silahkan dicek kembali Template yang disubmit!"
                GoTo step001
            Else
                cls.SupplierID = Trim(ExcelSheet.Range("H5").Value.ToString & "")
            End If

step001:
            If checkHeader = False Then
                '03. Move to Error Folder
                If Not IsNothing(ExcelBook) Then
                    ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If

                log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                LogName.Refresh()
                Exit Try
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Read Header [" & sheetNumber & "], FileName [" & pFileName & "]")

            If ExcelSheet.Range("I28").Value Is Nothing Then
                '03. Move to Error Folder
                If Not IsNothing(ExcelBook) Then
                    ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
                pub_ErrorMessage = "Surat Jalan No. Blank, Please check this file again!"
                sendEmailtoSupplier(GB, pAtttacment, pFileName, "", "", cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                LogName.Refresh()
                Exit Try
            Else
                cls.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I28").Value.ToString & "", 20)
            End If

            If ExcelSheet.Range("AE13").Value Is Nothing Then
                '03. Move to Error Folder
                If Not IsNothing(ExcelBook) Then
                    ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
                pub_ErrorMessage = "PO No. Blank, Please check this file again!"
                sendEmailtoSupplier(GB, pAtttacment, pFileName, "", "", cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                LogName.Refresh()
                Exit Try
            Else
                tempPONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
            End If

            If ExcelSheet.Range("AE15").Value Is Nothing Then
                tempOrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
            Else
                tempOrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE15").Value.ToString & "", 20)
            End If

            If tempPONo <> tempOrderNo Then
                cls.PONo = tempOrderNo
                cls.OrderNo = tempPONo
            Else
                cls.PONo = tempPONo
                cls.OrderNo = tempOrderNo
            End If

            If ExcelSheet.Range("AP11").Value Is Nothing Then
                cls.CommercialCls = "1"
            Else
                cls.CommercialCls = Trim(ExcelSheet.Range("AP11").Value.ToString)
                cls.CommercialCls = IIf(cls.CommercialCls.ToUpper = "YES", "1", "0")
            End If

            'Check PO already Upload?
            If cekSJSubmited(cls.PONo, cls.OrderNo, cls.SuratJalanNo, cls.SupplierID, cls.AffiliateID, cfg.ConnectionString, errMsg) = True Then
                If Not IsNothing(ExcelBook) Then
                    ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
                pub_ErrorMessage = "PO No. [" & cls.PONo & "], Affiliate [" & cls.AffiliateID & "] already upload in PASI system, Please check this file again!"
                sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                LogName.Refresh()
                Exit Try
            End If

            'Refresh Log
            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Upload File [" & pFileName & "]", LogName)
            LogName.Refresh()

            log.WriteToProcessLog(Date.Now, pScreenName, "Read Detail [" & sheetNumber & "], FileName [" & pFileName & "]")

            Dim recExp As New List(Of clsDNExportProperty)
            Dim startRow As Integer = "34"

            For i = startRow To 10000
                If ExcelSheet.Range("B" & i).Value.ToString = "E" Then
                    Exit For
                End If

                Dim setNilai As New clsDNExportProperty
                setNilai.SupplierID = cls.SupplierID
                setNilai.AffiliateID = cls.AffiliateID
                setNilai.DeliveryLocation = cls.DeliveryLocation
                setNilai.PONo = cls.PONo
                setNilai.OrderNo = cls.OrderNo
                setNilai.SuratJalanNo = cls.SuratJalanNo
                setNilai.PartNo = Trim(ExcelSheet.Range("I" & i).Value.ToString & "")
                setNilai.QtyBox = clsGeneral.getQtyBox(GB, setNilai.PartNo, setNilai.SupplierID, setNilai.AffiliateID)


                Try
                    setNilai.BoxNoFrom = Trim(Trim(ExcelSheet.Range("W" & i).Value))
                    setNilai.BoxNoTo = Trim(Trim(ExcelSheet.Range("Z" & i).Value))
                    If setNilai.BoxNoTo = "" Then
                        setNilai.BoxNoTo = setNilai.BoxNoFrom
                    End If
                Catch ex As Exception
                    setNilai.BoxNoTo = setNilai.BoxNoFrom
                End Try

                'Check Prefix Box NO From and TO must be same
                Try
                    If Microsoft.VisualBasic.Left(Trim(setNilai.BoxNoFrom), 2) <> Microsoft.VisualBasic.Left(Trim(setNilai.BoxNoTo), 2) Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Prefix BoxNo different, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                Catch ex As Exception
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "ROW [" & i & "] Prefix BoxNo different, Please check this file again!"
                    sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                End Try


                'Set Total Nilai DO Qty
                Try
                    setNilai.DOQty = IIf(IsNumeric(ExcelSheet.Range("AO" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AO" & i).Value))
                Catch ex As Exception
                    setNilai.DOQty = 0
                End Try

                'Check Prefix Box NO From and TO must be same
                Try
                    If setNilai.DOQty Mod setNilai.QtyBox <> 0 Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] DO Qty must be multiply from Qty/Box, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                Catch ex As Exception
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "ROW [" & i & "] DO Qty must be multiply from Qty/Box, Please check this file again!"
                    sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                End Try


                Dim i_rec1 As Integer = Microsoft.VisualBasic.Right(Trim(setNilai.BoxNoFrom), 7)
                Dim i_rec2 As Integer = Microsoft.VisualBasic.Right(Trim(setNilai.BoxNoTo), 7)

                tempTotalBox = (i_rec2 - i_rec1) + 1

                setNilai.TotalBox = setNilai.DOQty / setNilai.QtyBox

                If setNilai.DOQty > 0 Then
                    If tempTotalBox <> setNilai.TotalBox Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Total Box is different with Box No From and Box No To, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If

                    Dim i_PrefixLabelNo As String = Microsoft.VisualBasic.Left(Trim(setNilai.BoxNoFrom), 2)

                    For i_rec1 = i_rec1 To i_rec2
                        '2.1 check Register
                        If cekPartNoAndLabelRegister(cls.PONo, cls.OrderNo, setNilai.PartNo, i_PrefixLabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7), cfg.ConnectionString, errMsg) = False Then
                            If Not IsNothing(ExcelBook) Then
                                ExcelBook.Save()
                                xlApp.Workbooks.Close()
                                xlApp.Quit()
                            End If
                            pub_ErrorMessage = "ROW [" & i & "], Box NO [" & i_PrefixLabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7) & "] BoxNo not found with PASI System, Please check this file again!"
                            sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                            log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                            log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                            LogName.Refresh()
                            Exit Try
                        End If

                        '2.1 check already exists
                        If cekPartNoAndLabel(cls.SuratJalanNo, cls.PONo, cls.OrderNo, setNilai.PartNo, i_PrefixLabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7), cls.AffiliateID, cls.SupplierID, cfg.ConnectionString, errMsg) = True Then
                            If Not IsNothing(ExcelBook) Then
                                ExcelBook.Save()
                                xlApp.Workbooks.Close()
                                xlApp.Quit()
                            End If
                            pub_ErrorMessage = "ROW [" & i & "], Box NO [" & i_PrefixLabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7) & "] BoxNo already exists in PASI System with SuratJalanNo [" & errMsg & "], Please check this file again!"
                            sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                            log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                            log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                            LogName.Refresh()
                            Exit Try
                        End If
                    Next
                End If

                setNilai.SeqNo = CekDataDetailDNSupplier(setNilai.PONo, setNilai.OrderNo, setNilai.SupplierID, setNilai.AffiliateID, setNilai.PartNo, setNilai.SuratJalanNo, cfg.ConnectionString, errMsg)

                If setNilai.DOQty > 0 Then
                    recExp.Add(setNilai)
                End If
            Next

            Dim opt As Transactions.TransactionOptions
            opt.IsolationLevel = Transactions.IsolationLevel.ReadCommitted
            opt.Timeout = TimeSpan.FromMinutes(5)
            Using scope As New TransactionScope(Transactions.TransactionScopeOption.Required, opt)
                For i = 0 To recExp.Count - 1
                    Try
                        Try
                            insertDetail(recExp(i), cfg.ConnectionString, errMsg)
                            insertDetailBox(recExp(i), cfg.ConnectionString, errMsg)
                            insertMaster(recExp(i), cfg.ConnectionString, errMsg)                            
                        Catch ex As Exception
                            errMsg = ex.Message.ToString
                        End Try
                        If errMsg <> "" Then
                            If Not IsNothing(ExcelBook) Then
                                ExcelBook.Save()
                                xlApp.Workbooks.Close()
                                xlApp.Quit()
                            End If
                            pub_ErrorMessage = errMsg
                            sendEmailtoTOS(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                            log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                            log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                            LogName.Refresh()
                            Exit Try
                        End If
                    Catch ex As Exception
                        errMsg = ex.Message.ToString
                    End Try
                Next
                scope.Complete()
            End Using
        Catch ex As Exception
            If Not IsNothing(ExcelBook) Then
                ExcelBook.Save()
                xlApp.Workbooks.Close()
                xlApp.Quit()
            End If

            pub_ErrorMessage = "File Excel Corrupt. Please check this file again!"
            sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

            log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & ex.Message)

            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & ex.Message, LogName)
            LogName.Refresh()
            Exit Try
        Finally
            If Not xlApp Is Nothing Then
                xlApp.DisplayAlerts = False
                clsGeneral.NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                clsGeneral.NAR(ExcelBook)
                xlApp.Quit()
                clsGeneral.NAR(xlApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If

            If pub_ErrorMessage <> "" Then
                clsFilterEmail.up_MoveErrorFile(pAtttacment & "\", pResultExp & "\BACKUP ERROR FILE" & "\", pFileName)
            Else
                clsFilterEmail.up_MoveFile(pAtttacment & "\", pResultExp & "\", pFileName)
                sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.SuratJalanNo, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)
            End If
        End Try

    End Sub

    Public Shared Function CekDataDetailDNSupplier(ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSupplierID As String, _
                                                   ByVal pAffiliateID As String, ByVal PartNo As String, ByVal pSuratJalanNo As String, _
                                                   ByVal pConstr As String, ByRef errMsg As String) As Integer
        Dim sql As String = ""

        sql = "SELECT AffiliateID, MAX(SeqNo) + 1 SeqNo FROM DOSupplier_Detail_Export ds " & vbCrLf

        sql = sql + " WHERE PONo = '" & Trim(pPONo) & "' and AffiliateID ='" & Trim(pAffiliateID) & "' and SupplierID ='" & Trim(pSupplierID) & "' and OrderNo = '" & Trim(pOrderNo) & "'" & vbCrLf & _
                    " And PartNo = '" & Trim(PartNo) & "' and SuratJalanNo = '" & Trim(pSuratJalanNo) & "'" & vbCrLf & _
                    " GROUP BY AffiliateID"
        
        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    CekDataDetailDNSupplier = dt.Tables(0).Rows(0)("SeqNo")
                Else
                    CekDataDetailDNSupplier = 1
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Public Shared Function cekSJSubmited(ByVal PONo As String, ByVal OrderNo As String, _
                                           ByVal SuratJalanNo As String, _
                                           ByVal SupplierID As String, ByVal AffiliateID As String, _
                                           ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = "select * from DOSupplier_Master_Export WHERE SuratjalanNo = '" & Trim(SuratJalanNo) & "' and SupplierID = '" & Trim(SupplierID) & "' AND AffiliateID = '" & Trim(AffiliateID) & "' " & vbCrLf & _
                  " AND OrderNo = '" & Trim(OrderNo) & "' AND PONo = '" & Trim(PONo) & "'"
        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    cekSJSubmited = True
                Else
                    cekSJSubmited = False
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Public Shared Function cekPartNoAndPONo(ByVal pMaster As clsReceivingExportProperty, ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = " select PartNo from PO_Detail_Export " & vbCrLf & _
                  " where PONo = '" & pMaster.PONo & "' and OrderNo1 = '" & pMaster.OrderNo & "' and AffiliateID = '" & pMaster.AffiliateID & "' " & vbCrLf & _
                  "       and SupplierID = '" & pMaster.SupplierID & "' and PartNo = '" & pMaster.PartNo & "' " & vbCrLf

        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    cekPartNoAndPONo = True
                Else
                    cekPartNoAndPONo = False
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Public Shared Function cekPartNoAndLabelRegister(ByVal PONo As String, ByVal OrderNo As String, _
                                            ByVal PartNo As String, ByVal pBoxNo As String, ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = "select * from PrintLabelExport " & vbCrLf & _
              "where PONo = '" & Trim(PONo) & "' and OrderNo = '" & Trim(OrderNo) & "' and PartNo = '" & Trim(PartNo) & "' and LabelNo = '" & Trim(pBoxNo) & "' "
        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    cekPartNoAndLabelRegister = True
                Else
                    cekPartNoAndLabelRegister = False
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Public Shared Function cekPartNoAndLabel(ByVal SuratJalanNo As String, ByVal PONo As String, ByVal OrderNo As String, _
                                            ByVal PartNo As String, ByVal pBoxNo As String, _
                                            ByVal AffiliateID As String, ByVal SupplierID As String, ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""


        sql = "SELECT * FROM DOSupplier_DetailBox_Export ds " & vbCrLf

        sql = sql + " WHERE PONo = '" & Trim(PONo) & "' and AffiliateID ='" & Trim(AffiliateID) & "' and SupplierID ='" & Trim(SupplierID) & "' and OrderNo = '" & Trim(OrderNo) & "'" & vbCrLf & _
                    " And PartNo = '" & Trim(PartNo) & "' and BoxNo = '" & Trim(pBoxNo) & "' and not exists " & vbCrLf & _
                    "(" & vbCrLf & _
                    "   select * from ReceiveForwarder_DetailBox a " & vbCrLf & _
                    "   WHERE a.PartNo = ds.PartNo and (DS.BoxNo between a.Label1 and a.Label2) and StatusDefect = '1'" & vbCrLf & _
                    ")" & vbCrLf
        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    cekPartNoAndLabel = True
                    errMsg = dt.Tables(0).Rows(0)("SuratJalanNo")
                Else
                    cekPartNoAndLabel = False
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Public Shared Function UpdateLabelPrint_RecEX(ByVal pMaster As clsReceivingExportProperty, ByVal pConstr As String, ByRef errMsg As String) As Integer
        Dim sql As String = ""
        Dim i As Integer

        Try
            Dim xstatus As String
            Dim ds As New DataSet

            Using cn As New SqlConnection(pConstr)
                cn.Open()

                If pMaster.TotalBoxG > 0 Then xstatus = "0" Else xstatus = "1"

                sql = " Update PrintLabelExport SET SuratJalanNo_FWD = '" & Trim(pMaster.SuratJalanNo) & "', StatusDefect = '" & xstatus & "' " & vbCrLf & _
                      " WHERE SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PONo = '" & Trim(pMaster.PONo) & "' " & vbCrLf & _
                      " AND OrderNo = '" & Trim(pMaster.OrderNo) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "' and (LabelNo between '" & Trim(pMaster.BoxNoFrom) & "' and '" & Trim(pMaster.BoxNoTo) & "') "

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Function

    Shared Sub insertMaster(ByVal pMaster As clsDNExportProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()
                sql = " UPDATE dbo.DOSupplier_Master_Export " & vbCrLf & _
                        " SET UpdateDate = getdate(), " & vbCrLf & _
                        " 	  UpdateUser = 'AdmUpload' " & vbCrLf & _
                        " WHERE SuratJalanNo = '" & Trim(pMaster.SuratJalanNo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' " & vbCrLf & _
                        " AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND OrderNo = '" & Trim(pMaster.OrderNo) & "'" & vbCrLf & _
                        " AND PONo = '" & Trim(pMaster.PONo) & "'"

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO dbo.DOSupplier_Master_Export " & vbCrLf & _
                           " (suratjalanno, supplierID, AffiliateID, PONo, OrderNo, EntryDate, EntryUser, UpdateDate, UpdateUser, ExcelCls, DeliveryDate, CommercialCls) " & vbCrLf & _
                           " VALUES  ( '" & Trim(pMaster.SuratJalanNo) & "' ," & vbCrLf & _
                           "           '" & Trim(pMaster.SupplierID) & "' ," & vbCrLf & _
                           "           '" & Trim(pMaster.AffiliateID) & "' ," & vbCrLf & _
                           "           '" & Trim(pMaster.PONo) & "' ," & vbCrLf & _
                           "           '" & Trim(pMaster.OrderNo) & "' ," & vbCrLf

                    sql = sql + "           GETDATE() ,  " & vbCrLf & _
                                "           'AdmUpload' ,  " & vbCrLf & _
                                "           GETDATE() , " & vbCrLf & _
                                "           'AdmUpload', '1', Getdate(), '" & pMaster.CommercialCls & "' " & vbCrLf & _
                                "         ) "
                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Sub insertDetail(ByVal pMaster As clsDNExportProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()

                sql = " UPDATE dbo.DOSUpplier_Detail_Export " & vbCrLf & _
                        " SET DOQty = " & Trim(pMaster.DOQty) & " " & vbCrLf & _
                        " WHERE SuratJalanNo = '" & Trim(pMaster.SuratJalanNo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' " & vbCrLf & _
                        " AND OrderNo = '" & Trim(pMaster.OrderNo) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "' " & vbCrLf & _
                        " AND PONo = '" & Trim(pMaster.PONo) & "' and SeqNo = '" & pMaster.SeqNo & "'"

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO dbo.DOSUpplier_Detail_Export " & vbCrLf & _
                            " (SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, SeqNo, DOQty) " & vbCrLf & _
                            " VALUES  ( '" & Trim(pMaster.SuratJalanNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pMaster.SupplierID) & "' , " & vbCrLf & _
                            "           '" & Trim(pMaster.AffiliateID) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pMaster.PONo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pMaster.PartNo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pMaster.OrderNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pMaster.SeqNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pMaster.DOQty) & "'" & vbCrLf & _
                            "       ) "
                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Sub insertDetailBox(ByVal pMaster As clsDNExportProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Dim ds As New DataSet

            Using cn As New SqlConnection(pConstr)
                cn.Open()

                Dim i_rec1 As Integer = Microsoft.VisualBasic.Right(Trim(pMaster.BoxNoFrom), 7)
                Dim i_rec2 As Integer = Microsoft.VisualBasic.Right(Trim(pMaster.BoxNoTo), 7)
                Dim i_PrefixLabelNo As String = Microsoft.VisualBasic.Left(Trim(pMaster.BoxNoFrom), 2)

                Dim cmd As New SqlCommand
                cmd.Connection = cn

                For i_rec1 = i_rec1 To i_rec2
                    Dim pBoxNo As String = i_PrefixLabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7)
                    sql = " UPDATE dbo.DOSUpplier_DetailBox_Export " & vbCrLf & _
                            " SET BoxNo = '" & Trim(pBoxNo) & "' " & vbCrLf & _
                            " WHERE SuratJalanNo = '" & Trim(pMaster.SuratJalanNo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PONo = '" & Trim(pMaster.PONo) & "' " & vbCrLf & _
                            " AND OrderNo = '" & Trim(pMaster.OrderNo) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "' " & vbCrLf & _
                            " AND PONo = '" & Trim(pMaster.PONo) & "'" & vbCrLf & _
                            " AND BoxNo = '" & Trim(pBoxNo) & "' and SeqNo = '" & pMaster.SeqNo & "'"
                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery

                    If i = 0 Then
                        sql = " INSERT INTO dbo.DOSUpplier_DetailBox_Export " & vbCrLf & _
                               " (SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, BoxNo, SeqNo) " & vbCrLf & _
                               " VALUES  ( '" & Trim(pMaster.SuratJalanNo) & "' , " & vbCrLf & _
                               "           '" & Trim(pMaster.SupplierID) & "' , " & vbCrLf & _
                               "           '" & Trim(pMaster.AffiliateID) & "' ,  " & vbCrLf & _
                               "           '" & Trim(pMaster.PONo) & "' ,  " & vbCrLf & _
                               "           '" & Trim(pMaster.PartNo) & "' ,  " & vbCrLf & _
                               "           '" & Trim(pMaster.OrderNo) & "' , " & vbCrLf & _
                               "           '" & Trim(pBoxNo) & "'," & vbCrLf & _
                               "           '" & Trim(pMaster.SeqNo) & "'" & vbCrLf & _
                               "       ) "

                        cmd.CommandText = sql
                        i = cmd.ExecuteNonQuery()
                    End If
                Next

            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Sub deleteMaster(ByVal pMaster As clsReceivingExportProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()
                sql = " Delete dbo.receiveForwarder_master " & vbCrLf & _
                      " WHERE SuratJalanNo = '" & Trim(pMaster.SuratJalanNo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' " & vbCrLf & _
                      " AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND OrderNo = '" & Trim(pMaster.OrderNo) & "' AND PONo = '" & Trim(pMaster.PONo) & "'" & vbCrLf
                sql = sql + " Delete dbo.receiveForwarder_Detail " & vbCrLf & _
                            " WHERE SuratJalanNo = '" & Trim(pMaster.SuratJalanNo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' " & vbCrLf & _
                            " AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND OrderNo = '" & Trim(pMaster.OrderNo) & "' AND PONo = '" & Trim(pMaster.PONo) & "'" & vbCrLf
                sql = sql + " Delete dbo.ReceiveForwarder_DetailBox " & vbCrLf & _
                            " WHERE SuratJalanNo = '" & Trim(pMaster.SuratJalanNo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' " & vbCrLf & _
                            " AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND OrderNo = '" & Trim(pMaster.OrderNo) & "' AND PONo = '" & Trim(pMaster.PONo) & "'" & vbCrLf

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pSuratJalanNo As String, ByVal pOrderNo As String, ByVal pAffiliate As String, ByVal pSupplierID As String, ByVal pErrorMsg As String, ByRef errMsg As String) As Boolean
        Dim dsEmail As New DataSet

        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoSupplier = True

            dsEmail = clsGeneral.getEmailAddressSupplier(GB, "PASI", pSupplierID, "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                End If
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "SUPP" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "SUPP" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("EmailCC")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailCC")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If fromEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Send Surat Jalan No. [" & pSuratJalanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplierID & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Send Surat Jalan No, [" & pSuratJalanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplierID & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If pErrorMsg = "" Then
                ls_Subject = "FEEDBACK SUCCESS - DN Surat Jalan No: " & pSuratJalanNo & "-" & pOrderNo & "-" & pAffiliate
                ls_Attachment = ""
                ls_Body = clsNotification.GetNotification("102", , , , pSuratJalanNo, , , , pErrorMsg)
            Else
                ls_Subject = "FEEDBACK FAILED - DN Surat Jalan No: " & pSuratJalanNo & "-" & pOrderNo & "-" & pAffiliate
                ls_Attachment = Trim(pPathFile) & "\" & pFileName
                ls_Body = clsNotification.GetNotification("103", , , , pSuratJalanNo, , , , pErrorMsg)
            End If


            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Send Shipping Instruction No [" & pSuratJalanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplierID & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Function sendEmailtoTOS(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pSuratJalanNo As String, ByVal pOrderNo As String, ByVal pAffiliate As String, ByVal pSupplierID As String, ByVal pErrorMsg As String, ByRef errMsg As String) As Boolean
        Dim dsEmail As New DataSet

        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoTOS = True

            dsEmail = clsGeneral.getEmailAddressSupplier(GB, "PASI", pSupplierID, "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                End If
                'If dsEmail.Tables(0).Rows(iRow)("FLAG") = "FWD" Then
                '    If receiptEmail = "" Then
                '        receiptEmail = dsEmail.Tables(0).Rows(iRow)("EmailTO")
                '    Else
                '        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailTO")
                '    End If
                'End If
                'If dsEmail.Tables(0).Rows(iRow)("FLAG") = "FWD" Then
                '    If receiptCCEmail = "" Then
                '        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("EmailCC")
                '    Else
                '        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailCC")
                '    End If
                'End If
            Next

            receiptEmail = "yudha@tos.co.id;jemmy@tos.co.id;edi@tos.co.id"
            receiptCCEmail = "yudha@tos.co.id;jemmy@tos.co.id;edi@tos.co.id"

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If fromEmail = "" Then
                sendEmailtoTOS = False
                errMsg = "Process Send Surat JalanNo No. [" & pSuratJalanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplierID & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoTOS = False
                errMsg = "Process Send Surat JalanNo No, [" & pSuratJalanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplierID & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If pErrorMsg = "" Then
                ls_Subject = "FEEDBACK SUCCESS - DN Surat Jalan No: " & pSuratJalanNo & "-" & pOrderNo & "-" & pAffiliate
                ls_Attachment = ""
                ls_Body = clsNotification.GetNotification("100", , , , pSuratJalanNo, , , , pErrorMsg)
            Else
                ls_Subject = "FEEDBACK FAILED - DN Surat Jalan No: " & pSuratJalanNo & "-" & pOrderNo & "-" & pAffiliate
                ls_Attachment = Trim(pPathFile) & "\" & pFileName
                ls_Body = clsNotification.GetNotification("101", , , , pSuratJalanNo, , , , pErrorMsg)
            End If


            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoTOS = False
                Exit Function
            End If

            sendEmailtoTOS = True

        Catch ex As Exception
            sendEmailtoTOS = False
            errMsg = "Process Send Shipping Instruction No [" & pSuratJalanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplierID & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function
End Class
