Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Transactions

Public Class clsPOMonthly
    Shared Sub up_POMonthly(ByVal cfg As GlobalSetting.clsConfig,
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

        Dim cls As New clsPOEmergencyProperty
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application

        Dim sheetNumber As Integer = 1
        Dim checkHeader As Boolean = True

        Dim tempPONo As String
        Dim tempOrderNo As String

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

            If ExcelSheet.Range("I9").Value Is Nothing Then
                '03. Move to Error Folder
                If Not IsNothing(ExcelBook) Then
                    ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
                pub_ErrorMessage = "PO No. Blank, Please check this file again!"
                sendEmailtoSupplier(GB, pAtttacment, pFileName, "", cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                LogName.Refresh()
                Exit Try
            Else
                tempPONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I9").Value.ToString & "", 20)
            End If

            If ExcelSheet.Range("I11").Value Is Nothing Then
                tempOrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I9").Value.ToString & "", 20)
            Else
                tempOrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I11").Value.ToString & "", 20)
            End If

            If tempPONo <> tempOrderNo Then
                cls.PONo = tempOrderNo.Trim
                cls.OrderNo = tempPONo.Trim
            Else
                cls.PONo = tempPONo.Trim
                cls.OrderNo = tempOrderNo.Trim
            End If

            Try
                cls.ETDVendor = Format(clsGeneral.getEmergencyETD(GB, cls.PONo, cls.OrderNo, cls.SupplierID, cls.AffiliateID), "yyyy-MM-dd")
            Catch ex As Exception
                cls.ETDVendor = Now
            End Try

            'Check PO already Upload?
            If cekAutoApprove(cls.PONo, cls.OrderNo, cls.DeliveryLocation, cls.SupplierID, cls.AffiliateID, cfg.ConnectionString, errMsg) = True Then
                If Not IsNothing(ExcelBook) Then
                    ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
                pub_ErrorMessage = "PO No. [" & cls.PONo & "], Affiliate [" & cls.AffiliateID & "] already upload in PASI system, Please check this file again!"
                sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

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

            Dim recExp As New List(Of clsPOMonthlyProperty)
            Dim startRow As Integer = "37"

            For i = startRow To 10000
                If ExcelSheet.Range("B" & i).Value.ToString = "E" Then
                    Exit For
                End If

                Dim setNilai As New clsPOMonthlyProperty
                setNilai.SupplierID = cls.SupplierID.Trim
                setNilai.AffiliateID = cls.AffiliateID.Trim
                setNilai.DeliveryLocation = cls.DeliveryLocation.Trim
                setNilai.PONo = cls.PONo.Trim
                setNilai.OrderNo = cls.OrderNo.Trim

                setNilai.PartNo = Trim(ExcelSheet.Range("D" & i).Value.ToString & "")
                setNilai.QtyBox = clsGeneral.getQtyBox(GB, setNilai.PartNo, setNilai.SupplierID, setNilai.AffiliateID)
                setNilai.Week = clsGeneral.getQtyWeek(GB, setNilai.PartNo, setNilai.SupplierID, setNilai.AffiliateID, setNilai.PONo, setNilai.OrderNo)

                If cekPartNoAndPONoRegister(cls.PONo, cls.OrderNo, setNilai.PartNo, setNilai.DeliveryLocation, cls.SupplierID, cls.AffiliateID, cfg.ConnectionString, errMsg) = False Then
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "ROW [" & i & "], PartNo [" & setNilai.PartNo & "] not found in PASI System, Please check this file again!"
                    sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                End If

                'Check 1-5 ETD FORWARDER and Week
                Try
                    If ExcelSheet.Range("AE" & i).Value Is Nothing = False Then
                        setNilai.Week1 = ExcelSheet.Range("AE" & i).Value
                    End If

                    If ExcelSheet.Range("AO" & i).Value Is Nothing = False Then
                        setNilai.Week2 = ExcelSheet.Range("AO" & i).Value
                    End If

                    If ExcelSheet.Range("AY" & i).Value Is Nothing = False Then
                        setNilai.Week3 = ExcelSheet.Range("AY" & i).Value
                    End If

                    If ExcelSheet.Range("BI" & i).Value Is Nothing = False Then
                        setNilai.Week4 = ExcelSheet.Range("BI" & i).Value
                    End If

                    If ExcelSheet.Range("BS" & i).Value Is Nothing = False Then
                        setNilai.Week5 = ExcelSheet.Range("BS" & i).Value
                    End If


                    If ExcelSheet.Range("Z" & i).Value Is Nothing = False Then
                        setNilai.ETDVendor1 = ExcelSheet.Range("Z" & i).Value
                    End If

                    If ExcelSheet.Range("AJ" & i).Value Is Nothing = False Then
                        setNilai.ETDVendor2 = ExcelSheet.Range("AJ" & i).Value
                    End If

                    If ExcelSheet.Range("AT" & i).Value Is Nothing = False Then
                        setNilai.ETDVendor3 = ExcelSheet.Range("AT" & i).Value
                    End If

                    If ExcelSheet.Range("BD" & i).Value Is Nothing = False Then
                        setNilai.ETDVendor4 = ExcelSheet.Range("BD" & i).Value
                    End If

                    If ExcelSheet.Range("BN" & i).Value Is Nothing = False Then
                        setNilai.ETDVendor5 = ExcelSheet.Range("BN" & i).Value
                    End If

                Catch ex As Exception
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "ROW [" & i & "], PartNo [" & setNilai.PartNo & "] , Please check this file again! " & ex.Message
                    sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                End Try

                'Check Approve Qty OLD <> 0
                Try
                    'setNilai.we = clsGeneral.getApproveQtyPO(GB, setNilai.PONo, setNilai.OrderNo, setNilai.SupplierID, setNilai.AffiliateID, setNilai.PartNo)
                    Dim Qty1 As Double = IIf(IsNothing(setNilai.Week1), 0, setNilai.Week1)
                    Dim Qty2 As Double = IIf(IsNothing(setNilai.Week2), 0, setNilai.Week2)
                    Dim Qty3 As Double = IIf(IsNothing(setNilai.Week3), 0, setNilai.Week3)
                    Dim Qty4 As Double = IIf(IsNothing(setNilai.Week4), 0, setNilai.Week4)
                    Dim Qty5 As Double = IIf(IsNothing(setNilai.Week5), 0, setNilai.Week5)

                    If setNilai.Week <> (Qty1 + Qty2 + Qty3 + Qty4 + Qty5) Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Qty Approve is different with Qty Order, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

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
                    pub_ErrorMessage = "ROW [" & i & "] " & ex.Message & ", Please check this file again!"
                    sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                End Try

                'Check Approve Qty Mod Qty/Box
                Try
                    If Not IsNothing(setNilai.Week1) And setNilai.Week1 Mod setNilai.QtyBox <> 0 Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Approve Qty must be multiply from Qty/Box, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                    If Not IsNothing(setNilai.Week2) And setNilai.Week2 Mod setNilai.QtyBox <> 0 Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Approve Qty must be multiply from Qty/Box, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                    If Not IsNothing(setNilai.Week3) And setNilai.Week3 Mod setNilai.QtyBox <> 0 Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Approve Qty must be multiply from Qty/Box, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                    If Not IsNothing(setNilai.Week4) And setNilai.Week4 Mod setNilai.QtyBox <> 0 Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Approve Qty must be multiply from Qty/Box, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                    If Not IsNothing(setNilai.Week5) And setNilai.Week5 Mod setNilai.QtyBox <> 0 Then
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Approve Qty must be multiply from Qty/Box, Please check this file again!"
                        sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

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
                    pub_ErrorMessage = "ROW [" & i & "] Approve Qty must be multiply from Qty/Box, Please check this file again!"
                    sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                End Try

                recExp.Add(setNilai)
            Next


            Dim opt As Transactions.TransactionOptions
            opt.IsolationLevel = Transactions.IsolationLevel.ReadCommitted
            opt.Timeout = TimeSpan.FromMinutes(5)
            Using scope As New TransactionScope(Transactions.TransactionScopeOption.Required, opt)
                For i = 0 To recExp.Count - 1
                    Try
                        Try
                            'insertDetail(recExp(i), cfg.ConnectionString, errMsg)
                            'insertMaster(recExp(i), cfg.ConnectionString, errMsg)
                            'UpdateMaster(recExp(i), cfg.ConnectionString, errMsg)
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
                            sendEmailtoTOS(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

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
            sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)

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
                sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)
            End If
        End Try

    End Sub

    Public Shared Function cekPartNoAndPONoRegister(ByVal PONo As String, ByVal OrderNo As String, _
                                            ByVal PartNo As String, ByVal DeliveryLocation As String, _
                                            ByVal SupplierID As String, ByVal AffiliateID As String, _
                                            ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = "select * from PO_Detail_Export " & vbCrLf & _
              " WHERE SupplierID = '" & Trim(SupplierID) & "' AND AffiliateID = '" & Trim(AffiliateID) & "' and PartNo = '" & Trim(PartNo) & "'" & vbCrLf & _
              " AND OrderNo1 = '" & Trim(OrderNo) & "' and PONo = '" & Trim(PONo) & "' and ForwarderID = '" & DeliveryLocation & "'"
        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    cekPartNoAndPONoRegister = True
                Else
                    cekPartNoAndPONoRegister = False
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Public Shared Function cekAutoApprove(ByVal PONo As String, ByVal OrderNo As String, _
                                            ByVal DeliveryLocation As String, _
                                            ByVal SupplierID As String, ByVal AffiliateID As String, _
                                            ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = "select * from PO_Master_Export " & vbCrLf & _
              " WHERE SupplierID = '" & Trim(SupplierID) & "' AND AffiliateID = '" & Trim(AffiliateID) & "' " & vbCrLf & _
              " AND OrderNo1 = '" & Trim(OrderNo) & "' and PONo = '" & Trim(PONo) & "' and ForwarderID = '" & DeliveryLocation & "' and SupplierApproveUser = 'AUTO APPROVED'"
        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    cekAutoApprove = True
                Else
                    cekAutoApprove = False
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Shared Sub insertMaster(ByVal pMaster As clsPOEmergencyProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()
                sql = " UPDATE dbo.PO_MasterUpload_Export " & vbCrLf & _
                      " SET UpdateDate = getdate(), " & vbCrLf & _
                      " 	UpdateUser = 'AdmUpload' " & vbCrLf & _
                      " WHERE SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' " & vbCrLf & _
                      " AND OrderNo1 = '" & Trim(pMaster.OrderNo) & "' and PONo = '" & Trim(pMaster.PONo) & "' and ForwarderID = '" & pMaster.DeliveryLocation & "'"

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO dbo.PO_MasterUpload_Export " & vbCrLf & _
                            " (PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1, ETDVendor1, EntryDate, EntryUser) " & vbCrLf & _
                            " VALUES  ( '" & Trim(pMaster.PONo) & "' ," & vbCrLf & _
                            "           '" & Trim(pMaster.AffiliateID) & "' ," & vbCrLf & _
                            "           '" & Trim(pMaster.SupplierID) & "' ," & vbCrLf & _
                            "           '" & Trim(pMaster.DeliveryLocation) & "', " & vbCrLf & _
                            "           '" & Trim(pMaster.OrderNo) & "', " & vbCrLf & _
                            "           '" & Trim(pMaster.ETDVendor) & "', " & vbCrLf
                    sql = sql + "           GETDATE() , " & vbCrLf & _
                                "           'AdmUpload' " & vbCrLf & _
                                "         ) "

                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Sub UpdateMaster(ByVal pMaster As clsPOEmergencyProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()
                sql = " UPDATE dbo.PO_Master_Export " & vbCrLf & _
                      " SET SupplierApproveDate = GETDATE(), " & vbCrLf & _
                      " SupplierApproveUser = 'AdmUpload' " & vbCrLf & _
                      " WHERE SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' " & vbCrLf & _
                      " AND OrderNo1 = '" & Trim(pMaster.OrderNo) & "' and PONo = '" & Trim(pMaster.PONo) & "' and ForwarderID = '" & pMaster.DeliveryLocation & "'"

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Sub insertDetail(ByVal pMaster As clsPOEmergencyProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()

                sql = " UPDATE dbo.PO_DetailUpload_Export " & vbCrLf & _
                      " SET UpdateDate = getdate(), " & vbCrLf & _
                      " 	UpdateUser = 'AdmUpload' " & vbCrLf & _
                      " WHERE SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' " & vbCrLf & _
                      " AND OrderNo1 = '" & Trim(pMaster.OrderNo) & "' and PONo = '" & Trim(pMaster.PONo) & "' and ForwarderID = '" & pMaster.DeliveryLocation & "'"

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO dbo.PO_DetailUpload_Export " & vbCrLf & _
                            " (PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1, PartNo, Week1, Week1Old, TotalPOQty, TotalPOQtyOld, EntryDate, EntryUser) " & vbCrLf & _
                            " VALUES  ( '" & Trim(pMaster.PONo) & "' , " & vbCrLf & _
                            "           '" & Trim(pMaster.AffiliateID) & "' , " & vbCrLf & _
                            "           '" & Trim(pMaster.SupplierID) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pMaster.DeliveryLocation) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pMaster.OrderNo) & "' ,  " & vbCrLf & _
                            "           '" & Trim(pMaster.PartNo) & "' , " & vbCrLf & _
                            "           '" & Trim(pMaster.ApproveQty) & "',  " & vbCrLf & _
                            "           '" & Trim(pMaster.ApproveQtyOld) & "',  " & vbCrLf & _
                            "           '" & Trim(pMaster.ApproveQty) & "',  " & vbCrLf & _
                            "           '" & Trim(pMaster.ApproveQtyOld) & "',  " & vbCrLf

                    sql = sql + "           GETDATE() , " & vbCrLf & _
                                "           'AdmUpload' " & vbCrLf & _
                                "         ) "
                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pOrderNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pErrorMsg As String, ByRef errMsg As String) As Boolean
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

            dsEmail = clsGeneral.getEmailAddressSupplier(GB, "PASI", pSupplier, "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

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
                errMsg = "Process Send PO No. [" & pOrderNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Send PO No, [" & pOrderNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If pErrorMsg = "" Then
                ls_Subject = "FEEDBACK SUCCESS - APPROVAL PO No: " & pOrderNo & "-" & pSupplier & "-" & pAffiliate
                ls_Attachment = ""
                ls_Body = clsNotification.GetNotification("106", , pOrderNo, , , , , , pErrorMsg)
            Else
                ls_Subject = "FEEDBACK FAILED - APPROVAL PO No: " & pOrderNo & "-" & pSupplier & "-" & pAffiliate
                ls_Attachment = Trim(pPathFile) & "\" & pFileName
                ls_Body = clsNotification.GetNotification("107", , pOrderNo, , , , , , pErrorMsg)
            End If


            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Send PO No [" & pOrderNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Function sendEmailtoTOS(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pOrderNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pErrorMsg As String, ByRef errMsg As String) As Boolean
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

            dsEmail = clsGeneral.getEmailAddressSupplier(GB, "PASI", pSupplier, "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)

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
                errMsg = "Process Send PO No. [" & pOrderNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoTOS = False
                errMsg = "Process Send PO No. [" & pOrderNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If pErrorMsg = "" Then
                ls_Subject = "FEEDBACK SUCCESS - APPROVAL PO No: " & pOrderNo & "-" & pSupplier & "-" & pAffiliate
                ls_Attachment = ""
                ls_Body = clsNotification.GetNotification("106", , pOrderNo, , , , , , pErrorMsg)
            Else
                ls_Subject = "FEEDBACK FAILED - APPROVAL PO No: " & pOrderNo & "-" & pSupplier & "-" & pAffiliate
                ls_Attachment = Trim(pPathFile) & "\" & pFileName
                ls_Body = clsNotification.GetNotification("107", , pOrderNo, , , , , , pErrorMsg)
            End If


            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoTOS = False
                Exit Function
            End If

            sendEmailtoTOS = True

        Catch ex As Exception
            sendEmailtoTOS = False
            errMsg = "Process Send PO No [" & pOrderNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

End Class
