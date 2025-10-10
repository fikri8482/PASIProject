Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Transactions

Public Class clsPO
    Shared Sub up_PODom(ByVal cfg As GlobalSetting.clsConfig,
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

        Dim cls As New clsPOProperty
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application

        Dim sheetNumber As Integer = 1
        Dim checkHeader As Boolean = True

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
                cls.PONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I9").Value.ToString & "", 20)
            End If

            If ExcelSheet.Range("J25").Value Is Nothing Then
                cls.PIC = ""
            Else
                cls.PIC = Microsoft.VisualBasic.Left(ExcelSheet.Range("J25").Value.ToString & "", 20)
            End If

            If ExcelSheet.Range("J27").Value Is Nothing Then
                cls.Remarks = ""
            Else
                cls.Remarks = Microsoft.VisualBasic.Left(ExcelSheet.Range("J27").Value.ToString & "", 100)
            End If

            'Check PO already Upload?
            If cekAutoApprove(cls.PONo, cls.SupplierID, cls.AffiliateID, cfg.ConnectionString, errMsg) = True Then
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

            Dim recExp As New List(Of clsPOProperty)
            Dim startRow As Integer = "36"

            For i = startRow To 10000
                If ExcelSheet.Range("B" & i).Value.ToString = "E" Then
                    Exit For
                End If

                Dim setNilai As New clsPOProperty
                setNilai.SupplierID = cls.SupplierID.Trim
                setNilai.AffiliateID = cls.AffiliateID.Trim            
                setNilai.PONo = cls.PONo.Trim

                setNilai.PartNo = Trim(ExcelSheet.Range("D" & i).Value.ToString & "")
                setNilai.QtyBox = clsGeneral.getQtyBox(GB, setNilai.PartNo, setNilai.SupplierID, setNilai.AffiliateID)

                setNilai.QtyPOOld = clsGeneral.getQtyDomestic(GB, setNilai.PartNo, setNilai.SupplierID, setNilai.AffiliateID, setNilai.PONo)

                If cekPartNoAndPONoRegister(cls.PONo, setNilai.PartNo, cls.SupplierID, cls.AffiliateID, cfg.ConnectionString, errMsg) = False Then
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

                'get new Qty Approval
                Try
                    If ExcelSheet.Range("AX" & i + 1).Value Is Nothing = False Then setNilai.Qty01 = ExcelSheet.Range("AX" & i + 1).Value
                    If ExcelSheet.Range("AZ" & i + 1).Value Is Nothing = False Then setNilai.Qty02 = ExcelSheet.Range("AZ" & i + 1).Value
                    If ExcelSheet.Range("BB" & i + 1).Value Is Nothing = False Then setNilai.Qty03 = ExcelSheet.Range("BB" & i + 1).Value
                    If ExcelSheet.Range("BD" & i + 1).Value Is Nothing = False Then setNilai.Qty04 = ExcelSheet.Range("BD" & i + 1).Value
                    If ExcelSheet.Range("BF" & i + 1).Value Is Nothing = False Then setNilai.Qty05 = ExcelSheet.Range("BF" & i + 1).Value
                    If ExcelSheet.Range("BH" & i + 1).Value Is Nothing = False Then setNilai.Qty06 = ExcelSheet.Range("BH" & i + 1).Value
                    If ExcelSheet.Range("BJ" & i + 1).Value Is Nothing = False Then setNilai.Qty07 = ExcelSheet.Range("BJ" & i + 1).Value
                    If ExcelSheet.Range("BL" & i + 1).Value Is Nothing = False Then setNilai.Qty08 = ExcelSheet.Range("BL" & i + 1).Value
                    If ExcelSheet.Range("BN" & i + 1).Value Is Nothing = False Then setNilai.Qty09 = ExcelSheet.Range("BN" & i + 1).Value
                    If ExcelSheet.Range("BP" & i + 1).Value Is Nothing = False Then setNilai.Qty10 = ExcelSheet.Range("BP" & i + 1).Value

                    If ExcelSheet.Range("BR" & i + 1).Value Is Nothing = False Then setNilai.Qty11 = ExcelSheet.Range("BR" & i + 1).Value
                    If ExcelSheet.Range("BT" & i + 1).Value Is Nothing = False Then setNilai.Qty12 = ExcelSheet.Range("BT" & i + 1).Value
                    If ExcelSheet.Range("BV" & i + 1).Value Is Nothing = False Then setNilai.Qty13 = ExcelSheet.Range("BV" & i + 1).Value
                    If ExcelSheet.Range("BX" & i + 1).Value Is Nothing = False Then setNilai.Qty14 = ExcelSheet.Range("BX" & i + 1).Value
                    If ExcelSheet.Range("BZ" & i + 1).Value Is Nothing = False Then setNilai.Qty15 = ExcelSheet.Range("BZ" & i + 1).Value
                    If ExcelSheet.Range("CB" & i + 1).Value Is Nothing = False Then setNilai.Qty16 = ExcelSheet.Range("CB" & i + 1).Value
                    If ExcelSheet.Range("CD" & i + 1).Value Is Nothing = False Then setNilai.Qty17 = ExcelSheet.Range("CD" & i + 1).Value
                    If ExcelSheet.Range("CF" & i + 1).Value Is Nothing = False Then setNilai.Qty18 = ExcelSheet.Range("CF" & i + 1).Value
                    If ExcelSheet.Range("CH" & i + 1).Value Is Nothing = False Then setNilai.Qty19 = ExcelSheet.Range("CH" & i + 1).Value
                    If ExcelSheet.Range("CJ" & i + 1).Value Is Nothing = False Then setNilai.Qty20 = ExcelSheet.Range("CJ" & i + 1).Value

                    If ExcelSheet.Range("CL" & i + 1).Value Is Nothing = False Then setNilai.Qty21 = ExcelSheet.Range("CL" & i + 1).Value
                    If ExcelSheet.Range("CN" & i + 1).Value Is Nothing = False Then setNilai.Qty22 = ExcelSheet.Range("CN" & i + 1).Value
                    If ExcelSheet.Range("CP" & i + 1).Value Is Nothing = False Then setNilai.Qty23 = ExcelSheet.Range("CP" & i + 1).Value
                    If ExcelSheet.Range("CR" & i + 1).Value Is Nothing = False Then setNilai.Qty24 = ExcelSheet.Range("CR" & i + 1).Value
                    If ExcelSheet.Range("CT" & i + 1).Value Is Nothing = False Then setNilai.Qty25 = ExcelSheet.Range("CT" & i + 1).Value
                    If ExcelSheet.Range("CV" & i + 1).Value Is Nothing = False Then setNilai.Qty26 = ExcelSheet.Range("CV" & i + 1).Value
                    If ExcelSheet.Range("CX" & i + 1).Value Is Nothing = False Then setNilai.Qty27 = ExcelSheet.Range("CX" & i + 1).Value
                    If ExcelSheet.Range("CZ" & i + 1).Value Is Nothing = False Then setNilai.Qty28 = ExcelSheet.Range("CZ" & i + 1).Value
                    If ExcelSheet.Range("DB" & i + 1).Value Is Nothing = False Then setNilai.Qty29 = ExcelSheet.Range("DB" & i + 1).Value
                    If ExcelSheet.Range("DD" & i + 1).Value Is Nothing = False Then setNilai.Qty30 = ExcelSheet.Range("DD" & i + 1).Value
                    If ExcelSheet.Range("DF" & i + 1).Value Is Nothing = False Then setNilai.Qty31 = ExcelSheet.Range("DF" & i + 1).Value

                    setNilai.QtyPO = setNilai.Qty01 + setNilai.Qty02 + setNilai.Qty03 + setNilai.Qty04 + setNilai.Qty05 _
                                   + setNilai.Qty06 + setNilai.Qty07 + setNilai.Qty08 + setNilai.Qty09 + setNilai.Qty10 _
                                   + setNilai.Qty11 + setNilai.Qty12 + setNilai.Qty13 + setNilai.Qty14 + setNilai.Qty15 _
                                   + setNilai.Qty16 + setNilai.Qty17 + setNilai.Qty18 + setNilai.Qty19 + setNilai.Qty20 _
                                   + setNilai.Qty21 + setNilai.Qty22 + setNilai.Qty23 + setNilai.Qty24 + setNilai.Qty25 _
                                   + setNilai.Qty26 + setNilai.Qty27 + setNilai.Qty28 + setNilai.Qty29 + setNilai.Qty30 + setNilai.Qty31
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

                'Check Approve Qty Mod Qty/Box
                Try
                    If (setNilai.Qty01 Mod setNilai.QtyBox <> 0 And setNilai.Qty11 Mod setNilai.QtyBox <> 0 And setNilai.Qty21 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty02 Mod setNilai.QtyBox <> 0 And setNilai.Qty12 Mod setNilai.QtyBox <> 0 And setNilai.Qty22 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty03 Mod setNilai.QtyBox <> 0 And setNilai.Qty13 Mod setNilai.QtyBox <> 0 And setNilai.Qty23 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty04 Mod setNilai.QtyBox <> 0 And setNilai.Qty14 Mod setNilai.QtyBox <> 0 And setNilai.Qty24 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty05 Mod setNilai.QtyBox <> 0 And setNilai.Qty15 Mod setNilai.QtyBox <> 0 And setNilai.Qty25 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty06 Mod setNilai.QtyBox <> 0 And setNilai.Qty16 Mod setNilai.QtyBox <> 0 And setNilai.Qty26 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty07 Mod setNilai.QtyBox <> 0 And setNilai.Qty17 Mod setNilai.QtyBox <> 0 And setNilai.Qty27 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty08 Mod setNilai.QtyBox <> 0 And setNilai.Qty18 Mod setNilai.QtyBox <> 0 And setNilai.Qty28 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty09 Mod setNilai.QtyBox <> 0 And setNilai.Qty19 Mod setNilai.QtyBox <> 0 And setNilai.Qty29 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty10 Mod setNilai.QtyBox <> 0 And setNilai.Qty20 Mod setNilai.QtyBox <> 0 And setNilai.Qty30 Mod setNilai.QtyBox <> 0 And _
                        setNilai.Qty31 Mod setNilai.QtyBox <> 0) Then

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
                i = i + 1
            Next


            Dim opt As Transactions.TransactionOptions
            opt.IsolationLevel = Transactions.IsolationLevel.ReadCommitted
            opt.Timeout = TimeSpan.FromMinutes(5)
            Using scope As New TransactionScope(Transactions.TransactionScopeOption.Required, opt)
                For i = 0 To recExp.Count - 1
                    Try
                        Try
                            insertDetail(recExp(i), cfg.ConnectionString, errMsg)
                            insertMaster(recExp(i), cfg.ConnectionString, errMsg)
                            UpdateMaster(recExp(i), cfg.ConnectionString, errMsg)
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
                clsFilterEmail.up_MoveErrorFile(pAtttacment & "\", pResultDom & "\BACKUP ERROR FILE" & "\", pFileName)
            Else
                clsFilterEmail.up_MoveFile(pAtttacment & "\", pResultDom & "\", pFileName)
                sendEmailtoSupplier(GB, pAtttacment, pFileName, cls.PONo, cls.AffiliateID, cls.SupplierID, pub_ErrorMessage, errMsg)
            End If
        End Try
    End Sub

    Public Shared Function cekPartNoAndPONoRegister(ByVal PONo As String, ByVal PartNo As String, _
                                            ByVal SupplierID As String, ByVal AffiliateID As String, _
                                            ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = "select * from PO_Detail " & vbCrLf & _
              " WHERE SupplierID = '" & Trim(SupplierID) & "' AND AffiliateID = '" & Trim(AffiliateID) & "' and PartNo = '" & Trim(PartNo) & "'" & vbCrLf & _
              " and PONo = '" & Trim(PONo) & "' "
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

    Public Shared Function cekAutoApprove(ByVal PONo As String, _
                                            ByVal SupplierID As String, ByVal AffiliateID As String, _
                                            ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = "select * from PO_Master " & vbCrLf & _
              " WHERE SupplierID = '" & Trim(SupplierID) & "' AND AffiliateID = '" & Trim(AffiliateID) & "' " & vbCrLf & _
              " AND PONo = '" & Trim(PONo) & "' and (SupplierApproveUser = 'AUTO APPROVED' or SupplierApprovePendingUser = 'AUTO APPROVED' or SupplierUnApproveUser = 'AUTO APPROVED')"
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

    Shared Sub insertMaster(ByVal pMaster As clsPOProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()
                sql = " UPDATE dbo.PO_MasterUpload " & vbCrLf & _
                      " SET UpdateDate = getdate(), " & vbCrLf & _
                      " 	UpdateUser = '" & pMaster.PIC & "' " & vbCrLf & _
                      " WHERE SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' " & vbCrLf & _
                      " AND PONo = '" & Trim(pMaster.PONo) & "' "

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO dbo.PO_MasterUpload " & vbCrLf & _
                            " (PONo, AffiliateID, SupplierID, Remarks, EntryDate, EntryUser) " & vbCrLf & _
                            " VALUES  ( '" & Trim(pMaster.PONo) & "' ," & vbCrLf & _
                            "           '" & Trim(pMaster.AffiliateID) & "' ," & vbCrLf & _
                            "           '" & Trim(pMaster.SupplierID) & "' ," & vbCrLf & _
                            "           '" & Trim(pMaster.Remarks) & "', " & vbCrLf 
                    sql = sql + "           GETDATE() , " & vbCrLf & _
                                "           '" & Trim(pMaster.PIC) & "' " & vbCrLf & _
                                "         ) "

                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Sub insertDetail(ByVal pMaster As clsPOProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()

                sql = " UPDATE dbo.PO_DetailUpload " & vbCrLf & _
                      " SET UpdateDate = getdate(), " & vbCrLf & _
                      " 	UpdateUser = '" & pMaster.PIC & "' " & vbCrLf & _
                      " WHERE SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' " & vbCrLf & _
                      " AND PONo = '" & Trim(pMaster.PONo) & "'"

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO [PO_DetailUpload] " & vbCrLf & _
                            "            ([PONo],[AffiliateID],[SupplierID],[PartNo] " & vbCrLf & _
                            "            ,[POQty],[POQtyOld] " & vbCrLf & _
                            " 		     ,[DeliveryD1],[DeliveryD1Old],[DeliveryD2],[DeliveryD2Old],[DeliveryD3],[DeliveryD3Old],[DeliveryD4],[DeliveryD4Old],[DeliveryD5],[DeliveryD5Old] " & vbCrLf & _
                            "            ,[DeliveryD6],[DeliveryD6Old],[DeliveryD7],[DeliveryD7Old],[DeliveryD8],[DeliveryD8Old],[DeliveryD9],[DeliveryD9Old],[DeliveryD10],[DeliveryD10Old] " & vbCrLf & _
                            "            ,[DeliveryD11],[DeliveryD11Old],[DeliveryD12],[DeliveryD12Old],[DeliveryD13],[DeliveryD13Old],[DeliveryD14],[DeliveryD14Old],[DeliveryD15],[DeliveryD15Old] " & vbCrLf & _
                            "            ,[DeliveryD16],[DeliveryD16Old],[DeliveryD17],[DeliveryD17Old],[DeliveryD18],[DeliveryD18Old],[DeliveryD19],[DeliveryD19Old],[DeliveryD20],[DeliveryD20Old] " & vbCrLf & _
                            "            ,[DeliveryD21],[DeliveryD21Old],[DeliveryD22],[DeliveryD22Old],[DeliveryD23],[DeliveryD23Old],[DeliveryD24],[DeliveryD24Old],[DeliveryD25],[DeliveryD25Old] " & vbCrLf & _
                            "            ,[DeliveryD26],[DeliveryD26Old],[DeliveryD27],[DeliveryD27Old],[DeliveryD28],[DeliveryD28Old],[DeliveryD29],[DeliveryD29Old],[DeliveryD30],[DeliveryD30Old] " & vbCrLf & _
                            "            ,[DeliveryD31],[DeliveryD31Old],[EntryDate],[EntryUser] " & vbCrLf & _
                            "            ) " & vbCrLf & _
                            " VALUES  (" & vbCrLf

                    sql = sql + "           '" & Trim(pMaster.PONo) & "' , " & vbCrLf & _
                                "           '" & Trim(pMaster.AffiliateID) & "' , " & vbCrLf & _
                                "           '" & Trim(pMaster.SupplierID) & "' ,  " & vbCrLf & _
                                "           '" & Trim(pMaster.PartNo) & "' ,  " & vbCrLf & _
                                "           '" & Trim(pMaster.QtyPO) & "' ,  " & vbCrLf & _
                                "           '" & Trim(pMaster.QtyPOOld) & "' , " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty01) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD1 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty02) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD2 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty03) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD3 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty04) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD4 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty05) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD5 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty06) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD6 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty07) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD7 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty08) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD8 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty09) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD9 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty10) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD10 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf

                    sql = sql + "           '" & Trim(pMaster.Qty11) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD11 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty12) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD12 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty13) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD13 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty14) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD14 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty15) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD15 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty16) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD16 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty17) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD17 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty18) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD18 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty19) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD19 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty20) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD20 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf

                    sql = sql + "           '" & Trim(pMaster.Qty21) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD21 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty22) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD22 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty23) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD23 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty24) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD24 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty25) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD25 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty26) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD26 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty27) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD27 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty28) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD28 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty29) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD29 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty30) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD30 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf & _
                                "           '" & Trim(pMaster.Qty31) & "',  " & vbCrLf & _
                                "           (SELECt DeliveryD31 FROM PO_Detail WHERE PONo = '" & Trim(pMaster.PONo) & "' AND SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PartNo = '" & Trim(pMaster.PartNo) & "'),  " & vbCrLf

                    sql = sql + "           GETDATE() , " & vbCrLf & _
                                "           '" & pMaster.PIC & "' " & vbCrLf & _
                                "         ) "

                   
                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Sub UpdateMaster(ByVal pMaster As clsPOProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()
                sql = " UPDATE dbo.PO_Master " & vbCrLf & _
                      " SET SupplierApproveDate = GETDATE(), " & vbCrLf & _
                      " SupplierApproveUser = 'AdmUpload' " & vbCrLf & _
                      " WHERE SupplierID = '" & Trim(pMaster.SupplierID) & "' AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' " & vbCrLf & _
                      " and PONo = '" & Trim(pMaster.PONo) & "' "

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery
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

            dsEmail = clsGeneral.getEmailAddressSupplierDOM(GB, "PASI", pSupplier, "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

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
                ls_Body = clsNotification.GetNotification("108", , pOrderNo, , , , , , pErrorMsg)
            Else
                ls_Subject = "FEEDBACK FAILED - APPROVAL PO No: " & pOrderNo & "-" & pSupplier & "-" & pAffiliate
                ls_Attachment = Trim(pPathFile) & "\" & pFileName
                ls_Body = clsNotification.GetNotification("109", , pOrderNo, , , , , , pErrorMsg)
            End If


            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
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

            dsEmail = clsGeneral.getEmailAddressSupplierDOM(GB, "PASI", pSupplier, "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                End If
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
                ls_Body = clsNotification.GetNotification("108", , pOrderNo, , , , , , pErrorMsg)
            Else
                ls_Subject = "FEEDBACK FAILED - APPROVAL PO No: " & pOrderNo & "-" & pSupplier & "-" & pAffiliate
                ls_Attachment = Trim(pPathFile) & "\" & pFileName
                ls_Body = clsNotification.GetNotification("109", , pOrderNo, , , , , , pErrorMsg)
            End If


            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
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
