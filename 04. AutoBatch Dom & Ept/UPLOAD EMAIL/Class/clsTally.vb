Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Transactions

Public Class clsTally
    Shared Sub up_Tally(ByVal cfg As GlobalSetting.clsConfig,
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

        Dim cls As New clsTallyProperty
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application

        Dim sheetNumber As Integer = 1
        Dim totalSheet As Integer = 1
        Dim k As Integer = 1
        Dim checkHeader As Boolean = True
        Dim checkMaster As Boolean = True
        Dim checkDetail As Boolean = True

        Try
            Dim ls_file As String = pAtttacment & "\" & pFileName
            ExcelBook = xlApp.Workbooks.Open(ls_file)
            totalSheet = ExcelBook.Worksheets.Count

            For k = 1 To totalSheet
                checkDetail = True
                checkMaster = True
                checkHeader = True

                pub_ErrorMessage = ""

                ExcelSheet = CType(ExcelBook.Worksheets(k), Excel.Worksheet)

                log.WriteToProcessLog(Date.Now, pScreenName, "Read Sheet [" & k & "], FileName [" & pFileName & "]")

                If ExcelSheet.Range("H3").Value Is Nothing Then
                    checkHeader = False
                    pub_ErrorMessage = "Header Template tidak sesuai (Consignee Code kosong). Silahkan dicek kembali Template yang disubmit!"
                    GoTo step001
                Else
                    cls.AffiliateID = Trim(ExcelSheet.Range("H3").Value.ToString & "")
                    cls.AffiliateID = clsGeneral.AffiliateConsignee(GB, cls.AffiliateID)
                End If

                If ExcelSheet.Range("H4").Value Is Nothing Then
                    checkHeader = False
                    pub_ErrorMessage = "Header Template tidak sesuai (Delivery Location kosong). Silahkan dicek kembali Template yang disubmit!"
                    GoTo step001
                Else
                    cls.DeliveryLocation = Trim(ExcelSheet.Range("H4").Value.ToString & "")
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

                log.WriteToProcessLog(Date.Now, pScreenName, "Read Header [" & k & "], FileName [" & pFileName & "]")


                '####01. Read Tally Master
                If ExcelSheet.Range("I8").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Invoice No. Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, "", cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.ShippingInstructionNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I8").Value.ToString.Trim & "", 20)
                End If

                If ExcelSheet.Range("I10").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Container No. Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.ContainerNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I10").Value.ToString.Trim & "", 30)
                End If

                If ExcelSheet.Range("I12").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Seal No. Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.SealNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I12").Value.ToString & "", 30)
                End If

                If ExcelSheet.Range("I14").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Tare Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    If IsNumeric(ExcelSheet.Range("I14").Value) = False Then
                        '03. Move to Error Folder
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "Tare must be numeric, Please check this file again!"
                        sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                    cls.Tare = ExcelSheet.Range("I14").Value
                End If

                If ExcelSheet.Range("I16").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Gross Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    If IsNumeric(ExcelSheet.Range("I16").Value.ToString) = False Then
                        '03. Move to Error Folder
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "Gross must be numeric, Please check this file again!"
                        sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                    cls.Gross = ExcelSheet.Range("I16").Value
                End If

                If ExcelSheet.Range("I18").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Total Carton Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    If IsNumeric(ExcelSheet.Range("I18").Value.ToString) = False Then
                        '03. Move to Error Folder
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "Total Carton must be numeric, Please check this file again!"
                        sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If
                    cls.TotalCarton = ExcelSheet.Range("I18").Value
                End If


                If ExcelSheet.Range("AA8").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Vessel Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.Vessel = Microsoft.VisualBasic.Left(ExcelSheet.Range("AA8").Value.ToString.Trim & "", 50)
                End If

                If ExcelSheet.Range("AA10").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Size Container Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.SizeContainer = Microsoft.VisualBasic.Left(ExcelSheet.Range("AA10").Value.ToString.Trim & "", 30)
                End If

                If ExcelSheet.Range("AA12").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "ETD PORT Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.ETDPort = Microsoft.VisualBasic.Left(ExcelSheet.Range("AA12").Value.ToString & "", 30)
                End If

                If ExcelSheet.Range("AA14").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Shipping Line Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.ShippingLine = Microsoft.VisualBasic.Left(ExcelSheet.Range("AA14").Value.ToString & "", 30)
                End If

                If ExcelSheet.Range("AA16").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Destination Port Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.DestinationPort = Microsoft.VisualBasic.Left(ExcelSheet.Range("AA16").Value.ToString & "", 30)
                End If

                If ExcelSheet.Range("AQ8").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Vessel Name Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.VesselName = Microsoft.VisualBasic.Left(ExcelSheet.Range("AQ8").Value.ToString & "", 30)
                End If

                If ExcelSheet.Range("AQ10").Value Is Nothing Then
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "Stuffing Date Blank, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                Else
                    cls.StuffingDate = Microsoft.VisualBasic.Left(ExcelSheet.Range("AQ10").Value.ToString & "", 30)
                End If

                'Check SI and Container already Upload?
                If cekShippingIntructionNo(cls.ShippingInstructionNo, cls.AffiliateID, cls.ContainerNo, cfg.ConnectionString, errMsg) = True Then
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    pub_ErrorMessage = "SI No. [" & cls.ShippingInstructionNo & "], Container No. [" & cls.ContainerNo & "] already exists in PASI system, Please check this file again!"
                    sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                    log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                    LogName.Refresh()
                    Exit Try
                End If

                'Refresh Log
                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Upload File [" & pFileName & "]", LogName)
                LogName.Refresh()

                log.WriteToProcessLog(Date.Now, pScreenName, "Read Detail [" & k & "], FileName [" & pFileName & "]")
                Dim recExp As New List(Of clsTallyProperty)
                Dim startRow As Integer = "23"
                Dim temp_PalletNo As String = ""

                Dim temp_Length As Double = 0
                Dim temp_Width As Double = 0
                Dim temp_Height As Double = 0
                Dim temp_M3 As Double = 0
                Dim temp_Weight As Double = 0

                For i = startRow To 10000
                    If ExcelSheet.Range("B" & i).Value.ToString = "E" Then
                        Exit For
                    End If

                    Dim setNilai As New clsTallyProperty

                    'Check Pallet Kosong
                    Try
                        temp_PalletNo = IIf(IsNothing(ExcelSheet.Range("D" & i).Value.ToString & ""), "", ExcelSheet.Range("D" & i).Value).ToString & ""
                        setNilai.PalletNo = temp_PalletNo
                    Catch ex As Exception
                        setNilai.PalletNo = temp_PalletNo
                    End Try

                    'Check Pallet Pertama Kosong
                    If setNilai.PalletNo = "" Then
                        '03. Move to Error Folder
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Pallet No Blank, Please check this file again!"
                        sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End If


                    'Check PO tidak boleh kosong
                    Try
                        setNilai.PONo = ExcelSheet.Range("I" & i).Value.ToString & ""
                    Catch ex As Exception
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Order No Blank, Please check this file again!"
                        sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End Try


                    'Check Part No tidak boleh kosong
                    Try
                        setNilai.PartNo = Trim(ExcelSheet.Range("O" & i).Value.ToString & "")
                    Catch ex As Exception
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If
                        pub_ErrorMessage = "ROW [" & i & "] Part No Blank, Please check this file again!"
                        sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End Try

                    'Check Box No
                    Try
                        setNilai.BoxNoFrom = Trim(Trim(ExcelSheet.Range("AE" & i).Value))
                        setNilai.BoxNoTo = Trim(Trim(ExcelSheet.Range("AK" & i).Value))
                        If setNilai.BoxNoTo = "" Then
                            setNilai.BoxNoTo = setNilai.BoxNoFrom
                        End If
                    Catch ex As Exception
                        setNilai.BoxNoTo = setNilai.BoxNoFrom
                    End Try

                    'Check Prefix Box NO From and TO must be same
                    Try
                        'setNilai.PartNo = Trim(ExcelSheet.Range("O" & i).Value.ToString & "")
                        If Microsoft.VisualBasic.Left(Trim(setNilai.BoxNoFrom), 2) <> Microsoft.VisualBasic.Left(Trim(setNilai.BoxNoTo), 2) Then
                            If Not IsNothing(ExcelBook) Then
                                ExcelBook.Save()
                                xlApp.Workbooks.Close()
                                xlApp.Quit()
                            End If
                            pub_ErrorMessage = "ROW [" & i & "] Prefix BoxNo different, Please check this file again!"
                            sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

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
                        sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                        log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                        LogName.Refresh()
                        Exit Try
                    End Try

                    'Validasi buat check data Part No, Order No, dan Label must be in Shipping Instruction
                    Dim i_rec1 As Integer = Microsoft.VisualBasic.Right(Trim(setNilai.BoxNoFrom), 7)
                    Dim i_rec2 As Integer = Microsoft.VisualBasic.Right(Trim(setNilai.BoxNoTo), 7)
                    Dim i_PrefixLabelNo As String = Microsoft.VisualBasic.Left(Trim(setNilai.BoxNoFrom), 2)

                    setNilai.TotalBox = (i_rec2 - i_rec1) + 1

                    For i_rec1 = i_rec1 To i_rec2
                        '2.1 check
                        If cekPartNoAndLabel(cls.ShippingInstructionNo, cls.AffiliateID, setNilai.PONo, setNilai.PartNo, i_PrefixLabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7), cfg.ConnectionString, errMsg) = False Then
                            If Not IsNothing(ExcelBook) Then
                                ExcelBook.Save()
                                xlApp.Workbooks.Close()
                                xlApp.Quit()
                            End If
                            pub_ErrorMessage = "ROW [" & i & "], Box NO [" & i_PrefixLabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7) & "] BoxNo not found with Shipping Intruction, Please check this file again!"
                            sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                            log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                            log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                            LogName.Refresh()
                            Exit Try
                        End If
                    Next

                    Try
                        temp_Length = IIf(IsNothing(ExcelSheet.Range("AQ" & i).Value), "", ExcelSheet.Range("AQ" & i).Value).ToString & ""
                        setNilai.Length = temp_Length
                    Catch ex As Exception
                        setNilai.Length = temp_Length
                    End Try

                    Try
                        temp_Width = IIf(IsNothing(ExcelSheet.Range("AT" & i).Value), "", ExcelSheet.Range("AT" & i).Value).ToString & ""
                        setNilai.Width = temp_Width
                    Catch ex As Exception
                        setNilai.Width = temp_Width
                    End Try

                    Try
                        temp_Height = IIf(IsNothing(ExcelSheet.Range("AW" & i).Value), "", ExcelSheet.Range("AW" & i).Value).ToString & ""
                        setNilai.Height = temp_Height
                    Catch ex As Exception
                        setNilai.Height = temp_Height
                    End Try

                    Try
                        temp_M3 = IIf(IsNothing(ExcelSheet.Range("AZ" & i).Value), "", ExcelSheet.Range("AZ" & i).Value).ToString & ""
                        setNilai.M3 = temp_M3
                    Catch ex As Exception
                        setNilai.M3 = temp_M3
                    End Try

                    Try
                        temp_Weight = IIf(IsNothing(ExcelSheet.Range("BC" & i).Value), "", ExcelSheet.Range("BC" & i).Value).ToString & ""
                        setNilai.WeightPallet = temp_Weight
                    Catch ex As Exception
                        setNilai.WeightPallet = temp_Weight
                    End Try

                    setNilai.ShippingInstructionNo = cls.ShippingInstructionNo
                    setNilai.AffiliateID = cls.AffiliateID
                    setNilai.DeliveryLocation = cls.DeliveryLocation
                    setNilai.ContainerNo = cls.ContainerNo
                    setNilai.SealNo = cls.SealNo
                    setNilai.Tare = cls.Tare
                    setNilai.Gross = cls.Gross
                    setNilai.TotalCarton = cls.TotalCarton
                    setNilai.Vessel = cls.Vessel
                    setNilai.SizeContainer = cls.SizeContainer
                    setNilai.ETDPort = cls.ETDPort
                    setNilai.ShippingLine = cls.ShippingLine
                    setNilai.DestinationPort = cls.DestinationPort
                    setNilai.VesselName = cls.VesselName
                    setNilai.StuffingDate = cls.StuffingDate


                    recExp.Add(setNilai)
                Next

                Dim opt As Transactions.TransactionOptions
                opt.IsolationLevel = Transactions.IsolationLevel.ReadCommitted
                opt.Timeout = TimeSpan.FromMinutes(5)
                Using scope As New TransactionScope(Transactions.TransactionScopeOption.Required, opt)
                    For i = 0 To recExp.Count - 1
                        Try
                            insertDetail(recExp(i), cfg.ConnectionString, errMsg)
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
                            sendEmailtoTOS(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

                            log.WriteToErrorLog(pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                            log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage)

                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & pFileName & "] Failed, because " & pub_ErrorMessage, LogName)
                            LogName.Refresh()
                            Exit Try
                        End If
                    Next                  

                    scope.Complete()
                End Using
            Next
        Catch ex As Exception
            If Not IsNothing(ExcelBook) Then
                ExcelBook.Save()
                xlApp.Workbooks.Close()
                xlApp.Quit()
            End If

            pub_ErrorMessage = "File Excel Corrupt. Please check this file again!"
            sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)

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
                sendEmailtoForwarder(GB, pAtttacment, pFileName, cls.ShippingInstructionNo, cls.AffiliateID, cls.DeliveryLocation, pub_ErrorMessage, errMsg)
            End If
        End Try
    End Sub

    Shared Sub insertDetail(ByVal pMaster As clsTallyProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()
                sql = " UPDATE dbo.Tally_Detail " & vbCrLf & _
                      " SET UpdateDate = getdate(), " & vbCrLf & _
                      " 	UpdateUser = 'AdmUpload' " & vbCrLf & _
                      " WHERE ShippingInstructionNo = '" & Trim(pMaster.ShippingInstructionNo) & "' AND ForwarderID = '" & Trim(pMaster.DeliveryLocation) & "' " & vbCrLf & _
                      " AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND PalletNo = '" & Trim(pMaster.PalletNo) & "' and OrderNo = '" & Trim(pMaster.PONo) & "'" & vbCrLf & _
                      " AND PartNo = '" & Trim(pMaster.PartNo) & "' AND CaseNo = '" & Trim(pMaster.BoxNoFrom) & "' and CaseNo2 = '" & Trim(pMaster.BoxNoTo) & "'" & vbCrLf & _
                      " AND ContainerNo = '" & pMaster.ContainerNo & "'"

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO Tally_Detail " & vbCrLf & _
                          " (ShippingInstructionNo,ForwarderID,AffiliateID, PalletNo, OrderNo, PartNo, " & vbCrLf & _
                          " CaseNo, Length, Width, Height, M3, WeightPallet, caseNo2, TotalBox, ContainerNo ) " & vbCrLf & _
                          "  "
                    sql = sql + " VALUES  ( '" & Trim(pMaster.ShippingInstructionNo) & "','" & Trim(pMaster.DeliveryLocation) & "' , " & vbCrLf & _
                                "           '" & Trim(pMaster.AffiliateID) & "' , '" & Trim(pMaster.PalletNo) & "', " & vbCrLf & _
                                "           '" & Trim(pMaster.PONo) & "' , '" & Trim(pMaster.PartNo) & "' , " & vbCrLf & _
                                "           '" & Trim(pMaster.BoxNoFrom) & "', '" & Trim(pMaster.Length) & "'," & vbCrLf & _
                                "           '" & Trim(pMaster.Width) & "', '" & Trim(pMaster.Height) & "'," & vbCrLf & _
                                "           '" & Trim(pMaster.M3) & "', '" & Trim(pMaster.WeightPallet) & "', " & vbCrLf & _
                                "           '" & Trim(pMaster.BoxNoTo) & "'," & vbCrLf & _
                                "           " & Trim(pMaster.TotalBox) & ", " & vbCrLf & _
                                "           '" & Trim(pMaster.ContainerNo) & "' " & vbCrLf & _
                                "           )" & vbCrLf

                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Sub insertMaster(ByVal pMaster As clsTallyProperty, ByVal pConstr As String, ByRef errMsg As String)
        Dim sql As String = ""
        Dim i As Integer

        Try
            Using cn As New SqlConnection(pConstr)
                cn.Open()
                sql = " UPDATE dbo.Tally_Master " & vbCrLf & _
                      " SET UpdateDate = getdate(), " & vbCrLf & _
                      " 	UpdateUser = 'AdmUpload' " & vbCrLf & _
                      " WHERE ShippingInstructionNo = '" & Trim(pMaster.ShippingInstructionNo) & "' AND ForwarderID = '" & Trim(pMaster.DeliveryLocation) & "' " & vbCrLf & _
                      "       AND AffiliateID = '" & Trim(pMaster.AffiliateID) & "' AND ContainerNo = '" & Trim(pMaster.ContainerNo) & "' " & vbCrLf

                Dim cmd As New SqlCommand(sql, cn)
                i = cmd.ExecuteNonQuery

                If i = 0 Then
                    sql = " INSERT INTO Tally_Master " & vbCrLf & _
                      " (ShippingInstructionNo,ForwarderID,AffiliateID, ContainerNo,SealNo,Tare, " & vbCrLf & _
                      " Gross, TotalCarton, Vessel, ContainerSize, ETD, ShippingLine, DestinationPort, NamaKapal, StuffingDate, TallyCls2) " & vbCrLf & _
                      "  "
                    sql = sql + " VALUES  ( '" & Trim(pMaster.ShippingInstructionNo) & "','" & Trim(pMaster.DeliveryLocation) & "' , " & vbCrLf & _
                                "           '" & Trim(pMaster.AffiliateID) & "' , '" & Trim(pMaster.ContainerNo) & "', " & vbCrLf & _
                                "           '" & Trim(pMaster.SealNo) & "' , '" & Trim(pMaster.Tare) & "' , " & vbCrLf & _
                                "           '" & Trim(pMaster.Gross) & "', '" & Trim(pMaster.TotalCarton) & "'," & vbCrLf & _
                                "           '" & Trim(pMaster.Vessel) & "', " & vbCrLf & _
                                "           '" & Trim(pMaster.SizeContainer) & "', '" & Trim(pMaster.ETDPort) & "', " & vbCrLf & _
                                "           '" & Trim(pMaster.ShippingLine) & "', '" & Trim(pMaster.DestinationPort) & "'," & vbCrLf & _
                                "           '" & Trim(pMaster.VesselName) & "', " & vbCrLf & _
                                "           " & IIf(pMaster.StuffingDate = "", "NULL", "'" & pMaster.StuffingDate & "'") & ",'1')" & vbCrLf


                    cmd.CommandText = sql
                    i = cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As SqlException
            errMsg = ex.Message.ToString
        End Try
    End Sub

    Shared Function cekPartNoAndLabel(ByVal ShippingInstructionNo As String, _
                                            ByVal AffiliateID As String, ByVal PONo As String, _
                                            ByVal PartNo As String, ByVal pBoxNo As String, ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = " select  * from ShippingInstruction_Detail " & vbCrLf & _
              "     where ShippingInstructionNo = '" & ShippingInstructionNo & "' " & vbCrLf & _
              "        and AffiliateID = '" & AffiliateID & "' " & vbCrLf & _
              "        and OrderNo = '" & PONo & "' and PartNo = '" & PartNo & "' " & vbCrLf & _
              "        and ('" & pBoxNo & "' between SUBSTRING(BoxNo,1,9) and SUBSTRING(BoxNo,11,9)) " & vbCrLf

        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    cekPartNoAndLabel = True
                Else
                    cekPartNoAndLabel = False
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Shared Function cekShippingIntructionNo(ByVal ShippingInstructionNo As String, _
                                            ByVal AffiliateID As String, ByVal ContainerNo As String, ByVal pConstr As String, ByRef errMsg As String) As Boolean
        Dim sql As String = ""

        sql = " select  * from Tally_Master " & vbCrLf & _
              "     where ShippingInstructionNo = '" & ShippingInstructionNo & "' " & vbCrLf & _
              "        and AffiliateID = '" & AffiliateID & "' " & vbCrLf & _
              "        and ContainerNo = '" & ContainerNo & "' " & vbCrLf

        Try
            Using Cn As New SqlConnection(pConstr)
                Cn.Open()
                Dim cmd As New SqlCommand(sql, Cn)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataSet

                da.Fill(dt)

                If dt.Tables(0).Rows.Count > 0 Then
                    cekShippingIntructionNo = True
                Else
                    cekShippingIntructionNo = False
                End If
            End Using
        Catch ex As Exception
            errMsg = ex.Message.ToString
        End Try
    End Function

    Shared Function sendEmailtoForwarder(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pShippingInstructionNo As String, ByVal pAffiliate As String, ByVal pForwarderID As String, ByVal pErrorMsg As String, ByRef errMsg As String) As Boolean
        Dim dsEmail As New DataSet

        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoForwarder = True

            dsEmail = clsGeneral.getEmailAddressForwarder(GB, "PASI", pForwarderID, "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                End If
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "FWD" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "FWD" Then
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
                sendEmailtoForwarder = False
                errMsg = "Process Send Shipping Instruction No. [" & pShippingInstructionNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarderID & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoForwarder = False
                errMsg = "Process Send Shipping Instruction No, [" & pShippingInstructionNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarderID & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If pErrorMsg = "" Then
                ls_Subject = "FEEDBACK SUCCESS - TALLY DATA Shipping Instruction: " & pShippingInstructionNo & "-" & pAffiliate
                ls_Attachment = ""
                ls_Body = clsNotification.GetNotification("104", , , , , , , pShippingInstructionNo, pErrorMsg)
            Else
                ls_Subject = "FEEDBACK FAILED - TALLY DATA Shipping Instruction: " & pShippingInstructionNo & "-" & pAffiliate
                ls_Attachment = Trim(pPathFile) & "\" & pFileName
                ls_Body = clsNotification.GetNotification("105", , , , , , , pShippingInstructionNo, pErrorMsg)
            End If
            

            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoForwarder = False
                Exit Function
            End If

            sendEmailtoForwarder = True

        Catch ex As Exception
            sendEmailtoForwarder = False
            errMsg = "Process Send Shipping Instruction No [" & pShippingInstructionNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarderID & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Function sendEmailtoTOS(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pShippingInstructionNo As String, ByVal pAffiliate As String, ByVal pForwarderID As String, ByVal pErrorMsg As String, ByRef errMsg As String) As Boolean
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

            dsEmail = clsGeneral.getEmailAddressForwarder(GB, "PASI", pForwarderID, "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)

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
                errMsg = "Process Send Shipping Instruction No. [" & pShippingInstructionNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarderID & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoTOS = False
                errMsg = "Process Send Shipping Instruction No, [" & pShippingInstructionNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarderID & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If pErrorMsg = "" Then
                ls_Subject = "FEEDBACK SUCCESS - TALLY DATA Shipping Instruction: " & pShippingInstructionNo & "-" & pAffiliate
                ls_Attachment = ""
                ls_Body = clsNotification.GetNotification("104", , , , , , , pShippingInstructionNo, pErrorMsg)
            Else
                ls_Subject = "FEEDBACK FAILED - TALLY DATA Shipping Instruction: " & pShippingInstructionNo & "-" & pAffiliate
                ls_Attachment = Trim(pPathFile) & "\" & pFileName
                ls_Body = clsNotification.GetNotification("105", , , , , , , pShippingInstructionNo, pErrorMsg)
            End If


            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoTOS = False
                Exit Function
            End If

            sendEmailtoTOS = True

        Catch ex As Exception
            sendEmailtoTOS = False
            errMsg = "Process Send Shipping Instruction No [" & pShippingInstructionNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarderID & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function
End Class
