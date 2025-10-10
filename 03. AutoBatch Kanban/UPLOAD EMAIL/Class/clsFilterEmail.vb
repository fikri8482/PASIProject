Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class clsFilterEmail
    Shared Sub up_FilterEmail(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResultDom As String,
                              ByVal pResultExp As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        'Dim ExcelBook As Excel.Workbook
        'Dim ExcelSheet As Excel.Worksheet
        'Dim xlApp = New Excel.Application

        Dim fi = From f In New IO.DirectoryInfo(pAtttacment).GetFiles().Cast(Of IO.FileInfo)() _
                  Where f.Extension = ".xlsm" OrElse f.Extension = ".xlsx"
                  Order By f.Name
                  Select f

        Dim fi1 As IO.FileInfo
        Dim excelName As String = ""
        Dim sheetNumber As Integer = 1

        Dim startRow As Integer = 0

        Dim lsOpen As Boolean = True

        For Each fi1 In fi
            Dim ExcelBook As Excel.Workbook
            Dim ExcelSheet As Excel.Worksheet
            Dim xlApp = New Excel.Application

            Dim ls_file As String = pAtttacment & "\" & fi1.Name
            excelName = fi1.Name

            Dim checkHeader As Boolean = True

            '00. Check bisa dibuka tidak excelnya
            Try
                lsOpen = True
                ExcelBook = xlApp.Workbooks.Open(ls_file)
                ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)
            Catch ex As Exception
                If Not IsNothing(ExcelBook) Then
                    ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
                pub_ErrorMessage = "File Excel Corrupt. Silahkan dicek kembali Template yang disubmit!"
                up_MoveErrorFile(pAtttacment & "\", pAtttacment & "\BACKUP ERROR FILE" & "\", excelName)
                lsOpen = False
            End Try

            Try
                If lsOpen = False Then GoTo step002

                '00. Can't Blank
                pub_TemplateCode = Trim(ExcelSheet.Range("H1").Value.ToString & "")

                If Not xlApp Is Nothing Then
                    xlApp.DisplayAlerts = False
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If

                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & fi1.Name & "]", LogName)
                LogName.Refresh()

                '01. Domestic
                If pub_TemplateCode = "KB" Or pub_TemplateCode = "DO" Or pub_TemplateCode = "DO (PO non Kanban)" _
                    Or pub_TemplateCode = "PO" Or pub_TemplateCode = "POR" Or pub_TemplateCode = "INV" Then
                    '01.1 CheckPO
                    If pub_TemplateCode = "PO" Then
                        clsPO.up_PODom(cfg, log, GB, LogName, pAtttacment, pResultDom, pResultDom, pScreenName, fi1.Name, errMsg, ErrSummary)
                    End If

                    '01.1 CheckDN Kanban
                    If pub_TemplateCode = "DO" Or pub_TemplateCode = "DO (PO non Kanban)" Then

                    End If

                    '01.1 CheckINV
                    If pub_TemplateCode = "INV" Then

                    End If
                ElseIf pub_TemplateCode = "REC-EX" Or pub_TemplateCode = "TALLY" Or pub_TemplateCode = "POEM" _
                    Or pub_TemplateCode = "POEE" Or pub_TemplateCode = "DO-EX" Or pub_TemplateCode = "INV-EX" Then

                    '02. Export OK
                    If pub_TemplateCode = "REC-EX" Then
                        clsReceivingExport.up_RecExport(cfg, log, GB, LogName, pAtttacment, pResultDom, pResultExp, pScreenName, fi1.Name, errMsg, ErrSummary)
                    End If

                    'OK
                    If pub_TemplateCode = "TALLY" Then                        
                        clsTally.up_Tally(cfg, log, GB, LogName, pAtttacment, pResultDom, pResultExp, pScreenName, fi1.Name, errMsg, ErrSummary)
                    End If

                    'OK
                    If pub_TemplateCode = "POEE" Then
                        clsPOEmergency.up_POEmergency(cfg, log, GB, LogName, pAtttacment, pResultDom, pResultExp, pScreenName, fi1.Name, errMsg, ErrSummary)
                    End If

                    'OK
                    If pub_TemplateCode = "DO-EX" Then
                        clsDNExport.up_DNExport(cfg, log, GB, LogName, pAtttacment, pResultDom, pResultExp, pScreenName, fi1.Name, errMsg, ErrSummary)
                    End If

                    'Check()
                    If pub_TemplateCode = "POEM" Then
                        clsPOMonthly.up_POMonthly(cfg, log, GB, LogName, pAtttacment, pResultDom, pResultExp, pScreenName, fi1.Name, errMsg, ErrSummary)
                    End If
                Else
                    '03. Move to Error Folder
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If

                    up_MoveErrorFile(pAtttacment & "\", pAtttacment & "\BACKUP ERROR FILE" & "\", excelName)

                    GoTo step003
                End If

step002:
                '00.1. Check Header Template, if error replay e-mail to supplier or Forwarder
                If checkHeader = False Then
                    '00.1.1 Send Email to Supplier or Forwarder

                End If


            Catch ex As Exception
                '03.1 Move to Error Folder
                'BACKUP ERROR FILE
                If Not IsNothing(ExcelBook) Then
                    ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If

                up_MoveErrorFile(pAtttacment & "\", pAtttacment & "\BACKUP ERROR FILE" & "\", excelName)

                log.WriteToErrorLog(pScreenName, "Process Read File [" & fi1.Name & "] Failed, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                log.WriteToProcessLog(Date.Now, pScreenName, "Process Read File [" & fi1.Name & "] Failed, because " & ex.Message)

                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Read File [" & fi1.Name & "] Failed, because " & ex.Message, LogName)
                LogName.Refresh()
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
            End Try
step003:
        Next
    End Sub

    Public Shared Sub up_MoveErrorFile(ByVal pPathSource As String, ByVal pPathDestination As String, ByVal excelName As String)
        Try
            If Not System.IO.Directory.Exists(pPathSource) Then
                System.IO.Directory.CreateDirectory(pPathSource)
            End If

            If Not System.IO.Directory.Exists(pPathDestination) Then
                System.IO.Directory.CreateDirectory(pPathDestination)
            End If

            My.Computer.FileSystem.MoveFile(pPathSource & excelName, pPathDestination & excelName, True)

        Catch ex As Exception
            
        End Try
    End Sub

    Public Shared Sub up_MoveFile(ByVal pPathSource As String, ByVal pPathDestination As String, ByVal excelName As String)
        Try
            If Not System.IO.Directory.Exists(pPathSource) Then
                System.IO.Directory.CreateDirectory(pPathSource)
            End If

            If Not System.IO.Directory.Exists(pPathDestination) Then
                System.IO.Directory.CreateDirectory(pPathDestination)
            End If

            My.Computer.FileSystem.MoveFile(pPathSource & excelName, pPathDestination & excelName, True)
        Catch ex As Exception
            
        End Try
    End Sub

    'Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pErrorMessage As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String) As Boolean
    '    Dim dsEmail As New DataSet

    '    Try
    '        Dim receiptEmail As String = ""
    '        Dim receiptCCEmail As String = ""
    '        Dim fromEmail As String = ""

    '        Dim ls_Subject As String = ""
    '        Dim ls_Body As String = ""
    '        Dim ls_Attachment As String = ""
    '        Dim ls_URl As String = ""

    '        sendEmailtoSupplier = True

    '        dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", pSupplier, "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

    '        For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '            If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
    '                fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
    '            End If
    '            If dsEmail.Tables(0).Rows(iRow)("FLAG") = "SUPP" Then
    '                If receiptEmail = "" Then
    '                    receiptEmail = dsEmail.Tables(0).Rows(iRow)("EmailTO")
    '                Else
    '                    receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailTO")
    '                End If
    '            End If
    '            If dsEmail.Tables(0).Rows(iRow)("FLAG") = "SUPP" Then
    '                If receiptCCEmail = "" Then
    '                    receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("EmailCC")
    '                Else
    '                    receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailCC")
    '                End If
    '            End If
    '        Next

    '        receiptCCEmail = Replace(receiptCCEmail, " ", "")
    '        receiptEmail = Replace(receiptEmail, " ", "")
    '        fromEmail = Replace(fromEmail, " ", "")

    '        receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '        receiptEmail = Replace(receiptEmail, ",", ";")

    '        If fromEmail = "" Then
    '            sendEmailtoSupplier = False
    '            errMsg = "Process Send PO [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
    '            Exit Function
    '        End If

    '        If receiptEmail = "" Then
    '            sendEmailtoSupplier = False
    '            errMsg = "Process Send PO [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
    '            Exit Function
    '        End If

    '        ls_Subject = "Send To Supplier PO No: " & pPONo & "-" & pSupplier
    '        ls_Body = clsNotification.GetNotification("11", "", pPONo.Trim)
    '        ls_Attachment = Trim(pPathFile) & "\" & pFileName

    '        If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
    '            sendEmailtoSupplier = False
    '            Exit Function
    '        End If

    '        sendEmailtoSupplier = True

    '    Catch ex As Exception
    '        sendEmailtoSupplier = False
    '        errMsg = "Process Send PO [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
    '        Exit Function
    '    Finally
    '        If Not dsEmail Is Nothing Then
    '            dsEmail.Dispose()
    '        End If
    '    End Try
    'End Function
End Class
