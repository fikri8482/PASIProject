Imports GlobalSetting
Imports System.Threading

Public Class frmSendMailExport

#Region "Declaration"
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal revert As Integer) As Integer
    Private Declare Function EnableMenuItem Lib "user32" (ByVal menu As Integer, ByVal ideEnableItem As Integer, ByVal enable As Integer) As Integer
    Private Const SC_CLOSE As Integer = &HF060
    Private Const MF_BYCOMMAND As Integer = &H0
    Private Const MF_GRAYED As Integer = &H1
    Private Const MF_ENABLED As Integer = &H0

    Dim intervalpro As TimeSpan
    Dim processTime As Boolean

    Dim cls As clsGlobal
    Dim Log As clsLog
    Dim cfg As New clsConfig

    Dim UserLogin = "admin"

    Dim AutoAppExport As Boolean = False ' untuk aktifkan module auto approve po export
    Dim DNExport As Boolean = False 'untuk aktifkan module DN send to supplier
    Dim ReceivingFWD As Boolean = False 'untuk aktifkan module Receiving FWD
    Dim ReceivingToSupplier As Boolean = False
    Dim TallyData As Boolean = False 'untuk aktifkan module Tally Data

    Public SubjectEmail As String = "[TRIAL] "

    Dim screenName As String = ""

#End Region

#Region "Event"
    Private Sub frmSendMail_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Label12.Text = "Version " & Me.ProductVersion

        Try
            cls = New clsGlobal(cfg.ConnectionString, UserLogin)
            Log = New clsLog(cfg.ConnectionString, UserLogin)

            rtbProcess.Text = ""
            txtMsg.Text = ""
            lblDB.Text = "SERVER: [" & Trim(cfg.Server) & "], DATABASE: [" & cfg.Database & "]"

            loadSetting()

            timerProcess.Enabled = True

            If (CDbl(txtSechedule.Text)) = "0" Then
                timerProcess.Interval = 100
            Else
                timerProcess.Interval = CDbl(txtSechedule.Text) * 1000 '1 menit
            End If

            txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            intervalpro = TimeSpan.FromSeconds(CDbl(txtSechedule.Text))
            Dim Last As Date = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")

            btnManual.Enabled = True
            txtAttachmentDOM.Enabled = False
            txtSaveAsDOM.Enabled = False
            txtSechedule.Enabled = False
            btnExit.Enabled = True

            processTime = False
        Catch ex As Exception
            cls.up_ShowMsg(ex.Message, txtMsg, GlobalSetting.clsGlobal.MsgTypeEnum.ErrorMsg)
            Log.WriteToErrorLog(Me.Tag, txtMsg.Text, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
        End Try
    End Sub

    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnManual_Click(sender As System.Object, e As System.EventArgs) Handles btnManual.Click
        Me.Cursor = Cursors.WaitCursor
        sendExcel()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub frmGetMail_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        DisableCloseButton(Me)
    End Sub

    Private Sub timerProcess_Tick(sender As Object, e As System.EventArgs) Handles timerProcess.Tick
        If Format(Now, "yyyy-MM-dd HH:mm:ss") > txtNext.Text And processTime = False Then
            Me.Cursor = Cursors.WaitCursor
            sendExcel()
            Me.Cursor = Cursors.Default
        End If
    End Sub
#End Region

#Region "Procedure"
    Public Shared Sub DisableCloseButton(ByVal form As System.Windows.Forms.Form)
        Select Case EnableMenuItem(GetSystemMenu(form.Handle.ToInt32, 0), SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED)
            Case MF_ENABLED
            Case MF_GRAYED
            Case &HFFFFFFFF
                Throw New Exception("The Close menu item does not exist.")
            Case Else
        End Select
    End Sub

    Private Sub loadSetting()
        Try
            Dim ls_SQL As String = ""

            ls_SQL = " SELECT OriginalTemplateFolder, SaveAsTemplateFolder, IntervalSendExcel FROM MS_EmailSetting_Export "
            Dim ds As New DataSet
            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                txtAttachmentDOM.Text = ds.Tables(0).Rows(0)("OriginalTemplateFolder")
                txtSaveAsDOM.Text = ds.Tables(0).Rows(0)("SaveAsTemplateFolder")
                txtSechedule.Text = ds.Tables(0).Rows(0)("IntervalSendExcel")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub sendExcel()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""

        Try
            timerProcess.Enabled = False
            processTime = True

            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Batch Process", rtbProcess)

            '14. Send PO Export Monthly
            'up_SendDataPOExportMonthly()

            '15. Send PO Export Emergency
            'up_SendDataPOExportEmergency()

            '16. Auto Approve PO Export
            'up_AutoApprovePOExport()

            '17. Send Customer Delivery Confirmation
            'up_SendDataDNExport()

            ''18. Send Receiving Forwarder
            'If ReceivingFWD = True Then
            '    up_SendDataReceivingFWD()
            'End If

            ''19. Send Receiving to Supplier
            'If ReceivingToSupplier = True Then
            '    up_SendDataReceivingSupplier()
            'End If

            '17. Send DN Replacement
            up_SendDataDNReplacementExport()

            '18. Send Receiving Forwarder OK
            'up_SendDataReceivingFWD()

            '20. Send Invoice OK 
            up_SendDataTallyData()

            '21. Send Invoice OK
            'up_SendDataInvoice()

            '22. Send Shipping Instruction OK
            'up_SendDataShippingInstruction()

            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Batch Process", rtbProcess)

        Catch ex As Exception
            cls.up_ShowMsg(ex.Message, txtMsg, GlobalSetting.clsGlobal.MsgTypeEnum.ErrorMsg)
            Log.WriteToErrorLog(Me.Tag, txtMsg.Text, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
        Finally
            txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            intervalpro = TimeSpan.FromSeconds(CDbl(txtSechedule.Text))
            Dim Last As Date = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")            
            processTime = False
            timerProcess.Enabled = True
        End Try

    End Sub

    Private Sub up_SendDataPOExportMonthly()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataPOExportMonthly"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send PO Export Monthly To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send PO Export Monthly To Supplier")

            clsPOExportMonthly.up_SendPOExportMonthly(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Export Monthly data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Export Monthly data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send PO Export Monthly To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send PO Export Monthly To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataPOExportEmergency()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataPOExportEmergency"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send PO Export Emergency To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send PO Export Emergency To Supplier")

            clsPOExportEmergency.up_SendPOExportEmergency(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Export Emergency data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Export Emergency data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send PO Export Emergency To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send PO Export Emergency To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_AutoApprovePOExport()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "AutoApprovePOExport"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Auto Approve PO Export", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Auto Approve PO Export")

            clsAutoApprovePOExport.up_AutoApprovePOExport(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Export Auto Approve data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Export Auto Approve data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Auto Approve PO Export", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Auto Approve PO Export")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataDNExport()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataDNExport"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send DN Export To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send DN Export To Supplier")

            clsDNExport.up_DNExport(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No DN Export data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No DN Export data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send DN Export To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send DN Export To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataDNRemainingExport()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataDNExport"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send DN Export To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send DN Export To Supplier")

            clsDNExport.up_DNExport(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No DN Export data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No DN Export data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send DN Export To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send DN Export To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataReceivingFWD()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataReceivingFWD"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Receiving To FWD", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Receiving To FWD")

            clsReceivingFWD.up_ReceivingFWD(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Receiving FWD data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Receiving FWD data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Receiving To FWD", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Receiving To FWD")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataReceivingSupplier()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataReceivingSupp"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Receiving To Supp", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Receiving To Supp")

            clsReceivingFWD.up_ReceivingFWD(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Receiving Supp data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Receiving Supp data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Receiving To Supp", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Receiving To Supp")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataDNReplacementExport()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataDNReplacementExport"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send DN Replacement Export To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send DN Replacement Export To Supplier")

            clsDNReplacementExport.up_SendDNReplacement(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No DN Replacement Export data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No DN Replacement Export data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send DN Replacement Export To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send DN Replacement Export To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataTallyData()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataTallyData"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Tally Data To FWD", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Tally Data To FWD")

            clsTallyData.up_SendShippingInstruction(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Tally Data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Tally Data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Tally Data To FWD", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Tally Data To FWD")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataInvoice()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataPackingList"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Packing List To FWD", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Packing List To FWD")

            clsInvoicePakcingList.up_SendInvoice(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Packing List data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Packing List data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Packing List To FWD", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Packing List To FWD")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendDataShippingInstruction()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendDataShippingInstruction"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Shipping Instruction To FWD", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Shipping Instruction To FWD")

            clsShippingInstruction.up_SendShippingInstruction(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Shipping Instruction data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Shipping Instruction data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Shipping Instruction To FWD", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Shipping Instruction To FWD")
            Thread.Sleep(500)
        End Try
    End Sub
#End Region
End Class
