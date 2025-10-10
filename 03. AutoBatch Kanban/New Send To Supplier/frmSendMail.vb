Imports GlobalSetting
Imports System.Threading

Public Class frmSendMail

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

    Dim Barcode As Boolean = True ' untuk aktifkan send label kanban barcode dalam bentuk pdf

    'Public SubjectEmail As String = "[TRIAL] "
    Public SubjectEmail As String = ""

    Dim screenName As String = ""

    Dim pIntervalNonKanban As String = "4"
    Dim pHourNonKanban As String = "04:00"
    Dim nextNonKanban As String
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
            nextNonKanban = pHourNonKanban

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
        If Format(Now, "yyyy-MM-dd HH:mm:ss") >= txtNext.Text And processTime = False Then
            Dim ls_SQL As String = ""

            'ls_SQL = "SELECT * FROM dbo.BatchProcessStatus"
            'Dim ds As New DataSet
            'ds = cls.uf_GetDataSet(ls_SQL)

            'If ds.Tables(0).Rows(0)("BatchProcessStatus") = "2" Then
            Me.Cursor = Cursors.WaitCursor

            sendExcel()

            'ls_SQL = "UPDATE dbo.BatchProcessStatus SET BatchProcessStatus = '3'"
            'cls.uf_ExecuteSql(ls_SQL)

            Me.Cursor = Cursors.Default
            'Else
            '    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Skip process to wait other process finished", rtbProcess)
            '    txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            '    intervalpro = TimeSpan.FromSeconds(CDbl(txtSechedule.Text))
            '    Dim Last As Date = FormatDateTime(txtLast.Text)
            '    txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")
            '    nextNonKanban = pHourNonKanban
            '    processTime = False
            '    timerProcess.Enabled = True
            'End If
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

            ls_SQL = " SELECT OriginalTemplateFolder, SaveAsTemplateFolder, IntervalSendExcel, IntervalNonKanbanApproval, NonKanbanApprovalHour FROM MS_EmailSetting "
            Dim ds As New DataSet
            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                txtAttachmentDOM.Text = ds.Tables(0).Rows(0)("OriginalTemplateFolder")
                txtSaveAsDOM.Text = ds.Tables(0).Rows(0)("SaveAsTemplateFolder")
                txtSechedule.Text = ds.Tables(0).Rows(0)("IntervalSendExcel")
                pIntervalNonKanban = ds.Tables(0).Rows(0)("IntervalNonKanbanApproval")
                pHourNonKanban = ds.Tables(0).Rows(0)("NonKanbanApprovalHour")
                'txtAttachmentDOM.Text = "D:\PASI EBWEB\IT TEST\01. TEMPLATE EXCEL"
                'txtSaveAsDOM.Text = "D:\PASI EBWEB\IT TEST\01. TEMPLATE EXCEL\RESULT"
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

            '''00. UPDATE DATA NORMALIZE DATA HT = SUDAH PERIKSA
            'up_UpdateReceivingData()

            '''01. Send Data PO Domestic = SUDAH PERIKSA
            'up_SendPO()

            '''02. Auto Approve PO Domestic = SUDAH PERIKSA
            'up_AutoApprovePODomestic()

            '''03. Create data for PO Non Kanban = SUDAH PERIKSA
            'uf_AutoPOFinalApprove()

            '''04. Send Remaining Delivery = SUDAH PERIKSA
            'up_SendRemainingDelivery()

            ' ''05. Send PO Kanban = SUDAH PERIKSA
            up_SendPOKanban()

            '''06. Send PO Non Kanban = SUDAH PERIKSA
            'If Format(Now, "HH:mm:ss") > nextNonKanban Then
            '    up_SendPONonKanban()
            'End If

            '''07. Auto Approve Kanban = SUDAH PERIKSA
            'up_AutoApproveKanban()

            '''''''08. Send data PO Revision Domestic
            ''''''up_SendPORev()

            '''''''09. Auto Approve PO Revision PO Domestic
            ''''''up_AutoApprovePORevDomestic()

            ''10. Send Receiving PASI = OK OK
            up_SendReceivingPASI()

            '''11. Send Receiving Affiliate = OK OK
            'up_SendReceivingAffiliate()

            '''12. Send SummaryOutstanding
            'up_SendSummaryOutstanding()

            '''13. Send SummaryForecast = SUDAH PERIKSA
            'up_SendSummaryForecast()

            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Batch Process", rtbProcess)

        Catch ex As Exception
            cls.up_ShowMsg(ex.Message, txtMsg, GlobalSetting.clsGlobal.MsgTypeEnum.ErrorMsg)
            Log.WriteToErrorLog(Me.Tag, txtMsg.Text, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
        Finally
            txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            intervalpro = TimeSpan.FromSeconds(CDbl(txtSechedule.Text))
            Dim Last As Date = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")
            nextNonKanban = pHourNonKanban
            processTime = False
            timerProcess.Enabled = True
        End Try

    End Sub

    Private Sub up_UpdateReceivingData()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SynchronizeData"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Synchronize Data", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Synchronize Data")

            clsSynchronizeData.up_SynchronizeData(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Synchronize Data to process."
            Else
                clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Success Synchronize Data", rtbProcess)
                Log.WriteToProcessLog(startTime, screenName, "Success Synchronize Data")
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Synchronize Data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Synchronize Data", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Synchronize Data")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendPO()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendPO"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send PO To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send PO To Supplier")

            clsPO.up_SendPODomestic(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send PO To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send PO To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_AutoApprovePODomestic()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "AutoApprovePODomestic"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Auto Approve PO Domestic", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Auto Approve PO Domestic")

            clsPOAutoApprove.up_AutoApprovePODomestic(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Domestic Auto Approve data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Domestic Auto Approve data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Auto Approve PO Domestic", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Auto Approve PO Domestic")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub uf_AutoPOFinalApprove()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "AutoPOFinalApprove"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Final Approve PO", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Final Approve PO")

            clsPOFinalApproval.up_FinalApprovePODomestic(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Final Approval data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Final Approval data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Final Approve PO", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Final Approve PO")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendPOKanban()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendPOKanban"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send PO Kanban To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send PO Kanban To Supplier")

            clsPOKanban.up_SendPOKanban(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, Barcode, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Kanban data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Kanban data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send PO Kanban To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send PO Kanban To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendPONonKanban()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendPONonKanban"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send PO Non Kanban To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send PO Non Kanban To Supplier")

            clsPONonKanban.up_SendPONonKanban(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, Barcode, pIntervalNonKanban, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Non Kanban data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Non Kanban data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send PO Non Kanban To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send PO Non Kanban To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_AutoApproveKanban()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "AutoApproveKanban"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Auto Approve Kanban", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Auto Approve Kanban")

            clsAutoApproveKanban.AutoApproveKanban(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Kanban Auto Approve data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Kanban Auto Approve data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Auto Approve Kanban", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Auto Approve Kanban")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendRemainingDelivery()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendRemainingDelivery"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Remaining Delivery To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Remaining Delivery To Supplier")

            clsDeliveryRemaining.up_SendRemainingDelivery(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Remaining Delivery data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Remaining Delivery data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Remaining Delivery To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Remaining Delivery To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendPORev()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendPORev"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send PO Rev To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send PO Rev To Supplier")

            clsPORev.up_SendPORevDomestic(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Rev data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Rev data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send PO Rev To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send PO Rev To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_AutoApprovePORevDomestic()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "AutoApprovePORevDomestic"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Auto Approve PO Domestic Rev", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Auto Approve PO Domestic Rev")

            clsPORevAutoApprove.up_AutoApprovePORevDomestic(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No PO Domestic Rev Auto Approve data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No PO Domestic Rev Auto Approve data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Auto Approve PO Domestic Rev", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Auto Approve PO Domestic Rev")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendReceivingPASI()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendReceivingPASI"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Receiving PASI To Supplier", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Receiving PASI To Supplier")

            clsReceivingPASI.up_SendReceivingPASI(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Receiving PASI data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Receiving PASI data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Receiving PASI To Supplier", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Receiving PASI To Supplier")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendReceivingAffiliate()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendReceivingAffiliate"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Receiving Affiliate to PASI", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Receiving Affiliate to PASI")

            clsReceivingAffiliate.up_SendReceivingAffiliate(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Receiving Affiliate data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Receiving Affiliate data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Receiving Affiliate to PASI", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Receiving Affiliate to PASI")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendSummaryOutstanding()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendSummaryOutstanding"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Summary Outstanding", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Summary Outstanding")

            clsSummaryOutstanding.up_SendSummaryOutstanding(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Summary Outstanding data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Summary Outstanding data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Summary Outstanding", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Summary Outstanding")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub up_SendSummaryForecast()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "SendSummaryForecast"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Summary Forecast", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Send Summary Forecast")

            clsSummaryForecast.up_SendSummaryForecast(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is No Summary Forecast data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is No Summary Forecast data to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Summary Forecast", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Send Summary Forecast")
            Thread.Sleep(500)
        End Try
    End Sub

#End Region

End Class
