Imports GlobalSetting
Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmSendMail

#Region "Declaration"
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal revert As Integer) As Integer
    Private Declare Function EnableMenuItem Lib "user32" (ByVal menu As Integer, ByVal ideEnableItem As Integer, ByVal enable As Integer) As Integer
    Private Const SC_CLOSE As Integer = &HF060
    Private Const MF_BYCOMMAND As Integer = &H0
    Private Const MF_GRAYED As Integer = &H1
    Private Const MF_ENABLED As Integer = &H0

    '------EXPORT--------
    Dim POExport As Boolean = True
    Dim AutoExport As Boolean = False
    Dim DNExport As Boolean = True
    Dim SendTally As Boolean = True

    Dim ls_newLabelEx As Boolean = True 'untuk format label baru export
    Dim newDN As Boolean = True
    Dim statusPartial = True 'Untuk Delivery Partial Export
    Dim RecExNew As Boolean = True 'untuk receiving baru
    Dim sendInvoiceEx As Boolean = True

    Dim pDate As Date
    Dim pForecastPeriod As Date
    Dim pAffCode As String
    Dim pPONo As String
    Dim pPORevNo As String
    Dim pSupplier As String
    Dim pDel As String
    Dim pDelivBy As String

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String
    Dim DefaultCredentials As Boolean
    Dim SSL As Boolean

    Dim ls_Body As String
    Dim UserName As String
    '------EXPORT--------

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

            ls_SQL = " SELECT a.OriginalTemplateFolder, a.SaveAsTemplateFolder, a.IntervalSendExcel, b.IntervalNonKanbanApproval, b.NonKanbanApprovalHour FROM MS_EmailSetting_Export a, MS_EmailSetting b "
            Dim ds As New DataSet
            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                txtAttachmentDOM.Text = ds.Tables(0).Rows(0)("OriginalTemplateFolder")
                txtSaveAsDOM.Text = ds.Tables(0).Rows(0)("SaveAsTemplateFolder")
                txtSechedule.Text = ds.Tables(0).Rows(0)("IntervalSendExcel")
                pIntervalNonKanban = ds.Tables(0).Rows(0)("IntervalNonKanbanApproval")
                pHourNonKanban = ds.Tables(0).Rows(0)("NonKanbanApprovalHour")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub sendExcel()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        txtMsg.Text = ""
        Try
            timerProcess.Enabled = False
            processTime = True

            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Batch Process", rtbProcess)

            ''00. UPDATE DATA NORMALIZE DATA HT = SUDAH PERIKSA
            '''20200624 Shut Down'''
            'up_UpdateReceivingData()

            ''01. Send Data PO Domestic = SUDAH PERIKSA
            up_SendPO()

            ''02. Auto Approve PO Domestic = SUDAH PERIKSA
            up_AutoApprovePODomestic()

            ''03. Create data for PO Non Kanban = SUDAH PERIKSA
            uf_AutoPOFinalApprove()

            ''04. Send Remaining Delivery = SUDAH PERIKSA
            up_SendRemainingDelivery()

            '' ''05. Send PO Kanban = SUDAH PERIKSA
            'up_SendPOKanban()

            ''06. Send PO Non Kanban = SUDAH PERIKSA
            If Format(Now, "HH:mm:ss") > nextNonKanban Then
                up_SendPONonKanban()
            End If

            ''07. Auto Approve Kanban = SUDAH PERIKSA
            up_AutoApproveKanban()

            ''''''08. Send data PO Revision Domestic
            '''''up_SendPORev()

            ''''''09. Auto Approve PO Revision PO Domestic
            '''''up_AutoApprovePORevDomestic()

            '''10. Send Receiving PASI = OK OK
            'up_SendReceivingPASI()

            ''11. Send Receiving Affiliate = OK OK
            up_SendReceivingAffiliate()

            ''12. Send SummaryOutstanding
            up_SendSummaryOutstanding()

            ''13. Send SummaryForecast = SUDAH PERIKSA
            up_SendSummaryForecast()

            ''14. Send Supplier Export
            up_SendExport()

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

    Private Function uf_PriceCls(ByVal pDesc As String)
        Dim ls_SQL As String = ""
        Dim ls_Cls As Integer = 0
        Dim ds As New DataSet
        'Using sqlConn As New SqlConnection(ConStr)
        '    sqlConn.Open()

        ls_SQL = " Select PriceCls From MS_PriceCls Where Description = '" & pDesc & "' "

        ds = cls.uf_GetDataSet(ls_SQL)

        'Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
        'Dim ds As New DataSet
        'sqlDA.SelectCommand.CommandTimeout = 200
        'sqlDA.Fill(ds)

        If ds.Tables(0).Rows.Count > 0 Then
            ls_Cls = ds.Tables(0).Rows(0).Item("PriceCls")
        End If
        'End Using

        Return ls_Cls
    End Function

    Private Function uf_DesPrice(ByVal pCls As String)
        Dim ls_SQL As String = ""
        Dim ls_Cls As String = ""
        Dim ds As New DataSet

        ls_SQL = " Select Description=RTRIM(Description) From MS_PriceCls Where PriceCls = '" & pCls & "' "

        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            ls_Cls = ds.Tables(0).Rows(0).Item("Description")
        End If

        Return ls_Cls
    End Function

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

#Region "EXPORT"

    Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
        With Rg
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
    End Sub

    Private Sub up_SendExport()
        If POExport = True Then
            '====================PO (EXPORT) MONTHLY================
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send PO Export To Supplier" & vbCrLf & _
                          " " & vbCrLf & _
                          rtbProcess.Text
            pGetExcelPOMonthly()
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send PO Export To Supplier" & vbCrLf & _
                         rtbProcess.Text

            '====================PO (EXPORT) MONTHLY================
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send PO Export To Supplier" & vbCrLf & _
                          " " & vbCrLf & _
                          rtbProcess.Text
            pGetExcelPOEmergency()
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send PO Export To Supplier" & vbCrLf & _
                         rtbProcess.Text

            If DNExport = True Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Customer Delivery Confirmation To Supplier" & vbCrLf & _
                          " " & vbCrLf & _
                          rtbProcess.Text
                Excel_CustomerDeliveryConfirmation()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Customer Delivery Confirmation To Supplier" & vbCrLf & _
                             rtbProcess.Text

                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Customer Delivery Confirmation Replace To Supplier" & vbCrLf & _
                          " " & vbCrLf & _
                          rtbProcess.Text
                Excel_DeliveryConfirmationReplacement()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Customer Delivery Confirmation Replace To Supplier" & vbCrLf & _
                             rtbProcess.Text

                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Delivery to Forwarder" & vbCrLf & _
                            " " & vbCrLf & _
                            rtbProcess.Text
                Excel_DeliveryToForwarder()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Delivery to Forwarder " & vbCrLf & _
                             rtbProcess.Text

                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Remaining Delivery to Forwarder" & vbCrLf & _
                            " " & vbCrLf & _
                            rtbProcess.Text
                'Excel_DeliveryConfirmationRemaining()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Remaining Delivery to Forwarder " & vbCrLf & _
                             rtbProcess.Text

                'Fungsi Baru Fikri => 2022-10-10 untuk kirim Invoice Export Excel ke Supplier
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Invoice Export and Good Receiving to Supplier " & vbCrLf &
                                " " & vbCrLf &
                                rtbProcess.Text
                up_InvoiceSupplierExport()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Invoice Export and Good Receiving to Supplier " & vbCrLf &
                             rtbProcess.Text
                'Fungsi Baru Fikri => 2022-10-10 untuk kirim Invoice Excel ke Supplier

                'Fungsi Baru Fikri => 2023-03-01 untuk kirim Moving Good ke Forwarder
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Moving Good to Forwarder " & vbCrLf & _
                                " " & vbCrLf & _
                                rtbProcess.Text
                up_MovingGoodExport()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Moving Good to Forwarder " & vbCrLf & _
                             rtbProcess.Text
                'Fungsi Baru Fikri => 2023-03-01 untuk kirim Moving Good ke Forwarder

                'rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Receiving To Supplier" & vbCrLf & _
                '            " " & vbCrLf & _
                '            rtbProcess.Text
                'Excel_GoodReceivingANDInvoice()
                'rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Receiving To Supplier" & vbCrLf & _
                '             rtbProcess.Text
            End If

            ''auto app
            'If AutoExport = True Then
            '    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Auto Approved PO Supplier" & vbCrLf & _
            '              " " & vbCrLf & _
            '              rtbProcess.Text
            '    AutoApprovePOEXPORT()
            '    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Auto Approved PO  Supplier" & vbCrLf & _
            '                 rtbProcess.Text
            '    'auto app

            '    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Customer Delivery Confirmation To Supplier" & vbCrLf & _
            '                         " " & vbCrLf & _
            '                         rtbProcess.Text
            'End If

            If SendTally = True Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Tally to Forwarder " & vbCrLf & _
                            " " & vbCrLf & _
                            rtbProcess.Text
                pGetExcelTally()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Tally to Forwarder " & vbCrLf & _
                             rtbProcess.Text

                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Tally PDF to Forwarder " & vbCrLf & _
                            " " & vbCrLf & _
                            rtbProcess.Text
                pGetPDFTally()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Tally PDF to Forwarder " & vbCrLf & _
                             rtbProcess.Text

                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Shipping Information to Forwarder " & vbCrLf & _
                            " " & vbCrLf & _
                            rtbProcess.Text
                pGetPDFShipping()
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Shipping Information to Forwarder " & vbCrLf & _
                             rtbProcess.Text

                If sendInvoiceEx = True Then
                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Invoice to Forwarder " & vbCrLf & _
                                " " & vbCrLf & _
                                rtbProcess.Text
                    pGetPDFINVEX()
                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Invoice to Forwarder " & vbCrLf & _
                                 rtbProcess.Text
                End If
            End If


            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Moving List to Forwarder " & vbCrLf & _
                            " " & vbCrLf & _
                            rtbProcess.Text
            Excel_MovingList()
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Moving List to Forwarder " & vbCrLf & _
                         rtbProcess.Text

            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send Stock Opname to Forwarder " & vbCrLf & _
                           " " & vbCrLf & _
                           rtbProcess.Text
            pGetExcelSTOCKOPNAME()
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send Stock Opname to Forwarder " & vbCrLf & _
                         rtbProcess.Text

            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Start Send PO Cancelation List " & vbCrLf & _
                            " " & vbCrLf & _
                            rtbProcess.Text
            Excel_CancelationList()
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " End Send PO Cancelation List " & vbCrLf & _
                         rtbProcess.Text
        End If
    End Sub

    Private Sub pGetExcelPOMonthly()
        Const ColorYellow As Single = 65535

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        'copy file from server to local
        Dim NewFileCopy As String

        'MdlConn.ReadConnection()
        Dim ls_SQL As String = ""
        Dim ls_Consignee As String = ""
        Dim ls_ExcelCls As String = ""

        Dim ds As New DataSet
        Dim dsSplit As New DataSet
        Dim dsDetail As New DataSet

        Dim pPeriod As Date
        Dim pOrderNo1 As String
        Dim pCommercialCls As String
        Dim pETDWeek1
        Dim pETDWeek2
        Dim pETDWeek3
        Dim pETDWeek4
        Dim pETDWeek5
        Dim pETDVendor

        Dim xlApp = New Excel.Application
        Dim booSplit As Boolean = False
        Dim temp_Filename1 As String = ""
        Dim temp_Filename2 As String = ""

        Try
            ls_SQL = " SELECT Consignee = isnull(MA.ConsigneeCode,''), ME.*, ETDWEEK1 = MEE1.ETAForwarder, ETDWEEK2 = MEE2.ETAForwarder, ETDWEEK3 = MEE3.ETAForwarder, " & vbCrLf & _
                  " ETDWEEK4 = MEE4.ETAForwarder, ETDWEEK5 = MEE5.ETAForwarder, ISNULL(ExcelCls, '0') ExcelCls, ISNULL(ME.SplitReffPONo, '') SplitReffPONo " & vbCrLf & _
                  " FROM dbo.PO_Master_Export ME " & vbCrLf & _
                  " LEFT JOIN MS_AFFILIATE MA ON MA.AffiliateID = ME.AffiliateID " & vbCrLf & _
                  " LEFT JOIN MS_ETD_Export MEE1 ON MEE1.AffiliateID = ME.AffiliateID " & vbCrLf & _
                  " 	AND MEE1.SupplierID = ME.SupplierID " & vbCrLf & _
                  " 	AND MEE1.Period = ME.Period " & vbCrLf & _
                  " 	AND MEE1.Week = '1' " & vbCrLf & _
                  " LEFT JOIN MS_ETD_Export MEE2 ON MEE2.AffiliateID = ME.AffiliateID " & vbCrLf & _
                  " 	AND MEE2.SupplierID = ME.SupplierID " & vbCrLf & _
                  " 	AND MEE2.Period = ME.Period " & vbCrLf & _
                  " 	AND MEE2.Week = '2' " & vbCrLf
            ls_SQL = ls_SQL + " LEFT JOIN MS_ETD_Export MEE3 ON MEE3.AffiliateID = ME.AffiliateID " & vbCrLf & _
                              " 	AND MEE3.SupplierID = ME.SupplierID " & vbCrLf & _
                              " 	AND MEE3.Period = ME.Period " & vbCrLf & _
                              " 	AND MEE3.Week = '3' " & vbCrLf & _
                              " LEFT JOIN MS_ETD_Export MEE4 ON MEE4.AffiliateID = ME.AffiliateID " & vbCrLf & _
                              " 	AND MEE4.SupplierID = ME.SupplierID " & vbCrLf & _
                              " 	AND MEE4.Period = ME.Period " & vbCrLf & _
                              " 	AND MEE4.Week = '4' " & vbCrLf & _
                              " LEFT JOIN MS_ETD_Export MEE5 ON MEE5.AffiliateID = ME.AffiliateID " & vbCrLf & _
                              " 	AND MEE5.SupplierID = ME.SupplierID " & vbCrLf & _
                              " 	AND MEE5.Period = ME.Period " & vbCrLf
            ls_SQL = ls_SQL + " 	AND MEE5.Week = '5' " & vbCrLf &
                              " WHERE ExcelCls IN ('1', '3') and  " & vbCrLf &
                              " EmergencyCls = 'M' ORDER BY OrderNo1"

            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1

                    temp_Filename1 = ""
                    temp_Filename2 = ""

                    ls_ExcelCls = ds.Tables(0).Rows(xi)("ExcelCls")
                    If ls_ExcelCls = "3" Then
                        ls_SQL = " SELECT Consignee = isnull(MA.ConsigneeCode,''), ME.*, ETDWEEK1 = MEE1.ETAForwarder, ETDWEEK2 = MEE2.ETAForwarder, ETDWEEK3 = MEE3.ETAForwarder, " & vbCrLf & _
                              " ETDWEEK4 = MEE4.ETAForwarder, ETDWEEK5 = MEE5.ETAForwarder, ISNULL(ExcelCls, '0') ExcelCls, ISNULL(ME.SplitReffPONo, '') SplitReffPONo " & vbCrLf & _
                              " FROM dbo.PO_Master_Export ME " & vbCrLf & _
                              " LEFT JOIN MS_AFFILIATE MA ON MA.AffiliateID = ME.AffiliateID " & vbCrLf & _
                              " LEFT JOIN MS_ETD_Export MEE1 ON MEE1.AffiliateID = ME.AffiliateID " & vbCrLf & _
                              " 	AND MEE1.SupplierID = ME.SupplierID " & vbCrLf & _
                              " 	AND MEE1.Period = ME.Period " & vbCrLf & _
                              " 	AND MEE1.Week = '1' " & vbCrLf & _
                              " LEFT JOIN MS_ETD_Export MEE2 ON MEE2.AffiliateID = ME.AffiliateID " & vbCrLf & _
                              " 	AND MEE2.SupplierID = ME.SupplierID " & vbCrLf & _
                              " 	AND MEE2.Period = ME.Period " & vbCrLf & _
                              " 	AND MEE2.Week = '2' " & vbCrLf
                        ls_SQL = ls_SQL + " LEFT JOIN MS_ETD_Export MEE3 ON MEE3.AffiliateID = ME.AffiliateID " & vbCrLf & _
                                          " 	AND MEE3.SupplierID = ME.SupplierID " & vbCrLf & _
                                          " 	AND MEE3.Period = ME.Period " & vbCrLf & _
                                          " 	AND MEE3.Week = '3' " & vbCrLf & _
                                          " LEFT JOIN MS_ETD_Export MEE4 ON MEE4.AffiliateID = ME.AffiliateID " & vbCrLf & _
                                          " 	AND MEE4.SupplierID = ME.SupplierID " & vbCrLf & _
                                          " 	AND MEE4.Period = ME.Period " & vbCrLf & _
                                          " 	AND MEE4.Week = '4' " & vbCrLf & _
                                          " LEFT JOIN MS_ETD_Export MEE5 ON MEE5.AffiliateID = ME.AffiliateID " & vbCrLf & _
                                          " 	AND MEE5.SupplierID = ME.SupplierID " & vbCrLf & _
                                          " 	AND MEE5.Period = ME.Period " & vbCrLf
                        ls_SQL = ls_SQL + " 	AND MEE5.Week = '5' " & vbCrLf & _
                                          " WHERE ME.PONo = '" & Trim(ds.Tables(0).Rows(xi)("PONo")) & "' " & vbCrLf & _
                                          " AND ME.AffiliateID = '" & Trim(ds.Tables(0).Rows(xi)("AffiliateID")) & "' " & vbCrLf & _
                                          " AND ME.SupplierID = '" & Trim(ds.Tables(0).Rows(xi)("SupplierID")) & "' " & vbCrLf & _
                                          " AND ME.SplitReffPONo = '" & Trim(ds.Tables(0).Rows(xi)("SplitReffPONo")) & "' " & vbCrLf & _
                                          " AND EmergencyCls = 'M' ORDER BY OrderNo1"

                        dsSplit = cls.uf_GetDataSet(ls_SQL)

                        booSplit = True
                        pDate = Now
                        pAffCode = dsSplit.Tables(0).Rows(0)("AffiliateID")
                        ls_Consignee = dsSplit.Tables(0).Rows(0)("Consignee")
                        pPONo = dsSplit.Tables(0).Rows(0)("PONo")
                        pSupplier = dsSplit.Tables(0).Rows(0)("SupplierID")
                        pPeriod = dsSplit.Tables(0).Rows(0)("Period")
                        pDel = dsSplit.Tables(0).Rows(0)("ForwarderID")
                        pOrderNo1 = dsSplit.Tables(0).Rows(0)("OrderNo1")
                        pCommercialCls = dsSplit.Tables(0).Rows(0)("CommercialCls")
                        pETDWeek1 = dsSplit.Tables(0).Rows(0)("ETDWEEK1")
                        pETDWeek2 = dsSplit.Tables(0).Rows(0)("ETDWEEK2")
                        pETDWeek3 = dsSplit.Tables(0).Rows(0)("ETDWEEK3")
                        pETDWeek4 = dsSplit.Tables(0).Rows(0)("ETDWEEK4")
                        pETDWeek5 = dsSplit.Tables(0).Rows(0)("ETDWEEK5")
                        pETDVendor = dsSplit.Tables(0).Rows(0)("ETDVendor1")
                    Else
Split:
                        booSplit = False
                        pDate = Now
                        pAffCode = ds.Tables(0).Rows(xi)("AffiliateID")
                        ls_Consignee = ds.Tables(0).Rows(xi)("Consignee")
                        pPONo = ds.Tables(0).Rows(xi)("PONo")
                        pSupplier = ds.Tables(0).Rows(xi)("SupplierID")
                        pPeriod = ds.Tables(0).Rows(xi)("Period")
                        pDel = ds.Tables(0).Rows(xi)("ForwarderID")
                        pOrderNo1 = ds.Tables(0).Rows(xi)("OrderNo1")
                        pCommercialCls = ds.Tables(0).Rows(xi)("CommercialCls")
                        pETDWeek1 = ds.Tables(0).Rows(xi)("ETDWEEK1")
                        pETDWeek2 = ds.Tables(0).Rows(xi)("ETDWEEK2")
                        pETDWeek3 = ds.Tables(0).Rows(xi)("ETDWEEK3")
                        pETDWeek4 = ds.Tables(0).Rows(xi)("ETDWEEK4")
                        pETDWeek5 = ds.Tables(0).Rows(xi)("ETDWEEK5")
                        pETDVendor = ds.Tables(0).Rows(xi)("ETDVendor1")

                        If ls_ExcelCls = "3" Then
                            pOrderNo1 = ds.Tables(0).Rows(xi)("SplitReffPONo")
                        End If
                    End If

                    Dim dsEmail As New DataSet
                    dsEmail = EmailToEmailCCPOMonthly(pAffCode, "PASI", pSupplier)
                    '1 CC Affiliate'2 CC PASI'3 CC & TO Supplier
                    For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(i)("affiliatepocc")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("affiliatepocc")
                        End If
                        If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                            fromEmail = dsEmail.Tables(0).Rows(i)("toEmail")
                        End If
                        If receiptEmail = "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(i)("affiliatepoto")
                        Else
                            receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(i)("affiliatepoto")
                        End If
                    Next

                    receiptCCEmail = Replace(receiptCCEmail, ",", ";")
                    receiptEmail = Replace(receiptEmail, ",", ";")

                    'Create Excel File
                    Dim fi As New FileInfo(Trim(txtAttachmentDOM.Text) & "\Template PO Export (Monthly).xlsm") 'File dari Local
                    If Not fi.Exists Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because file Excel isn't Found" & vbCrLf & _
                                        rtbProcess.Text
                        Exit Sub
                    End If

                    NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Template PO Export (Monthly).xlsm"
                    Dim ls_file As String = NewFileCopy
                    ExcelBook = xlApp.Workbooks.Open(ls_file)
                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                    ExcelSheet.Range("H1").Value = "POEM"
                    ExcelSheet.Range("H2").Value = fromEmail
                    ExcelSheet.Range("H3").Value = pAffCode
                    ExcelSheet.Range("H4").Value = pDel
                    ExcelSheet.Range("H5").Value = pSupplier
                    ExcelSheet.Range("S1").Value = pPONo
                    ExcelSheet.Range("S1").Font.Color = Color.White
                    ExcelSheet.Range("Y2").Value = ""

                    'Order No
                    If Trim(pPONo) = Trim(pOrderNo1) Then
                        ExcelSheet.Range("I9").Value = pOrderNo1
                    Else
                        ExcelSheet.Range("I9").Value = pOrderNo1
                        ExcelSheet.Range("I11").Value = pPONo
                    End If

                    'PO Date
                    ExcelSheet.Range("AE9").Value = Format(pPeriod, "yyyy-MM-dd")

                    'Commercial Cls
                    ExcelSheet.Range("AE11").Value = IIf(pCommercialCls = "1", "YES", "NO")

                    'To
                    ExcelSheet.Range("I13").Value = pSupplier
                    Dim dsSupp As New DataSet
                    dsSupp = Supplier(Trim(pSupplier))
                    ExcelSheet.Range("I14").Value = dsSupp.Tables(0).Rows(0)("Address")
                    ExcelSheet.Range("I14:X16").WrapText = True

                    'Buyer
                    ExcelSheet.Range("I18").Value = pAffCode
                    Dim dsAffp As New DataSet
                    dsAffp = Affiliate(Trim(pAffCode))
                    ExcelSheet.Range("I18").Value = dsAffp.Tables(0).Rows(0)("BuyerName")
                    ExcelSheet.Range("I19").Value = dsAffp.Tables(0).Rows(0)("BuyerAddress")
                    ExcelSheet.Range("I19:X21").WrapText = True

                    'Delivery To
                    ExcelSheet.Range("AE13").Value = pDel
                    Dim dsDelivery As New DataSet
                    dsDelivery = Forwarder(Trim(pDel))
                    ExcelSheet.Range("AE14").Value = dsDelivery.Tables(0).Rows(0)("Address")
                    ExcelSheet.Range("AE14:AT16").WrapText = True

                    ExcelSheet.Range("B38").Interior.Color = Color.White
                    ExcelSheet.Range("B38").Font.Color = Color.Black

                    dsDetail = bindDataDetailMonthly(pPeriod, pAffCode, pPONo, pSupplier, pOrderNo1)

                    If dsDetail.Tables(0).Rows.Count > 0 Then
                        For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                            'Header
                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).Merge() 'No
                            ExcelSheet.Range("D" & i + 37 & ": H" & i + 37).Merge() 'Part No
                            ExcelSheet.Range("I" & i + 37 & ": P" & i + 37).Merge() 'Part Name
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).Merge() 'UOM
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).Merge() 'MOQ
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).Merge() ' Total Order
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Merge() 'ETD Supplier
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Merge() 'Total Firm Edit Supp
                            ExcelSheet.Range("AJ" & i + 37 & ": AN" & i + 37).Merge() 'ETD Supplier Edit Supp
                            ExcelSheet.Range("AO" & i + 37 & ": AS" & i + 37).Merge()
                            ExcelSheet.Range("AT" & i + 37 & ": AX" & i + 37).Merge() 'ETD Supplier Edit Supp
                            ExcelSheet.Range("AY" & i + 37 & ": BC" & i + 37).Merge()
                            ExcelSheet.Range("BD" & i + 37 & ": BH" & i + 37).Merge()
                            ExcelSheet.Range("BI" & i + 37 & ": BM" & i + 37).Merge() 'ETD Supplier Edit Supp
                            ExcelSheet.Range("BN" & i + 37 & ": BR" & i + 37).Merge() 'ETD Supplier Edit Supp
                            ExcelSheet.Range("BS" & i + 37 & ": BW" & i + 37).Merge()
                            ExcelSheet.Range("BX" & i + 37 & ": CA" & i + 37).Merge() 'Forecast 1
                            ExcelSheet.Range("CB" & i + 37 & ": CE" & i + 37).Merge() 'Forecast 2
                            ExcelSheet.Range("CF" & i + 37 & ": CI" & i + 37).Merge() 'Forecast 3

                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("NoUrut"))
                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("D" & i + 37 & ": H" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                            ExcelSheet.Range("I" & i + 37 & ": P" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName"))
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("UOM"))
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("MOQ"))
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).NumberFormat = "#,##0"

                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).Interior.Color = ColorYellow

                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Value = Trim(IIf(IsDBNull(pETDWeek1), "", pETDWeek1)) & ""
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).NumberFormat = "yyy-MM-dd"
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Interior.Color = RGB(217, 217, 217)

                            If Not IsDBNull(pETDWeek1) Then
                                If Trim(dsDetail.Tables(0).Rows(i)("SplitReffPONo")) <> "" Then
                                    ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Value = dsDetail.Tables(0).Rows(i)("ETDVendor1")
                                    ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                                Else
                                    If pETDWeek1 = dsDetail.Tables(0).Rows(i)("ETDVendor1") Then
                                        ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                                    Else
                                        ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Value = ""
                                    End If

                                End If
                            Else
                                If Trim(dsDetail.Tables(0).Rows(i)("SplitReffPONo")) <> "" Then
                                    ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Value = dsDetail.Tables(0).Rows(i)("ETDVendor1")
                                    ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                                Else
                                    ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Value = ""
                                End If
                            End If
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Interior.Color = ColorYellow

                            ExcelSheet.Range("AJ" & i + 37 & ": AN" & i + 37).Value = Trim(IIf(IsDBNull(pETDWeek2), "", pETDWeek2)) & ""
                            ExcelSheet.Range("AJ" & i + 37 & ": AN" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("AJ" & i + 37 & ": AN" & i + 37).NumberFormat = "yyy-MM-dd"
                            ExcelSheet.Range("AJ" & i + 37 & ": AN" & i + 37).Interior.Color = RGB(217, 217, 217)

                            If Not IsDBNull(pETDWeek2) Then
                                If Trim(dsDetail.Tables(0).Rows(i)("SplitReffPONo")) <> "" Then
                                    ExcelSheet.Range("AJ" & i + 37 & ": AN" & i + 37).Value = ""
                                    ExcelSheet.Range("AO" & i + 37 & ": AS" & i + 37).Value = ""
                                Else
                                    If pETDWeek2 = dsDetail.Tables(0).Rows(i)("ETDVendor1") Then
                                        ExcelSheet.Range("AO" & i + 37 & ": AS" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                                    Else
                                        ExcelSheet.Range("AO" & i + 37 & ": AS" & i + 37).Value = ""
                                    End If
                                End If
                            Else
                                ExcelSheet.Range("AO" & i + 37 & ": AS" & i + 37).Value = ""
                            End If

                            ExcelSheet.Range("AO" & i + 37 & ": AS" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AO" & i + 37 & ": AS" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("AO" & i + 37 & ": AS" & i + 37).Interior.Color = ColorYellow

                            ExcelSheet.Range("AT" & i + 37 & ": AX" & i + 37).Value = Trim(IIf(IsDBNull(pETDWeek3), "", pETDWeek3)) & ""
                            ExcelSheet.Range("AT" & i + 37 & ": AX" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("AT" & i + 37 & ": AX" & i + 37).NumberFormat = "yyy-MM-dd"
                            ExcelSheet.Range("AT" & i + 37 & ": AX" & i + 37).Interior.Color = RGB(217, 217, 217)

                            If Not IsDBNull(pETDWeek3) Then
                                If Trim(dsDetail.Tables(0).Rows(i)("SplitReffPONo")) <> "" Then
                                    ExcelSheet.Range("AT" & i + 37 & ": AX" & i + 37).Value = ""
                                    ExcelSheet.Range("AY" & i + 37 & ": BC" & i + 37).Value = ""
                                Else
                                    If pETDWeek3 = dsDetail.Tables(0).Rows(i)("ETDVendor1") Then
                                        ExcelSheet.Range("AY" & i + 37 & ": BC" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                                    Else
                                        ExcelSheet.Range("AY" & i + 37 & ": BC" & i + 37).Value = ""
                                    End If
                                End If
                            Else
                                ExcelSheet.Range("AY" & i + 37 & ": BC" & i + 37).Value = ""
                            End If

                            ExcelSheet.Range("AY" & i + 37 & ": BC" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AY" & i + 37 & ": BC" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("AY" & i + 37 & ": BC" & i + 37).Interior.Color = ColorYellow

                            ExcelSheet.Range("BD" & i + 37 & ": BH" & i + 37).Value = Trim(IIf(IsDBNull(pETDWeek4), "", pETDWeek4)) & ""
                            ExcelSheet.Range("BD" & i + 37 & ": BH" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("BD" & i + 37 & ": BH" & i + 37).NumberFormat = "yyy-MM-dd"
                            ExcelSheet.Range("BD" & i + 37 & ": BH" & i + 37).Interior.Color = RGB(217, 217, 217)

                            If Not IsDBNull(pETDWeek4) Then
                                If Trim(dsDetail.Tables(0).Rows(i)("SplitReffPONo")) <> "" Then
                                    ExcelSheet.Range("BD" & i + 37 & ": BH" & i + 37).Value = ""
                                    ExcelSheet.Range("BI" & i + 37 & ": BM" & i + 37).Value = ""
                                Else
                                    If pETDWeek4 = dsDetail.Tables(0).Rows(i)("ETDVendor1") Then
                                        ExcelSheet.Range("BI" & i + 37 & ": BM" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                                    Else
                                        ExcelSheet.Range("BI" & i + 37 & ": BM" & i + 37).Value = ""
                                    End If
                                End If
                            Else
                                ExcelSheet.Range("BI" & i + 37 & ": BM" & i + 37).Value = ""
                            End If

                            ExcelSheet.Range("BI" & i + 37 & ": BM" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("BI" & i + 37 & ": BM" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("BI" & i + 37 & ": BM" & i + 37).Interior.Color = ColorYellow

                            ExcelSheet.Range("BN" & i + 37 & ": BR" & i + 37).Value = Trim(IIf(IsDBNull(pETDWeek5), "", pETDWeek5)) & ""
                            ExcelSheet.Range("BN" & i + 37 & ": BR" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("BN" & i + 37 & ": BR" & i + 37).NumberFormat = "yyy-MM-dd"
                            ExcelSheet.Range("BN" & i + 37 & ": BR" & i + 37).Interior.Color = RGB(217, 217, 217)

                            If Not IsDBNull(pETDWeek5) Then
                                If Trim(dsDetail.Tables(0).Rows(i)("SplitReffPONo")) <> "" Then
                                    ExcelSheet.Range("BN" & i + 37 & ": BR" & i + 37).Value = ""
                                    ExcelSheet.Range("BS" & i + 37 & ": BW" & i + 37).Value = ""
                                Else
                                    If pETDWeek5 = dsDetail.Tables(0).Rows(i)("ETDVendor1") Then
                                        ExcelSheet.Range("BS" & i + 37 & ": BW" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                                    Else
                                        ExcelSheet.Range("BS" & i + 37 & ": BW" & i + 37).Value = ""
                                    End If
                                End If
                            Else
                                ExcelSheet.Range("BS" & i + 37 & ": BW" & i + 37).Value = ""
                            End If

                            ExcelSheet.Range("BS" & i + 37 & ": BW" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("BS" & i + 37 & ": BW" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("BS" & i + 37 & ": BW" & i + 37).Interior.Color = ColorYellow

                            ExcelSheet.Range("BX" & i + 37 & ": CA" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Forecast1"))
                            ExcelSheet.Range("BX" & i + 37 & ": CA" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("BX" & i + 37 & ": CA" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("BX" & i + 37 & ": CA" & i + 37).Interior.Color = RGB(217, 217, 217)

                            ExcelSheet.Range("CB" & i + 37 & ": CE" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Forecast2"))
                            ExcelSheet.Range("CB" & i + 37 & ": CE" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("CB" & i + 37 & ": CE" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("CB" & i + 37 & ": CE" & i + 37).Interior.Color = RGB(217, 217, 217)

                            ExcelSheet.Range("CF" & i + 37 & ": CI" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Forecast3"))
                            ExcelSheet.Range("CF" & i + 37 & ": CI" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("CF" & i + 37 & ": CI" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("CF" & i + 37 & ": CI" & i + 37).Interior.Color = RGB(217, 217, 217)

                            DrawAllBorders(ExcelSheet.Range("B" & i + 37 & ": CI" & i + 37))
                            ExcelSheet.Range("B" & i + 37 & ": AD" & i + 37).Interior.Color = RGB(217, 217, 217)
                        Next
                    End If

                    ExcelSheet.Range("B" & i + 37).Value = "E"
                    ExcelSheet.Range("B" & i + 37).Interior.Color = Color.Black
                    ExcelSheet.Range("B" & i + 37).Font.Color = Color.White

                    xlApp.DisplayAlerts = False
                    If temp_Filename1 = "" Then
                        If Trim(pPONo) = Trim(pOrderNo1) Then
                            temp_Filename1 = "PO Monthly " & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"
                        Else
                            temp_Filename1 = "PO Monthly " & Trim(pPONo) & "-Split (" & Trim(pOrderNo1) & ")-" & Trim(pSupplier) & ".xlsm"
                        End If

                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\" & temp_Filename1)
                    Else
                        If Trim(pPONo) = Trim(pOrderNo1) Then
                            temp_Filename2 = "PO Monthly " & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"
                        Else
                            temp_Filename2 = "PO Monthly " & Trim(pPONo) & "-Split (" & Trim(pOrderNo1) & ")-" & Trim(pSupplier) & ".xlsm"
                        End If

                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\" & temp_Filename2)
                    End If

                    If dsDetail.Tables(0).Rows.Count = 0 Then
                        temp_Filename2 = ""
                    End If

                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & temp_Filename1 & temp_Filename2 & " OK. " & vbCrLf & _
                    rtbProcess.Text

                    If booSplit Then GoTo Split

                    If ls_ExcelCls = "3" Then
                        pOrderNo1 = ds.Tables(0).Rows(xi)("OrderNo1")
                    End If

                    If sendEmailPOtoSupllierMonthly(temp_Filename1, temp_Filename2, pPONo, pOrderNo1) = False Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & temp_Filename1 & temp_Filename2 & " NG. " & vbCrLf & _
                        rtbProcess.Text
                        GoTo keluar
                    Else
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & temp_Filename1 & temp_Filename2 & " OK. " & vbCrLf & _
                        rtbProcess.Text
                    End If

                    Call UpdateExcelPOMonthly(True, pAffCode, pPONo, pSupplier, pOrderNo1)

                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    Thread.Sleep(500)
keluar:
                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                Next
            Else
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because there is nothing PO to send " & vbCrLf & _
                    rtbProcess.Text
            End If
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Sub pGetExcelPOEmergency()
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        Dim booSplit As Boolean = False
        Dim ls_ExcelCls As String = ""
        Dim temp_Filename1 As String = ""
        Dim temp_Filename2 As String = ""

        'copy file from server to local
        Dim NewFileCopy As String

        'MdlConn.ReadConnection()
        Dim ls_SQL As String = ""
        Dim ds As New DataSet
        Dim dsHeader As New DataSet
        Dim dsDetail As New DataSet
        Dim dsEta As New DataSet
        Dim pPeriod As Date
        Dim pOrderNo1 As String
        Dim xlApp = New Excel.Application

        Try
            ls_SQL = "SELECT * FROM dbo.PO_Master_Export" & vbCrLf & _
                     "WHERE EmergencyCls = 'E' and ExcelCls In ('1','3')"
            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then


                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pDate = Now
                    pAffCode = ds.Tables(0).Rows(xi)("AffiliateID")
                    pPONo = ds.Tables(0).Rows(xi)("PONo")
                    pSupplier = ds.Tables(0).Rows(xi)("SupplierID")
                    pPeriod = ds.Tables(0).Rows(xi)("Period")
                    pDel = ds.Tables(0).Rows(xi)("ForwarderID")
                    pOrderNo1 = ds.Tables(0).Rows(xi)("OrderNo1")
                    ls_ExcelCls = ds.Tables(0).Rows(xi)("ExcelCls")

                    Dim dsEmail As New DataSet
                    dsEmail = EmailToEmailCCPOEmergency(pAffCode, "PASI", pSupplier)
                    '1 CC Affiliate'2 CC PASI'3 CC & TO Supplier
                    For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(i)("affiliatepocc")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("affiliatepocc")
                        End If
                        If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                            fromEmail = dsEmail.Tables(0).Rows(i)("toEmail")
                        End If
                        If receiptEmail = "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(i)("affiliatepoto")
                        Else
                            receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(i)("affiliatepoto")
                        End If
                    Next

                    receiptCCEmail = Replace(receiptCCEmail, ",", ";")
                    receiptEmail = Replace(receiptEmail, ",", ";")

                    temp_Filename1 = ""
                    temp_Filename2 = ""

                    If ls_ExcelCls = "3" Then
                        booSplit = True
                    Else
Split:
                        If ls_ExcelCls = "3" Then
                            pOrderNo1 = ds.Tables(0).Rows(xi)("SplitReffPONo")
                        End If
                        booSplit = False
                    End If

                    'Create Excel File
                    Dim fi As New FileInfo(Trim(txtAttachmentDOM.Text) & "\Template PO Export (Emergency).xlsm") 'File dari Local

                    If Not fi.Exists Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because file Excel isn't Found" & vbCrLf & _
                                        rtbProcess.Text
                        Exit Sub
                    End If
                    NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Template PO Export (Emergency).xlsm"

                    Dim ls_file As String = NewFileCopy
                    ExcelBook = xlApp.Workbooks.Open(ls_file)
                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)


                    ExcelSheet.Range("H1").Value = "POEE"
                    ExcelSheet.Range("H2").Value = fromEmail
                    ExcelSheet.Range("H3").Value = pAffCode
                    ExcelSheet.Range("H4").Value = pDel
                    ExcelSheet.Range("H5").Value = pSupplier

                    ExcelSheet.Range("S1").Value = pPONo

                    ExcelSheet.Range("Y2").Value = ""

                    'Order No
                    ExcelSheet.Range("I9").Value = pOrderNo1

                    'PO Date
                    ExcelSheet.Range("AE9").Value = If(IsDBNull(ds.Tables(0).Rows(xi)("Period")), "", Format(ds.Tables(0).Rows(xi)("Period"), "dd-MMM-yyyy"))
                    'Commercial Cls
                    ExcelSheet.Range("AE11").Value = IIf(ds.Tables(0).Rows(xi)("CommercialCls") = "1", "YES", "NO")

                    'To
                    ExcelSheet.Range("I11").Value = pSupplier
                    Dim dsSupp As New DataSet
                    dsSupp = Supplier(Trim(pSupplier))
                    ExcelSheet.Range("I12").Value = dsSupp.Tables(0).Rows(0)("Address")
                    ExcelSheet.Range("I12:X14").WrapText = True

                    'Buyer
                    ExcelSheet.Range("I16").Value = pAffCode
                    Dim dsAffp As New DataSet
                    dsAffp = Affiliate(Trim(pAffCode))
                    ExcelSheet.Range("I17").Value = dsAffp.Tables(0).Rows(0)("Address")
                    ExcelSheet.Range("I17:X19").WrapText = True

                    'Delivery To
                    ExcelSheet.Range("AE13").Value = pDel
                    Dim dsDelivery As New DataSet
                    dsDelivery = Forwarder(Trim(pDel))
                    ExcelSheet.Range("AE14").Value = dsDelivery.Tables(0).Rows(0)("Address")
                    ExcelSheet.Range("AE14:AT16").WrapText = True

                    ExcelSheet.Range("B38").Interior.Color = Color.White
                    ExcelSheet.Range("B38").Font.Color = Color.Black

                    dsDetail = bindDataDetailEmergency(pPeriod, pAffCode, pPONo, pSupplier, pOrderNo1)

                    If dsDetail.Tables(0).Rows.Count > 0 Then
                        For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                            'Header
                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).Merge() 'No
                            ExcelSheet.Range("D" & i + 37 & ": H" & i + 37).Merge() 'Part No
                            ExcelSheet.Range("I" & i + 37 & ": P" & i + 37).Merge() 'Part Name
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).Merge() 'UOM
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).Merge() 'MOQ
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).Merge() ' Total Order
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Merge()

                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Merge() 'ETD Supplier

                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("NoUrut"))
                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("D" & i + 37 & ": H" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                            ExcelSheet.Range("I" & i + 37 & ": P" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName"))
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("UOM"))
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("MOQ"))
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).NumberFormat = "#,##0"

                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).NumberFormat = "#,##0"

                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Value = Trim(ds.Tables(0).Rows(xi)("ETDVendor1"))
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).NumberFormat = "yyy-MM-dd"

                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).NumberFormat = "#,##0"

                            'ExcelSheet.Range("AR" & i + 37 & ": AU" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Forecast1"))
                            'ExcelSheet.Range("AR" & i + 37 & ": AU" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            'ExcelSheet.Range("AR" & i + 37 & ": AU" & i + 37).NumberFormat = "#,##0"
                            'ExcelSheet.Range("AV" & i + 37 & ": AY" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Forecast2"))
                            'ExcelSheet.Range("AV" & i + 37 & ": AY" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            'ExcelSheet.Range("AV" & i + 37 & ": AY" & i + 37).NumberFormat = "#,##0"
                            'ExcelSheet.Range("AZ" & i + 37 & ": BC" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Forecast3"))
                            'ExcelSheet.Range("AZ" & i + 37 & ": BC" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            'ExcelSheet.Range("AZ" & i + 37 & ": BC" & i + 37).NumberFormat = "#,##0"

                            DrawAllBorders(ExcelSheet.Range("B" & i + 37 & ": AI" & i + 37))
                            ExcelSheet.Range("AD" & i + 37 & ": AI" & i + 37).Interior.Color = ColorYellow
                            ExcelSheet.Range("B" & i + 37 & ": AD" & i + 37).Interior.Color = RGB(217, 217, 217)
                            'ExcelSheet.Range("AR" & i + 37 & ": BC" & i + 37).Interior.Color = RGB(217, 217, 217)
                        Next
                        'Else
                        '    xlApp.Workbooks.Close()
                        '    xlApp.Quit()
                        '    GoTo satuTemplate
                    End If

                    ExcelSheet.Range("B" & i + 37).Value = "E"
                    ExcelSheet.Range("B" & i + 37).Interior.Color = Color.Black
                    ExcelSheet.Range("B" & i + 37).Font.Color = Color.White

                    xlApp.DisplayAlerts = False

                    If temp_Filename1 = "" Then
                        If Trim(pPONo) = Trim(pOrderNo1) Then
                            temp_Filename1 = "PO Emergency " & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"
                        Else
                            temp_Filename1 = "PO Emergency " & Trim(pPONo) & "-Split (" & Trim(pOrderNo1) & ")-" & Trim(pSupplier) & ".xlsm"
                        End If

                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\" & temp_Filename1)
                    Else
                        If Trim(pPONo) = Trim(pOrderNo1) Then
                            temp_Filename2 = "PO Emergency " & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"
                        Else
                            temp_Filename2 = "PO Emergency " & Trim(pPONo) & "-Split (" & Trim(pOrderNo1) & ")-" & Trim(pSupplier) & ".xlsm"
                        End If

                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\" & temp_Filename2)
                    End If

                    If dsDetail.Tables(0).Rows.Count = 0 Then
                        temp_Filename2 = ""
                    End If

                    'Dim temp_Filename As String = "PO Emergency " & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"
                    'ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\" & temp_Filename)
                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & temp_Filename1 & temp_Filename2 & " OK. " & vbCrLf & _
                    rtbProcess.Text

                    If booSplit Then GoTo Split
                    'satuTemplate:

                    If ls_ExcelCls = "3" Then
                        pOrderNo1 = ds.Tables(0).Rows(xi)("OrderNo1")
                    End If

                    If sendEmailPOtoSupllierEmergency(temp_Filename1, temp_Filename2, pOrderNo1) = False Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & temp_Filename1 & temp_Filename2 & " NG. " & vbCrLf & _
                        rtbProcess.Text
                        GoTo keluar
                    Else
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & temp_Filename1 & temp_Filename2 & " OK. " & vbCrLf & _
                        rtbProcess.Text
                    End If


                    'Call sendEmailPOtoAffiliateMonthly(pOrderNo1)

                    Call UpdateExcelPOEmergency(True, pAffCode, pPONo, pOrderNo1, pSupplier)

                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    Thread.Sleep(500)
keluar:
                    xlApp.Workbooks.Close()
                    xlApp.Quit()


                    Thread.Sleep(500)
                Next
            Else
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because there is nothing PO to send " & vbCrLf & _
                    rtbProcess.Text
            End If
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If

        End Try
    End Sub

    Private Sub Excel_CustomerDeliveryConfirmation_SplitDelivery(dt As DataTable, xlApp As Excel.Application, ExcelBook As Excel.Workbook, ExcelSheet As Excel.Worksheet)
        Dim templateFile = ""
        Dim ls_Filename1 = "", ls_Filename2 = ""
        Dim receiptEmail = "", receiptCCEmail = ""
        Dim tmpOrderNo = ""

        Dim OrderNo1 = ""
        Dim OrderNo2 = ""
        Dim AffiliateID = Trim(dt.Rows(0)("AffiliateID"))
        Dim SupplierID = Trim(dt.Rows(0)("SupplierID"))
        Dim ForwarderID = "" 'Trim(dt.Rows(0)("ForwarderID"))
        Dim PONo = Trim(dt.Rows(0)("PONo"))
        Dim OrderNo = "" 'Trim(dt.Rows(0)("orderNo"))
        Dim SJDeliverySplit = Trim(dt.Rows(0)("SplitDelivery")) 'Replace(dt.Rows(0)("SplitDelivery")), "'", "''")

        Dim dsEmail As New DataSet
        dsEmail = EmailToEmailCCKanban_Export(AffiliateID, SupplierID)
        '1 CC Affiliate
        '2 CC PASI
        '3 CC & TO Supplier
        For i = 0 To dsEmail.Tables(0).Rows.Count - 1
            If receiptCCEmail = "" Then
                receiptCCEmail = dsEmail.Tables(0).Rows(i)("KanbanCC")
            Else
                receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("KanbanCC")
            End If

            If dsEmail.Tables(0).Rows(i)("KanbanTO") <> "" Then
                receiptEmail = dsEmail.Tables(0).Rows(i)("KanbanTO")
            End If
        Next

        Dim ls_sql As String = "Exec sp_DeliverySplitBatch_Select '" + AffiliateID + "', '" + SupplierID + "', '" + PONo + "', '" + SJDeliverySplit + "' "
        Dim dtSplitDelivery As New DataTable
        dtSplitDelivery = cls.uf_GetDataSet(ls_sql).Tables(0)

        If dtSplitDelivery.Rows.Count > 2 Then
            Throw New Exception("Data Excel Split More than 2, Please Check : " & ls_sql)
        End If

        For i = 0 To dtSplitDelivery.Rows.Count - 1
            OrderNo = dtSplitDelivery.Rows(i)("OrderNo").ToString()
            ForwarderID = dtSplitDelivery.Rows(0)("OldForwarderID").ToString() 'pasti yg pertama aja
            tmpOrderNo = tmpOrderNo & IIf(tmpOrderNo = "", "", " & ") & OrderNo

            templateFile = Trim(txtAttachmentDOM.Text) & "\Template Customer Delivery Confirmation.xlsm"
            Dim fi As New FileInfo(templateFile)

            If Not fi.Exists Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Customer Delivery Confirmation STOPPED, because File Excel isn't Found " & vbCrLf & _
                                rtbProcess.Text
                Exit Sub
            End If

            ExcelBook = xlApp.Workbooks.Open(templateFile)
            ExcelSheet = CType(ExcelBook.Worksheets(1), Excel.Worksheet)

            'Start Create Excel for Header
            ExcelSheet.Range("H2").Value = receiptEmail.Trim
            ExcelSheet.Range("H3").Value = AffiliateID
            ExcelSheet.Range("H4").Value = dtSplitDelivery.Rows(i)("ForwarderID").ToString()
            ExcelSheet.Range("H5").Value = SupplierID

            ExcelSheet.Range("AP11").Value = IIf(dtSplitDelivery.Rows(i)("CommercialCls").ToString() = "0", "NO", "YES")

            ExcelSheet.Range("I11:X11").Value = dtSplitDelivery.Rows(i)("SupplierName").ToString()
            ExcelSheet.Range("I12:X15").Value = dtSplitDelivery.Rows(i)("SupplierAddress").ToString()

            ExcelSheet.Range("I19:X19").Value = dtSplitDelivery.Rows(i)("ForwarderName").ToString()
            ExcelSheet.Range("I20:X22").Value = dtSplitDelivery.Rows(i)("ForwarderAddress").ToString()

            ExcelSheet.Range("I23:X23").Value = "ATTN : " & dtSplitDelivery.Rows(i)("attn").ToString() & "     TELP : " & dtSplitDelivery.Rows(i)("telp").ToString()

            ExcelSheet.Range("AE19:AT19").Value = dtSplitDelivery.Rows(i)("ConsigneeName").ToString()
            ExcelSheet.Range("AE20:AT22").Value = dtSplitDelivery.Rows(i)("ConsigneeAddress").ToString()

            ExcelSheet.Range("AE11:AI11").Value = Format((dtSplitDelivery.Rows(i)("Period")), "yyyy-MM")

            ExcelSheet.Range("AE13:AI13").Value = OrderNo
            ExcelSheet.Range("AE15:AI15").Value = PONo

            ExcelSheet.Range("AE17:AI17").Value = Format((dtSplitDelivery.Rows(i)("ETDVendor")), "yyyy-MM-dd")

            'Ending Create Excel for Header

            ls_sql = "Exec sp_DeliverySplitBatch_Select_Detail '" + AffiliateID + "', '" + SupplierID + "', '" + PONo + "', '" + OrderNo + "' "
            Dim dtSplitDelivery_Detail As New DataTable
            dtSplitDelivery_Detail = cls.uf_GetDataSet(ls_sql).Tables(0)

            'Start Create Excel for Detail
            With ExcelSheet
                For j = 0 To dtSplitDelivery_Detail.Rows.Count - 1
                    .Range("B" & j + 34 & ": C" & j + 34).Merge()   'No
                    .Range("D" & j + 34 & ": H" & j + 34).Merge()   'Order No
                    .Range("I" & j + 34 & ": M" & j + 34).Merge()   'Part No
                    .Range("N" & j + 34 & ": V" & j + 34).Merge()   'Part Name
                    .Range("W" & j + 34 & ": Y" & j + 34).Merge()   'Label No From
                    .Range("Z" & j + 34 & ": AB" & j + 34).Merge()  'Label No To
                    .Range("AC" & j + 34 & ": AD" & j + 34).Merge() 'UOM
                    .Range("AE" & j + 34 & ": AF" & j + 34).Merge() 'Qty/Box
                    .Range("AG" & j + 34 & ": AJ" & j + 34).Merge() 'Delivery Plan Qty
                    .Range("AK" & j + 34 & ": AN" & j + 34).Merge() 'Remaining Qty
                    .Range("AO" & j + 34 & ": AR" & j + 34).Merge() 'Delivery Qty
                    .Range("AS" & j + 34 & ": AV" & j + 34).Merge() 'Total Box

                    .Range("B" & j + 34 & ": C" & j + 34).Value = j + 1                                                     'No
                    .Range("D" & j + 34 & ": H" & j + 34).Value = OrderNo                                                   'Order No
                    .Range("I" & j + 34 & ": M" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("PartNo").ToString()       'Part No
                    .Range("N" & j + 34 & ": V" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("PartName").ToString()     'Part Name
                    .Range("W" & j + 34 & ": Y" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("LabelNo1").ToString()     'Label No From
                    .Range("Z" & j + 34 & ": AB" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("LabelNo2").ToString()    'Label No To
                    .Range("AC" & j + 34 & ": AD" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("UOM").ToString()        'UOM
                    .Range("AE" & j + 34 & ": AF" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("MOQ").ToString()        'Qty/Box
                    .Range("AG" & j + 34 & ": AJ" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("DelivQty").ToString()   'Delivery Plan Qty
                    .Range("AG" & j + 34 & ": AJ" & j + 34).NumberFormat = "#,##0"

                    .Range("AK" & j + 34 & ": AN" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("DelivQty").ToString()   'Remaining Qty
                    .Range("AK" & j + 34 & ": AN" & j + 34).NumberFormat = "#,##0"

                    .Range("AO" & j + 34 & ": AR" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("DelivQty").ToString()   'Delivery Qty
                    .Range("AO" & j + 34 & ": AR" & j + 34).NumberFormat = "#,##0"

                    .Range("AS" & j + 34 & ": AV" & j + 34).Value = dtSplitDelivery_Detail.Rows(j)("TotalPOQty").ToString() 'Total Box
                    .Range("AS" & j + 34 & ": AV" & j + 34).NumberFormat = "#,##0"
                Next

                'Edit Cell Style

                .Range("B35").Interior.Color = Color.White
                .Range("B35").Font.Color = Color.Black
                .Range("B" & dtSplitDelivery_Detail.Rows.Count + 34).Value = "E"
                .Range("B" & dtSplitDelivery_Detail.Rows.Count + 34).Interior.Color = Color.Black
                .Range("B" & dtSplitDelivery_Detail.Rows.Count + 34).Font.Color = Color.White

                DrawAllBorders(.Range("B34" & ": AV" & dtSplitDelivery_Detail.Rows.Count - 1 + 34))
                .Range("AM34" & ": AP" & dtSplitDelivery_Detail.Rows.Count - 1 + 34).Interior.Color = Color.Yellow
                .Range("W34" & ": AB" & dtSplitDelivery_Detail.Rows.Count - 1 + 34).Interior.Color = Color.Yellow

                Dim xlRange As Microsoft.Office.Interop.Excel.Range
                xlRange = .Range("W34:AB" & dtSplitDelivery_Detail.Rows.Count - 1 + 34)
                xlRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = .Range("B34:C" & dtSplitDelivery_Detail.Rows.Count - 1 + 34)
                xlRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter


            End With
            'Ending Create Excel for Detail

            xlApp.DisplayAlerts = False

            If i = 0 Then
                OrderNo1 = OrderNo
                ls_Filename1 = "\DELIVERY CONFIRMATION-" & PONo & " Split (" & OrderNo & ")-" & SupplierID & ".xlsm"
                ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & ls_Filename1)
            ElseIf i = 1 Then
                OrderNo2 = OrderNo
                ls_Filename2 = "\DELIVERY CONFIRMATION-" & PONo & " Split (" & OrderNo & ")-" & SupplierID & ".xlsm"
                ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & ls_Filename2)
            End If

            xlApp.Workbooks.Close()
            xlApp.Quit()
        Next

        If sendEmailPASI_EXPORT("Delivery Confirmation", "Delivery Confirmation : " & Trim(PONo) & "-" & SupplierID, SupplierID, ls_Filename1, ls_Filename2, PONo, tmpOrderNo) = True Then
            If sendEmailPASI_EXPORTForwarder_Information(SJDeliverySplit, AffiliateID, ForwarderID) = True Then
                If OrderNo1 <> "" Then Call UpdateStatusPOExport(AffiliateID, SupplierID, PONo, OrderNo1)
                If OrderNo2 <> "" Then Call UpdateStatusPOExport(AffiliateID, SupplierID, PONo, OrderNo2)
            End If
        End If
    End Sub

    Private Sub Excel_CustomerDeliveryConfirmation()
        'On Error GoTo ErrHandler
        Dim strFileSize As String = ""
        Dim ls_sql As String = ""

        Dim xlApp = New Excel.Application
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        'Dim ExcelSheet2 As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim sheetNumber2 As Integer = 3
        Dim i As Integer, xi As Integer
        'Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        'Dim fileTocopy As String
        Dim NewFileCopy As String
        'Dim NewFileCopyas As String

        Dim KTime1 As String = ""
        Dim KTime2 As String = ""
        Dim KTime3 As String = ""
        Dim KTime4 As String = ""
        'Dim pNamaFile As String = ""

        'Dim jkanbanno As String
        'Dim jQty As Long
        'Dim jQtyBox As Long
        'Dim jQtyPallet As Long

        Dim ds As New DataSet
        Dim dsSplit As New DataSet
        Dim dsHeader As New DataSet
        Dim dsDetail As New DataSet
        Dim dsETAETD As New DataSet
        Dim dsDetailDelivery As New DataSet

        Dim ls_SJ As String = ""
        Dim ls_Supplier As String = ""
        Dim ls_supplierName As String = ""
        Dim ls_supplierAdd As String = ""
        Dim ls_delivery As String = ""
        Dim ls_Attn As String = ""
        Dim ls_Telp As String = ""
        Dim ls_DeliveryName As String = ""
        Dim ls_deliveryAdd As String = ""
        Dim ls_orderNo As String = ""
        Dim ls_orderNoSplit As String = ""
        Dim ls_ETDV As Date
        Dim ls_ETDP As Date
        Dim ls_ETAP As Date
        Dim ls_ETAF As Date
        Dim ls_Aff As String = ""
        Dim ls_Consignee As String = ""
        Dim ls_AFFName As String = ""
        Dim ls_AffADD As String
        Dim ls_PEriod As String = ""
        Dim ls_FinalApprovalCls As String = ""
        Dim ls_orderNoReff As String = ""
        Dim ls_sts As String = ""

        Dim i_loop As Long
        Dim pFilename_e As String = ""
        Dim pFileName_e2 As String = ""
        Dim dsR As New DataSet
        Dim ls_Comercial As String = ""
        Dim sDefect As String = ""
        'Dim booSplit As Boolean = False

        Try
            ls_sql = "  --Normal  " & vbCrLf &
                     "  select distinct sts = '0', Consignee = isnull(MA.ConsigneeCode,''),sts = 'remaining', attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), PME.Period, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') SUPPAddress,   " & vbCrLf &
                     "  PME.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  FWDAddress, ETDVendor1 as ETDVendor, ETDPort1 as ETDPort, ETAPort1 as ETAPort, ETAFactory1 as ETAFactory,  " & vbCrLf &
                     "  PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'') AFFAddress, DefectRecQty = 0, '1' FinalApprovalCls, ISNULL(PME.SplitReffPONo, '') SplitReffPONo, PME.CommercialCls, SplitDelivery = ISNULL(SplitDelivery,'') " & vbCrLf &
                     "  From PO_Master_Export PME   " & vbCrLf &
                     "  LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID  " & vbCrLf &
                     "  LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID  " & vbCrLf &
                     "  LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID  " & vbCrLf &
                     "  where isnull(FinalApprovalCls,0) IN ('1', '3', '4') and isnull(PONO,'') <> '' OR OrderNo1 = 'YC5101Y-7360G85904-2' "

            'MdlConn.ReadConnection()
            ds = cls.uf_GetDataSet(ls_sql)

            Dim ls_file As String = ""

            For i_loop = 0 To ds.Tables(0).Rows.Count - 1
                'Fungsi Split Delivery
                If Trim(ds.Tables(0).Rows(i_loop)("FinalApprovalCls")) = "1" And Trim(ds.Tables(0).Rows(i_loop)("SplitDelivery")) <> "" Then
                    Excel_CustomerDeliveryConfirmation_SplitDelivery(ds.Tables(0).Rows(i_loop).Table, xlApp, ExcelBook, ExcelSheet)
                Else
                    pFilename_e = "" : pFileName_e2 = ""

                    '================================DELIVERY CONFIRMATION==================================
                    Dim fi As New FileInfo(Trim(txtAttachmentDOM.Text) & "\Template Customer Delivery Confirmation.xlsm")

                    If Not fi.Exists Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Customer Delivery Confirmation STOPPED, because File Excel isn't Found " & vbCrLf & _
                                        rtbProcess.Text
                        Exit Sub
                    End If

                    ls_Comercial = Trim(ds.Tables(0).Rows(i_loop)("CommercialCls"))

                    ls_FinalApprovalCls = Trim(ds.Tables(0).Rows(i_loop)("FinalApprovalCls"))
                    If ls_FinalApprovalCls = "3" And Trim(ds.Tables(0).Rows(i_loop)("SplitReffPONo")) <> "" Then
                        'booSplit = True

                        ls_sql = " select distinct sts = '0', Consignee = isnull(MA.ConsigneeCode,''),attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), PME.Period, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') SUPPAddress,  " & vbCrLf & _
                                 " PME.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  FWDAddress, ETDVendor1 as ETDVendor, ETDPort1 as ETDPort, ETAPort1 as ETAPort, ETAFactory1 as ETAFactory, " & vbCrLf & _
                                 " PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'') AFFAddress, DefectRecQty = 0, isnull(FinalApprovalCls,0) FinalApprovalCls, ISNULL(PME.SplitReffPONo, '') SplitReffPONo " & vbCrLf & _
                                 " FROM PO_Master_Export PME  " & vbCrLf & _
                                 " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                                 " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID " & vbCrLf & _
                                 " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                                 " WHERE PME.PONo = '" & Trim(ds.Tables(0).Rows(i_loop)("PONo")) & "' " & vbCrLf & _
                                 " AND PME.AffiliateID = '" & Trim(ds.Tables(0).Rows(i_loop)("AffiliateID")) & "' " & vbCrLf & _
                                 " AND PME.SupplierID = '" & Trim(ds.Tables(0).Rows(i_loop)("SupplierID")) & "' " & vbCrLf & _
                                 " AND PME.SplitReffPONo = '" & Trim(ds.Tables(0).Rows(i_loop)("SplitReffPONo")) & "' "

                        dsSplit = cls.uf_GetDataSet(ls_sql)

                        ls_SJ = ""
                        ls_orderNoSplit = Trim(dsSplit.Tables(0).Rows(0)("OrderNo"))
                        ls_Supplier = Trim(dsSplit.Tables(0).Rows(0)("supplierID"))
                        ls_supplierName = Trim(dsSplit.Tables(0).Rows(0)("suppliername"))
                        ls_supplierAdd = Trim(dsSplit.Tables(0).Rows(0)("SuppAddress"))
                        ls_delivery = Trim(dsSplit.Tables(0).Rows(0)("ForwarderID"))
                        ls_DeliveryName = Trim(dsSplit.Tables(0).Rows(0)("ForwarderName"))
                        ls_deliveryAdd = Trim(dsSplit.Tables(0).Rows(0)("FWDAddress"))
                        ls_orderNo = Trim(dsSplit.Tables(0).Rows(0)("PONo"))
                        ls_ETDV = Format((dsSplit.Tables(0).Rows(0)("ETDVendor")), "yyyy-MM-dd")
                        ls_ETDP = Format((dsSplit.Tables(0).Rows(0)("ETDPort")), "yyyy-MM-dd")
                        ls_ETAP = Format((dsSplit.Tables(0).Rows(0)("ETAPort")), "yyyy-MM-dd")
                        ls_ETAF = Format((dsSplit.Tables(0).Rows(0)("ETAFactory")), "yyyy-MM-dd")
                        ls_Aff = Trim(dsSplit.Tables(0).Rows(0)("AFF"))
                        ls_Consignee = Trim(dsSplit.Tables(0).Rows(0)("Consignee"))
                        ls_AFFName = Trim(dsSplit.Tables(0).Rows(0)("AFFName"))
                        ls_AffADD = Trim(dsSplit.Tables(0).Rows(0)("AFFAddress"))
                        ls_PEriod = Format((dsSplit.Tables(0).Rows(0)("Period")), "yyyy-MM")
                        ls_Attn = Trim(dsSplit.Tables(0).Rows(0)("attn"))
                        ls_Telp = Trim(dsSplit.Tables(0).Rows(0)("telp"))
                        ls_sts = Trim(dsSplit.Tables(0).Rows(0)("sts"))
                    Else
Split:
                        'booSplit = False
                        ls_SJ = ""
                        ls_orderNoSplit = Trim(ds.Tables(0).Rows(i_loop)("OrderNo"))
                        ls_Supplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))
                        ls_supplierName = Trim(ds.Tables(0).Rows(i_loop)("suppliername"))
                        ls_supplierAdd = Trim(ds.Tables(0).Rows(i_loop)("SuppAddress"))
                        ls_delivery = Trim(ds.Tables(0).Rows(i_loop)("ForwarderID"))
                        ls_DeliveryName = Trim(ds.Tables(0).Rows(i_loop)("ForwarderName"))
                        ls_deliveryAdd = Trim(ds.Tables(0).Rows(i_loop)("FWDAddress"))
                        ls_orderNo = Trim(ds.Tables(0).Rows(i_loop)("PONo"))
                        ls_ETDV = Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "yyyy-MM-dd")
                        ls_ETDP = Format((ds.Tables(0).Rows(i_loop)("ETDPort")), "yyyy-MM-dd")
                        ls_ETAP = Format((ds.Tables(0).Rows(i_loop)("ETAPort")), "yyyy-MM-dd")
                        ls_ETAF = Format((ds.Tables(0).Rows(i_loop)("ETAFactory")), "yyyy-MM-dd")
                        ls_Aff = Trim(ds.Tables(0).Rows(i_loop)("AFF"))
                        ls_Consignee = Trim(ds.Tables(0).Rows(i_loop)("Consignee"))
                        ls_AFFName = Trim(ds.Tables(0).Rows(i_loop)("AFFName"))
                        ls_AffADD = Trim(ds.Tables(0).Rows(i_loop)("AFFAddress"))
                        ls_PEriod = Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy-MM")
                        ls_Attn = Trim(ds.Tables(0).Rows(i_loop)("attn"))
                        ls_Telp = Trim(ds.Tables(0).Rows(i_loop)("telp"))
                        ls_sts = Trim(ds.Tables(0).Rows(i_loop)("sts"))

                        If ls_FinalApprovalCls = "1" Then
                            Call InsertPrintLabel(ls_orderNo, ls_orderNoSplit, ls_Aff, ls_Supplier, Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "yyyy"), Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "MM"))
                        End If
                    End If

                    dsDetailDelivery = BidDataDeliveryConfirm(ls_orderNo, ls_orderNoSplit, ls_Aff, ls_Supplier)
                    If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                        Dim dsEmail As New DataSet
                        dsEmail = EmailToEmailCCKanban_Export(ls_Aff, ls_Supplier)
                        '1 CC Affiliate
                        '2 CC PASI
                        '3 CC & TO Supplier
                        For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                            If receiptCCEmail = "" Then
                                receiptCCEmail = dsEmail.Tables(0).Rows(i)("KanbanCC")
                            Else
                                receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("KanbanCC")
                            End If

                            If dsEmail.Tables(0).Rows(i)("KanbanTO") <> "" Then
                                receiptEmail = dsEmail.Tables(0).Rows(i)("KanbanTO")
                            End If
                        Next

                        Dim k As Long
                        Dim dsAffiliate As New DataSet
                        dsAffiliate = Affiliate(Trim(ls_Aff))

                        Dim dsSupplier As New DataSet
                        dsSupplier = Supplier(Trim(ls_Supplier))

                        Dim status As Boolean
                        status = True

                        If dsDetailDelivery.Tables(0).Rows.Count = 0 Then
                            status = False
                        Else
                            If ls_FinalApprovalCls = "3" And Trim(ds.Tables(0).Rows(i_loop)("SplitReffPONo")) = "" Then
                                status = False
                            Else
                                status = True
                            End If
                        End If

                        If status = True Then
                            NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Template Customer Delivery Confirmation.xlsm"
                            ls_file = NewFileCopy

                            ExcelBook = xlApp.Workbooks.Open(ls_file)
                            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                            ExcelSheet.Range("H2").Value = receiptEmail.Trim
                            ExcelSheet.Range("H3").Value = ls_Aff.Trim
                            ExcelSheet.Range("H4").Value = ls_delivery.Trim
                            ExcelSheet.Range("H5").Value = ls_Supplier.Trim

                            If ls_Comercial = "0" Then
                                ExcelSheet.Range("AP11").Value = "NO"
                            Else
                                ExcelSheet.Range("AP11").Value = "YES"
                            End If

                            ExcelSheet.Range("I11:X11").Value = ls_supplierName.Trim
                            ExcelSheet.Range("I12:X15").Value = ls_supplierAdd.Trim

                            ExcelSheet.Range("I19:X19").Value = ls_DeliveryName.Trim
                            ExcelSheet.Range("I20:X22").Value = ls_deliveryAdd.Trim
                            ExcelSheet.Range("I23:X23").Value = "ATTN : " & Trim(ls_Attn) & "     TELP : " & Trim(ls_Telp)

                            ExcelSheet.Range("AE19:AT19").Value = Trim(dsAffiliate.Tables(0).Rows(0)("ConsigneeName"))
                            ExcelSheet.Range("AE20:AT22").Value = Trim(dsAffiliate.Tables(0).Rows(0)("ConsigneeAddress"))

                            ExcelSheet.Range("AE11:AI11").Value = ls_PEriod

                            If newDN = False Then
                                ExcelSheet.Range("AE13:AI13").Value = ls_orderNoSplit.Trim
                                If ls_orderNo <> ls_orderNoSplit Then
                                    ExcelSheet.Range("AE15:AI15").Value = ls_orderNo.Trim
                                End If
                            Else
                                ExcelSheet.Range("AE13:AI13").Value = ls_orderNoSplit.Trim
                                If ls_orderNo <> ls_orderNoSplit Then
                                    ExcelSheet.Range("AE15:AI15").Value = ls_orderNo.Trim
                                End If
                            End If

                            ExcelSheet.Range("AE17:AI17").Value = ls_ETDV

                            k = 0
                            For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                                Dim newKanbanNo As String = ""
                                If newDN = False Then
                                    ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Merge()
                                    ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Merge()
                                    ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Merge()
                                    ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Merge()
                                    ExcelSheet.Range("W" & k + 34 & ": Z" & k + 34).Merge()
                                    ExcelSheet.Range("AA" & k + 34 & ": AB" & k + 34).Merge()
                                    ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Merge()
                                    ExcelSheet.Range("AE" & k + 34 & ": AH" & k + 34).Merge()
                                    ExcelSheet.Range("AI" & k + 34 & ": AL" & k + 34).Merge()
                                    ExcelSheet.Range("AM" & k + 34 & ": AP" & k + 34).Merge()
                                    ExcelSheet.Range("AQ" & k + 34 & ": AT" & k + 34).Merge()

                                    ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Value = k + 1
                                    ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Value = ls_orderNo '
                                    ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Partno"))
                                    ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("PartName")
                                    ExcelSheet.Range("W" & k + 34 & ": Z" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("labelno") ''
                                    ExcelSheet.Range("AA" & k + 34 & ": AB" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("UOM")
                                    ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("MOQ")
                                    ExcelSheet.Range("AE" & k + 34 & ": AH" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                                    ExcelSheet.Range("AE" & i + 34 & ": AH" & i + 34).NumberFormat = "#,##0"
                                    ExcelSheet.Range("AI" & k + 34 & ": AL" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                                    ExcelSheet.Range("AI" & i + 34 & ": AL" & i + 34).NumberFormat = "#,##0"
                                    ExcelSheet.Range("AM" & k + 34 & ": AP" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                                    ExcelSheet.Range("AM" & i + 34 & ": AP" & i + 34).NumberFormat = "#,##0"
                                    ExcelSheet.Range("AQ" & k + 34 & ": AT" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalPOQty")
                                    ExcelSheet.Range("AQ" & i + 34 & ": AT" & i + 34).NumberFormat = "#,##0"
                                Else
                                    ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Merge()
                                    ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Merge()
                                    ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Merge()
                                    ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Merge()
                                    ExcelSheet.Range("W" & k + 34 & ": Y" & k + 34).Merge()
                                    ExcelSheet.Range("Z" & k + 34 & ": AB" & k + 34).Merge()
                                    ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Merge()
                                    ExcelSheet.Range("AE" & k + 34 & ": AF" & k + 34).Merge()
                                    ExcelSheet.Range("AG" & k + 34 & ": AJ" & k + 34).Merge()
                                    ExcelSheet.Range("AK" & k + 34 & ": AN" & k + 34).Merge()
                                    ExcelSheet.Range("AO" & k + 34 & ": AR" & k + 34).Merge()
                                    ExcelSheet.Range("AS" & k + 34 & ": AV" & k + 34).Merge()

                                    ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Value = k + 1
                                    ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Value = ls_orderNo '
                                    ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Partno"))
                                    ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("PartName")
                                    ExcelSheet.Range("W" & k + 34 & ": Y" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelno1"))
                                    ExcelSheet.Range("Z" & k + 34 & ": AB" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelno2"))
                                    ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("UOM")
                                    ExcelSheet.Range("AE" & k + 34 & ": AF" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("MOQ")
                                    ExcelSheet.Range("AG" & k + 34 & ": AJ" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                                    ExcelSheet.Range("AG" & i + 34 & ": AJ" & i + 34).NumberFormat = "#,##0"
                                    ExcelSheet.Range("AK" & k + 34 & ": AN" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                                    ExcelSheet.Range("AK" & i + 34 & ": AN" & i + 34).NumberFormat = "#,##0"
                                    ExcelSheet.Range("AO" & k + 34 & ": AR" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                                    ExcelSheet.Range("AO" & i + 34 & ": AR" & i + 34).NumberFormat = "#,##0"
                                    ExcelSheet.Range("AS" & k + 34 & ": AV" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalPOQty")
                                    ExcelSheet.Range("AS" & i + 34 & ": AV" & i + 34).NumberFormat = "#,##0"
                                End If
                                k = k + 1
                            Next

                            ExcelSheet.Range("B35").Interior.Color = Color.White
                            ExcelSheet.Range("B35").Font.Color = Color.Black
                            ExcelSheet.Range("B" & k + 34).Value = "E"
                            ExcelSheet.Range("B" & k + 34).Interior.Color = Color.Black
                            ExcelSheet.Range("B" & k + 34).Font.Color = Color.White

                            k = k - 1
                            DrawAllBorders(ExcelSheet.Range("B34" & ": AV" & k + 34))
                            ExcelSheet.Range("AM34" & ": AP" & k + 34).Interior.Color = Color.Yellow
                            ExcelSheet.Range("W34" & ": AB" & k + 34).Interior.Color = Color.Yellow

                            'Save ke Local
                            xlApp.DisplayAlerts = False

                            If pFilename_e = "" Then
                                If ls_orderNo <> ls_orderNoSplit Then
                                    ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm")
                                    pFilename_e = "\DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm"
                                Else
                                    ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm")
                                    pFilename_e = "\DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm"
                                End If
                            Else
                                If ls_orderNo <> ls_orderNoSplit Then
                                    ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm")
                                    pFileName_e2 = "\DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm"
                                Else
                                    ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm")
                                    pFileName_e2 = "\DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm"
                                End If
                            End If

                            xlApp.Workbooks.Close()
                            xlApp.Quit()

                            'If booSplit Then GoTo Split
                        End If
                    End If
                    '================================DELIVERY CONFIRMATION==================================

                    '=======================================LABEL===========================================
                    If ls_FinalApprovalCls = "1" Then
                        Dim fi2 As New FileInfo(Trim(txtAttachmentDOM.Text) & "\Print label2.xlsm")

                        If Not fi2.Exists Then
                            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Customer Delivery Confirmation STOPPED, because File Excel isn't Found " & vbCrLf & _
                                            rtbProcess.Text
                            Exit Sub
                        End If

                        dsDetailDelivery = BindDataLabelPrint(ls_orderNo, ls_orderNoSplit, ls_Aff, ls_Supplier)

                        If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                            Dim k As Long
                            Dim dsAffiliate As New DataSet
                            dsAffiliate = Affiliate(Trim(ls_Aff))

                            Dim dsSupplier As New DataSet
                            dsSupplier = Supplier(Trim(ls_Supplier))

                            Dim status As Boolean
                            status = True

                            If dsDetailDelivery.Tables(0).Rows.Count = 0 Then
                                status = False
                            Else
                                status = True
                            End If

                            If status = True Then
                                NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Print Label2.xlsm"
                                ls_file = NewFileCopy
                                ExcelBook = xlApp.Workbooks.Open(ls_file)
                                ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                                ExcelSheet.Range("I4:W4").Value = ls_AFFName
                                ExcelSheet.Range("I6:U6").Value = ls_DeliveryName
                                ExcelSheet.Range("I7:U9").Value = ls_deliveryAdd
                                ExcelSheet.Range("AC6:AQ6").Value = ls_supplierName
                                ExcelSheet.Range("AC7:AQ9").Value = ls_supplierAdd

                                ExcelSheet.Range("I12:P12").Value = ls_PEriod
                                ExcelSheet.Range("I14:P14").Value = Format(ls_ETDV, "dd-MMM-yyyy")

                                k = 0

                                For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                                    'For i = 0 To 3
                                    k = k
                                    Dim newKanbanNo As String = ""
                                    ExcelSheet.Range("B" & k + 20 & ": C" & k + 20).Merge()
                                    ExcelSheet.Range("D" & k + 20 & ": F" & k + 20).Merge()
                                    ExcelSheet.Range("G" & k + 20 & ": I" & k + 20).Merge()
                                    ExcelSheet.Range("J" & k + 20 & ": K" & k + 20).Merge()
                                    ExcelSheet.Range("L" & k + 20 & ": P" & k + 20).Merge()
                                    ExcelSheet.Range("Q" & k + 20 & ": X" & k + 20).Merge()
                                    ExcelSheet.Range("Y" & k + 20 & ": AD" & k + 20).Merge()
                                    ExcelSheet.Range("AE" & k + 20 & ": AI" & k + 20).Merge()
                                    ExcelSheet.Range("AJ" & k + 20 & ": AK" & k + 20).Merge()
                                    ExcelSheet.Range("AL" & k + 20 & ": AM" & k + 20).Merge()
                                    ExcelSheet.Range("AN" & k + 20 & ": AQ" & k + 20).Merge()
                                    ExcelSheet.Range("AR" & k + 20 & ": AT" & k + 20).Merge()

                                    ExcelSheet.Range("J" & k + 20 & ": K" & k + 20).Value = k + 1
                                    ExcelSheet.Range("D" & k + 20 & ": F" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("label1")
                                    ExcelSheet.Range("G" & k + 20 & ": I" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("label2")
                                    ExcelSheet.Range("L" & k + 20 & ": P" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("Partno")
                                    ExcelSheet.Range("Q" & k + 20 & ": X" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("Partname")
                                    ExcelSheet.Range("Y" & k + 20 & ": AD" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("labelNo")
                                    ExcelSheet.Range("AE" & k + 20 & ": AI" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("OrderNo")
                                    ExcelSheet.Range("AJ" & k + 20 & ": AK" & k + 20).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Uom"))

                                    ExcelSheet.Range("AL" & k + 20 & ": AM" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("QtyBox")
                                    ExcelSheet.Range("AN" & k + 20 & ": AQ" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("Qty")
                                    ExcelSheet.Range("AR" & k + 20 & ": AT" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("boxqty")
                                    ExcelSheet.Range("AU" & k + 20 & ": AU" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("DestinationPort")
                                    ExcelSheet.Range("AV" & k + 20 & ": AV" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("DestinationPoint")
                                    ExcelSheet.Range("AW" & k + 20 & ": AW" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("custname")
                                    ExcelSheet.Range("AX" & k + 20 & ": AX" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("custcode")
                                    ExcelSheet.Range("AY" & k + 20 & ": AY" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("consigneecode")

                                    k = k + 1
                                Next

                                ExcelSheet.Range("B21").Interior.Color = Color.White
                                ExcelSheet.Range("B21" & ": I" & k + 19).Interior.Color = Color.Yellow
                                ExcelSheet.Range("AN20" & ": AT" & k + 20).Font.Color = Color.Black
                                ExcelSheet.Range("AU20" & ": AY" & k + 20).Font.Color = Color.White

                                ExcelSheet.Range("D20" & ": J" & k + 20).HorizontalAlignment = Alignment.HorizontalCenterAlign
                                ExcelSheet.Range("AJ20" & ": AL" & k + 20).HorizontalAlignment = Alignment.HorizontalCenterAlign

                                ExcelSheet.Range("B21").Value = ""
                                ExcelSheet.Range("B21").Font.Color = Color.Black

                                ExcelSheet.Range("B" & k + 20).Value = "E"
                                ExcelSheet.Range("B" & k + 20).Interior.Color = Color.Black
                                ExcelSheet.Range("B" & k + 20).Font.Color = Color.White

                                DrawAllBorders(ExcelSheet.Range("B20" & ": AT" & k + 19))

                                'Save ke Local
                                xlApp.DisplayAlerts = False

                                If ls_orderNo <> ls_orderNoSplit Then
                                    ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\Print Label-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm")
                                    pFileName_e2 = "\Print Label-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm"
                                Else
                                    ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\Print Label-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm")
                                    pFileName_e2 = "\Print Label-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm"
                                End If

                                xlApp.Workbooks.Close()
                                xlApp.Quit()
                            End If

                            If ls_orderNo <> ls_orderNoSplit Then
                                If sendEmailPASI_EXPORT("Delivery Confirmation", "Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier & " Split (" & Trim(ls_orderNoSplit) & ")", ls_Supplier, pFilename_e, pFileName_e2, ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                            Else
                                If sendEmailPASI_EXPORT("Delivery Confirmation", "Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier, ls_Supplier, pFilename_e, pFileName_e2, ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                            End If
                        End If
                    ElseIf pFilename_e <> "" Then
                        If ls_orderNo <> ls_orderNoSplit Then
                            If pFileName_e2 = "" Then
                                If sendEmailPASI_EXPORT("Delivery Confirmation", "Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier & " Split (" & Trim(ls_orderNoSplit) & ")", ls_Supplier, pFilename_e, "", ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                            Else
                                If sendEmailPASI_EXPORT("Delivery Confirmation", "Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier & " Split (" & Trim(ls_orderNoSplit) & ")", ls_Supplier, pFilename_e, pFileName_e2, ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                            End If
                        Else
                            If pFileName_e2 = "" Then
                                If sendEmailPASI_EXPORT("Delivery Confirmation", "Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier, ls_Supplier, pFilename_e, "", ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                            Else
                                If sendEmailPASI_EXPORT("Delivery Confirmation", "Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier, ls_Supplier, pFilename_e, pFileName_e2, ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                            End If
                        End If
                    End If
                    '=======================================LABEL===========================================

                    '==================================ORDER CONFIRMATION===================================
                    pFilename_e = "" : pFileName_e2 = ""

                    dsDetailDelivery = BindDataOrderConfirmation(ls_orderNo, ls_orderNoSplit, ls_Aff, ls_Supplier)
                    If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                        Dim k As Long
                        Dim dsAffiliate As New DataSet
                        dsAffiliate = Affiliate(Trim(ls_Aff))

                        Dim dsSupplier As New DataSet
                        dsSupplier = Supplier(Trim(ls_Supplier))

                        Dim status As Boolean
                        status = True

                        If dsDetailDelivery.Tables(0).Rows.Count = 0 Then
                            status = False
                        Else
                            status = True
                        End If

                        If status = True Then
                            NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Template Customer Order Confirmation.xlsx"
                            ls_file = NewFileCopy
                            ExcelBook = xlApp.Workbooks.Open(ls_file)
                            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                            ExcelSheet.Range("A7:BA7").Value = IIf(dsDetailDelivery.Tables(0).Rows(0)("EmergencyCls") = "M", "PASI FINAL APPROVAL PO (MONTHLY)", "PASI FINAL APPROVAL PO (EMERGENCY)")
                            ExcelSheet.Range("H1").Value = "PFAM"
                            ExcelSheet.Range("H2").Value = receiptEmail.Trim
                            ExcelSheet.Range("H3").Value = ls_Consignee.Trim
                            ExcelSheet.Range("H4").Value = ls_delivery.Trim
                            ExcelSheet.Range("H5").Value = ls_Supplier.Trim

                            ExcelSheet.Range("G11:K11").Value = dsDetailDelivery.Tables(0).Rows(0)("period")
                            ExcelSheet.Range("G13:K13").Value = Trim(dsDetailDelivery.Tables(0).Rows(0)("OrderNo"))
                            ExcelSheet.Range("G15:K15").Value = Trim(IIf(dsDetailDelivery.Tables(0).Rows(0)("OrderNo") <> dsDetailDelivery.Tables(0).Rows(0)("poNO"), dsDetailDelivery.Tables(0).Rows(0)("PoNo"), ""))
                            ExcelSheet.Range("G17:K17").Value = IIf(dsDetailDelivery.Tables(0).Rows(0)("ShipBy") = "B", "BOAT", "AIR")

                            ExcelSheet.Range("R11:V11").Value = dsDetailDelivery.Tables(0).Rows(0)("ETDVENDOR")
                            ExcelSheet.Range("R13:V13").Value = dsDetailDelivery.Tables(0).Rows(0)("ETDPORT")
                            ExcelSheet.Range("R15:V15").Value = dsDetailDelivery.Tables(0).Rows(0)("ETAPORT")
                            ExcelSheet.Range("R17:V17").Value = dsDetailDelivery.Tables(0).Rows(0)("ETAFACTORY")

                            ExcelSheet.Range("G19:V19").Value = Trim(dsDetailDelivery.Tables(0).Rows(0)("AFFName"))
                            ExcelSheet.Range("G20:V22").Value = Trim(dsDetailDelivery.Tables(0).Rows(0)("AFFAdd"))

                            ExcelSheet.Range("AE11:AT11").Value = Trim(dsDetailDelivery.Tables(0).Rows(0)("SuppName"))
                            ExcelSheet.Range("AE12:AT15").Value = Trim(dsDetailDelivery.Tables(0).Rows(0)("SuppAdd"))

                            ExcelSheet.Range("AE19:AT19").Value = Trim(dsDetailDelivery.Tables(0).Rows(0)("FWDName"))
                            ExcelSheet.Range("AE20:AT22").Value = Trim(dsDetailDelivery.Tables(0).Rows(0)("FWDAdd"))

                            k = 0
                            For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                                'For i = 0 To 3
                                k = k
                                Dim newKanbanNo As String = ""

                                ExcelSheet.Range("B" & k + 27 & ": C" & k + 27).Merge()
                                ExcelSheet.Range("D" & k + 27 & ": H" & k + 27).Merge()
                                ExcelSheet.Range("i" & k + 27 & ": P" & k + 27).Merge()

                                ExcelSheet.Range("Q" & k + 27 & ": R" & k + 27).Merge()
                                ExcelSheet.Range("S" & k + 27 & ": T" & k + 27).Merge()
                                ExcelSheet.Range("U" & k + 27 & ": X" & k + 27).Merge()

                                ExcelSheet.Range("Y" & k + 27 & ": AB" & k + 27).Merge()
                                ExcelSheet.Range("AC" & k + 27 & ": AF" & k + 27).Merge()
                                ExcelSheet.Range("AG" & k + 27 & ": AJ" & k + 27).Merge()
                                ExcelSheet.Range("AK" & k + 27 & ": AN" & k + 27).Merge()
                                ExcelSheet.Range("AO" & k + 27 & ": AQ" & k + 27).Merge()
                                ExcelSheet.Range("AR" & k + 27 & ": AT" & k + 27).Merge()
                                ExcelSheet.Range("AU" & k + 27 & ": AW" & k + 27).Merge()
                                ExcelSheet.Range("AX" & k + 27 & ": AZ" & k + 27).Merge()
                                ExcelSheet.Range("BA" & k + 27 & ": BC" & k + 27).Merge()

                                ExcelSheet.Range("B" & k + 27 & ": C" & k + 27).Value = k + 1
                                ExcelSheet.Range("D" & k + 27 & ": H" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Partno"))
                                ExcelSheet.Range("i" & k + 27 & ": P" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("PartName")
                                ExcelSheet.Range("Q" & k + 27 & ": R" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("UOM")
                                ExcelSheet.Range("S" & k + 27 & ": T" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("QtyBox")
                                ExcelSheet.Range("S" & k + 27 & ": T" & k + 27).NumberFormat = "#,##0"
                                ExcelSheet.Range("U" & k + 27 & ": X" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("Qty")
                                ExcelSheet.Range("U" & k + 27 & ": X" & k + 27).NumberFormat = "#,##0"
                                ExcelSheet.Range("Y" & k + 27 & ": AB" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("SupplierQty")
                                ExcelSheet.Range("Y" & k + 27 & ": AB" & k + 27).NumberFormat = "#,##0"
                                ExcelSheet.Range("AC" & k + 27 & ": AF" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalBox")
                                ExcelSheet.Range("AC" & k + 27 & ": AF" & k + 27).NumberFormat = "#,##0"
                                ExcelSheet.Range("AG" & k + 27 & ": AJ" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalNet")
                                ExcelSheet.Range("AG" & k + 27 & ": AJ" & k + 27).NumberFormat = "#,##0.00"
                                ExcelSheet.Range("AK" & k + 27 & ": AN" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("Volume")
                                ExcelSheet.Range("AK" & k + 27 & ": AN" & k + 27).NumberFormat = "#,##0.00"
                                ExcelSheet.Range("AO" & k + 27 & ": AQ" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("Length")
                                ExcelSheet.Range("AO" & k + 27 & ": AQ" & k + 27).NumberFormat = "#,##0.00"
                                ExcelSheet.Range("AR" & k + 27 & ": AT" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("Width")
                                ExcelSheet.Range("AR" & k + 27 & ": AT" & k + 27).NumberFormat = "#,##0.00"
                                ExcelSheet.Range("AU" & k + 27 & ": AW" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("Height")
                                ExcelSheet.Range("AU" & k + 27 & ": AW" & k + 27).NumberFormat = "#,##0.00"
                                ExcelSheet.Range("AX" & k + 27 & ": AZ" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("Gross")
                                ExcelSheet.Range("AX" & k + 27 & ": AZ" & k + 27).NumberFormat = "#,##0.00"
                                ExcelSheet.Range("BA" & k + 27 & ": BC" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("BoxPallet")
                                ExcelSheet.Range("BA" & k + 27 & ": BC" & k + 27).NumberFormat = "#,##0"

                                k = k + 1
                            Next

                            DrawAllBorders(ExcelSheet.Range("B27" & ": BC" & k + 26))

                            'Save ke Local
                            xlApp.DisplayAlerts = False

                            If ls_orderNo <> ls_orderNoSplit Then
                                ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\PASI FINAL APPROVAL-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ").xlsx")
                                pFilename_e = "\PASI FINAL APPROVAL-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ").xlsx"
                            Else
                                ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\PASI FINAL APPROVAL-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & ".xlsx")
                                pFilename_e = "\PASI FINAL APPROVAL-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & ".xlsx"
                            End If

                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                        End If

                        If pFilename_e = "" Then GoTo keluar
                        If sendEmailtoSupllierDeliveryConfirmation("DELIVERY", pFilename_e, "", Trim(ls_orderNo), Trim(ls_delivery), Trim(ls_Aff), Trim(ls_orderNoSplit), Trim(ls_Supplier), Trim(ls_SJ)) = False Then GoTo keluar
                    End If
                    '==================================ORDER CONFIRMATION===================================

                    Call UpdateStatusPOExport(ls_Aff, ls_Supplier, ls_orderNo, ls_orderNoSplit)

keluar:
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
            Next
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation to Supplier STOPPED, because " & Err.Description & " " & vbCrLf & _
                                    rtbProcess.Text
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try

        'Exit Sub
        'ErrHandler:
        '        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation to Supplier STOPPED, because " & Err.Description & " " & vbCrLf & _
        '                            rtbProcess.Text
        '        xlApp.Workbooks.Close()
        '        xlApp.Quit()
    End Sub

    Private Function EmailToEmailCCPOMonthly(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                " select 'AFF' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailAffiliate_Export where AffiliateID='" & pAfffCode & "'" & vbCrLf & _
                " union all " & vbCrLf & _
                " --PASI TO -CC " & vbCrLf & _
                " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI_Export where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                " union all " & vbCrLf & _
                " --Supplier TO- CC " & vbCrLf & _
                " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail= '' from ms_emailSupplier_Export where SupplierID='" & Trim(pSupplierID) & "'"
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function Supplier(ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function Affiliate(ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function Forwarder(ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = "SELECT * FROM dbo.MS_Forwarder WHERE ForwarderID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function bindDataDetailMonthly(ByVal pDate As Date, ByVal pAffCode As String, ByVal pPONo As String, ByVal pSupplierID As String, ByVal pOrderNo1 As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = " SELECT Sort, NoUrut, PONo, PartNo, PartName, " & vbCrLf & _
                  " 	UOM, MOQ, QtyBox, BYWHAT, " & vbCrLf & _
                  " 	Week1, Week2, Week3, Week4, Week5, TotalPOQty, " & vbCrLf & _
                  " 	Forecast1, Forecast2, Forecast3, ETDVendor1, ISNULL(SplitReffPONo, '') SplitReffPONo " & vbCrLf & _
                  " FROM ( " & vbCrLf & _
                  " SELECT row_number() over (order by POD.PONo) as Sort, " & vbCrLf & _
                  " 	CONVERT(CHAR,row_number() over (order by POD.PONo)) as NoUrut, " & vbCrLf & _
                  " 	PONo = RTRIM(POD.PONo), PartNo = RTRIM(POD.PartNo), PartName = RTRIM(PartName), " & vbCrLf & _
                  " 	UOM = MU.Description, MOQ = CONVERT(CHAR,ISNULL(POD.POMOQ,MPM.MOQ)), QtyBox = CONVERT(CHAR,ISNULL(POD.POQtyBox,MPM.QtyBox)), " & vbCrLf & _
                  " 	'ORDER' BYWHAT, " & vbCrLf & _
                  " 	ISNULL(Week1,0)Week1, " & vbCrLf

        ls_SQL = ls_SQL + " 	ISNULL(Week2,0)Week2, " & vbCrLf & _
                          " 	ISNULL(Week3,0)Week3, " & vbCrLf & _
                          " 	ISNULL(Week4,0)Week4, " & vbCrLf & _
                          " 	ISNULL(Week5,0)Week5, " & vbCrLf & _
                          " 	ISNULL(TotalPOQty,0)TotalPOQty, " & vbCrLf & _
                          " 	Forecast1 = ISNULL(CONVERT(CHAR,Forecast1),0),  " & vbCrLf & _
                          " 	Forecast2 = ISNULL(CONVERT(CHAR,Forecast2),0),  " & vbCrLf & _
                          " 	Forecast3 = ISNULL(CONVERT(CHAR,Forecast3),0), ETDVendor1, ME.SplitReffPONo " & vbCrLf & _
                          " FROM dbo.PO_Detail_Export POD " & vbCrLf & _
                          " INNER JOIN PO_Master_Export ME ON ME.PONo = POD.PONo AND ME.AffiliateID = POD.AffiliateID AND ME.SupplierID = POD.SupplierID AND ME.PONo = POD.PONo AND POD.OrderNo1 = ME.OrderNo1" & vbCrLf & _
                          " LEFT JOIN dbo.MS_Parts MPART ON POD.PartNo = MPART.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.AffiliateID = POD.AffiliateID and MPM.SupplierID = POD.SupplierID and MPM.PartNo = POD.PartNo " & vbCrLf & _
                          " LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls " & vbCrLf

        ls_SQL = ls_SQL + " WHERE EmergencyCls = 'M' AND POD.PONo = '" & Trim(pPONo) & "' AND POD.OrderNo1 = '" & Trim(pOrderNo1) & "' AND POD.AffiliateID = '" & Trim(pAffCode) & "' AND POD.SupplierID = '" & Trim(pSupplierID) & "'" & vbCrLf & _
                          " GROUP BY POD.PONo, POD.PartNo, PartName, MU.Description, ISNULL(POD.POMOQ,MPM.MOQ), ISNULL(POD.POQtyBox,MPM.QtyBox), " & vbCrLf & _
                          " 	Week1, Week2, Week3, Week4, Week5, TotalPOQty, " & vbCrLf & _
                          " 	Forecast1, Forecast2, Forecast3, ETDVendor1, ME.SplitReffPONo )detail1 " & vbCrLf & _
                          "  "

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Private Function sendEmailPOtoSupllierMonthly(ByVal pFileName1 As String, ByVal pFileName2 As String, ByVal pPONo As String, ByVal pOrderNo1 As String) As Boolean
        Try
            Dim TempFilePath As String
            Dim TempFileName1 As String = ""
            Dim TempFileName2 As String = ""
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)
            TempFileName1 = "\" & pFileName1
            If pFileName2 <> "" Then TempFileName2 = "\" & pFileName2

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCCPOMonthly(pAffCode, "PASI", pSupplier)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPOtoSupllierMonthly = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPOtoSupllierMonthly = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            If Trim(pPONo) = Trim(pOrderNo1) Then
                mailMessage.Subject = "Send To Supplier Order No : " & pPONo.Trim & "-" & pSupplier.Trim & ""
            Else
                mailMessage.Subject = "Send To Supplier Order No : " & pPONo.Trim & "-" & pSupplier.Trim & " Split (" & pOrderNo1.Trim & ")"
            End If

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            'Dim receiptBCCEmail As String
            'If pSupplier.Trim.ToUpper = "KMK" Then
            '    receiptBCCEmail = "Ariefrosyadi91@gmail.com;Ristian_a@yahoo.com;pasi.purchase@gmail.com"
            'Else
            '    receiptBCCEmail = "pasi.purchase@gmail.com"
            'End If
            'If receiptBCCEmail <> "" Then
            '    For Each recipientBCC In receiptBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If

            GetSettingEmail_Export("PO")
            'uf_GetNotification("11")
            'ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
            'mailMessage.Body = ls_Body
            'ls_Body = clsNotification.GetNotification("16", "", pPONo.Trim)
            If Trim(pPONo) = Trim(pOrderNo1) Then
                ls_Body = clsNotification.GetNotification("16", "", pOrderNo1.Trim)
            Else
                ls_Body = clsNotification.GetNotification("16", "", pPONo.Trim & " Split (" & pOrderNo1.Trim & ")")
            End If
            mailMessage.Body = ls_Body

            Dim filename1 As String = TempFilePath & TempFileName1
            mailMessage.Attachments.Add(New Attachment(filename1))

            If TempFileName2 <> "" Then
                Dim filename2 As String = TempFilePath & TempFileName2
                mailMessage.Attachments.Add(New Attachment(filename2))
            End If

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            'smtp.Host = "smtp.atisicloud.com"
            'smtp.Host = "mail.fast.net.id"

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
            sendEmailPOtoSupllierMonthly = True
            'Delete the file
            'Kill(TempFilePath & TempFileName)
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO No: " & pPONo & " to Supplier SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
            Exit Function
        Catch ex As Exception
            sendEmailPOtoSupllierMonthly = False
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO No: " & pPONo & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text

        End Try

    End Function

    Private Sub GetSettingEmail_Export(ByVal ls_Value As String)
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = "SELECT * FROM dbo.Ms_EmailSetting_Export"
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            smtpClient = Trim(ds.Tables(0).Rows(0)("SMTP"))
            portClient = Trim(ds.Tables(0).Rows(0)("PORTSMTP"))
            usernameSMTP = If(IsDBNull(ds.Tables(0).Rows(0)("usernameSMTP")), "", ds.Tables(0).Rows(0)("usernameSMTP"))
            PasswordSMTP = If(IsDBNull(ds.Tables(0).Rows(0)("passwordSMTP")), "", ds.Tables(0).Rows(0)("passwordSMTP"))
            DefaultCredentials = IIf(ds.Tables(0).Rows(0)("DefaultCredentials") = "1", True, False)
            SSL = IIf(ds.Tables(0).Rows(0)("SSL") = "1", True, False)
        Else
            rtbProcess.Text = rtbProcess.Text & vbCrLf & _
                         Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send " & ls_Value & " to Supplier STOPPED, because Email Setting Empty "
        End If
    End Sub

    Private Sub UpdateExcelPOMonthly(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffCode As String = "", _
                         Optional ByVal pPONo As String = "", _
                         Optional ByVal pSuppCode As String = "", _
                         Optional ByVal pOrderNo1 As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.PO_Master_Export " & vbCrLf & _
                      " SET ExcelCls='2'" & vbCrLf & _
                      " WHERE PONo='" & pPONo & "'  " & vbCrLf & _
                      " AND OrderNo1='" & pOrderNo1 & "' " & vbCrLf & _
                      " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                      " AND SupplierID='" & pSuppCode & "' " & vbCrLf & _
                      " AND EmergencyCls = 'M' "
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
        End Try
    End Sub

    Private Function EmailToEmailCCPOEmergency(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                " select 'AFF' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailAffiliate_Export where AffiliateID='" & pAfffCode & "'" & vbCrLf & _
                " union all " & vbCrLf & _
                " --PASI TO -CC " & vbCrLf & _
                " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI_Export where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                " union all " & vbCrLf & _
                " --Supplier TO- CC " & vbCrLf & _
                " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail='' from ms_emailSupplier_Export where SupplierID='" & Trim(pSupplierID) & "'"
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function bindDataDetailEmergency(ByVal pDate As Date, ByVal pAffCode As String, ByVal pPONo As String, ByVal pSupplierID As String, ByVal pOrderNo1 As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = "  SELECT row_number() over (order by POD.PONo) as Sort,  " & vbCrLf & _
                  "  	CONVERT(CHAR,row_number() over (order by POD.PONo)) as NoUrut,  " & vbCrLf & _
                  "  	PONo = RTRIM(POD.PONo), PartNo = RTRIM(POD.PartNo), PartName = RTRIM(PartName),  " & vbCrLf & _
                  "  	UOM = MU.Description, MOQ = CONVERT(CHAR,ISNULL(POD.POMOQ,MPM.MOQ)), QtyBox = CONVERT(CHAR,ISNULL(POD.POQtyBox,MPM.QtyBox)),  " & vbCrLf & _
                  "  	'ORDER' BYWHAT,  " & vbCrLf & _
                  "  	ISNULL(Week1,0)Week1, ETDVendor1 " & vbCrLf & _
                  "  FROM dbo.PO_Detail_Export POD  " & vbCrLf & _
                  "  INNER JOIN PO_Master_Export ME ON ME.PONo = POD.PONo AND ME.AffiliateID = POD.AffiliateID AND ME.SupplierID = POD.SupplierID  AND ME.OrderNo1 = POD.OrderNo1 " & vbCrLf & _
                  "  LEFT JOIN dbo.MS_Parts MPART ON POD.PartNo = MPART.PartNo  " & vbCrLf & _
                  "  LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls  " & vbCrLf & _
                  " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo and MPM.AffiliateID = POD.AffiliateID and MPM.SupplierID = POD.SupplierID " & vbCrLf

        ls_SQL = ls_SQL + " WHERE EmergencyCls = 'E' AND POD.PONo = '" & Trim(pPONo) & "' AND POD.AffiliateID = '" & Trim(pAffCode) & "' AND POD.SupplierID = '" & Trim(pSupplierID) & "' AND POD.OrderNo1 = '" & Trim(pOrderNo1) & "'" & vbCrLf

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Private Function sendEmailPOtoSupllierEmergency(ByVal pFileName1 As String, ByVal pFileName2 As String, ByVal pOrderNo1 As String) As Boolean
        Try
            Dim TempFilePath As String
            'Dim TempFileName As String
            Dim TempFileName1 As String = ""
            Dim TempFileName2 As String = ""
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)
            'TempFileName = "\" & pFileName
            TempFileName1 = "\" & pFileName1
            If pFileName2 <> "" Then TempFileName2 = "\" & pFileName2

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCCPOEmergency(pAffCode, "PASI", pSupplier)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                sendEmailPOtoSupllierEmergency = False
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                Exit Function
            End If
            If fromEmail = "" Then
                sendEmailPOtoSupllierEmergency = False
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            'mailMessage.Subject = "Send To Supplier Order No: " & pOrderNo1 & " "
            If Trim(pPONo) = Trim(pOrderNo1) Then
                mailMessage.Subject = "Send To Supplier Order No : " & pPONo.Trim & "-" & pSupplier.Trim & ""
            Else
                mailMessage.Subject = "Send To Supplier Order No : " & pPONo.Trim & "-" & pSupplier.Trim & " Split (" & pOrderNo1.Trim & ")"
            End If

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If
            'Dim receiptBCCEmail As String
            'If pSupplier.Trim.ToUpper = "KMK" Then
            '    receiptBCCEmail = "Ariefrosyadi91@gmail.com;Ristian_a@yahoo.com;pasi.purchase@gmail.com"
            'Else
            '    receiptBCCEmail = "pasi.purchase@gmail.com"
            'End If
            'If receiptBCCEmail <> "" Then
            '    For Each recipientBCC In receiptBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If
            GetSettingEmail_Export("PO")
            ''uf_GetNotification("11")
            ''ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
            ''mailMessage.Body = ls_Body
            'ls_Body = clsNotification.GetNotification("11", "", pPONo.Trim)
            If Trim(pPONo) = Trim(pOrderNo1) Then
                ls_Body = clsNotification.GetNotification("11", "", pOrderNo1.Trim)
            Else
                ls_Body = clsNotification.GetNotification("11", "", pPONo.Trim & " Split (" & pOrderNo1.Trim & ")")
            End If
            mailMessage.Body = ls_Body

            'Dim filename As String = TempFilePath & TempFileName
            'mailMessage.Attachments.Add(New Attachment(filename))
            Dim filename1 As String = TempFilePath & TempFileName1
            mailMessage.Attachments.Add(New Attachment(filename1))

            If TempFileName2 <> "" Then
                Dim filename2 As String = TempFilePath & TempFileName2
                mailMessage.Attachments.Add(New Attachment(filename2))
            End If

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            'smtp.Host = "smtp.atisicloud.com"
            'smtp.Host = "mail.fast.net.id"

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
            'Delete the file
            'Kill(TempFilePath & TempFileName)
            sendEmailPOtoSupllierEmergency = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO No: " & pPONo & " to Supplier SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
        Catch ex As Exception
            sendEmailPOtoSupllierEmergency = False
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO No: " & pPONo & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text

        End Try

    End Function

    Private Sub UpdateExcelPOEmergency(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffCode As String = "", _
                         Optional ByVal pPONo As String = "", _
                         Optional ByVal pOrderNo As String = "", _
                         Optional ByVal pSuppCode As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.PO_Master_Export " & vbCrLf & _
                      " SET ExcelCls='2'" & vbCrLf & _
                      " WHERE PONo='" & pPONo & "'  " & vbCrLf & _
                      " AND OrderNo1='" & pOrderNo & "'  " & vbCrLf & _
                      " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                      " AND SupplierID='" & pSuppCode & "' " & vbCrLf & _
                      " AND EmergencyCls = 'E' "
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
        End Try
    End Sub

    Private Sub InsertPrintLabel(ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pTahun As Integer, ByVal pBulan As Integer)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"
        Dim i As Integer = 0
        Dim x As Integer = 0
        Dim dsData As New DataSet
        Dim dsSeqNo As New DataSet
        Dim dsAda As New DataSet
        Dim ls_Startno As Integer
        Dim LabelNo As String = ""
        Dim ls_awalT As Integer = 2016
        Dim ls_selisihT As Integer = 0
        Dim ls_charT As Integer = 65
        Dim ls_codeT As String = ""
        Dim ls_awalB As Integer = 1
        Dim ls_selisihB As Integer = 0
        Dim ls_charB As Integer = 65
        Dim ls_codeB As String = ""
        Dim ls_Period As String = ""

        Try
            'If ls_newLabelEx = False Then
            dsAda = SelectInsertBarcodeExport(pPono, pOrderNo, pAff, pSupp)
            'Else
            '    dsAda = SelectInsertBarcodeExportNew(pPono, pOrderNo, pAff, pSupp)
            'End If

            If dsAda.Tables(0).Rows.Count = 0 Then
                dsData = InsertBarcodeExport(pPono, pOrderNo, pAff, pSupp)
                If dsData.Tables(0).Rows.Count > 0 Then
                    ls_Period = Format(dsData.Tables(0).Rows(0)("Period"), "yyyy-MM-dd")
                    Using sqlConn As New SqlConnection(cfg.ConnectionString)
                        sqlConn.Open()

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CreateKanban")
                            Dim sqlCommNew As SqlCommand = sqlConn.CreateCommand
                            sqlCommNew.Connection = sqlConn
                            sqlCommNew.Transaction = sqlTran

                            '============== Cari code =============
                            If ls_newLabelEx = True Then
                                'TAHUN
                                ls_selisihT = pTahun - ls_awalT
                                If ls_selisihT = 0 Then
                                    ls_codeT = Chr(ls_charT)
                                Else
                                    If (ls_charT + ls_selisihT) >= 73 Then 'I
                                        ls_selisihT = ls_selisihT + 1

                                        If ls_selisihT >= 79 Then 'O
                                            ls_selisihT = ls_selisihT + 1

                                            If ls_selisihT >= 83 Then 'S
                                                ls_selisihT = ls_selisihT + 1
                                            End If
                                        End If
                                    End If
                                    ls_codeT = Chr(ls_charT + ls_selisihT)
                                End If
                                'BULAN
                                ls_selisihB = pBulan - ls_awalB
                                If ls_selisihB = 0 Then
                                    ls_codeB = Chr(ls_charB)
                                Else
                                    If (ls_charB + ls_selisihB) >= 73 Then 'I
                                        ls_selisihB = ls_selisihB + 1

                                        If ls_selisihB >= 79 Then 'O
                                            ls_selisihB = ls_selisihB + 1

                                            If ls_selisihB >= 83 Then 'S
                                                ls_selisihB = ls_selisihB + 1
                                            End If
                                        End If
                                    End If
                                    ls_codeB = Chr(ls_charB + ls_selisihB)
                                End If
                            End If
                            '============== Cari code =============

                            For i = 0 To dsData.Tables(0).Rows.Count - 1
                                If ls_newLabelEx = False Then
                                    dsSeqNo = GetLABELNO(pPono, pOrderNo, pAff, pSupp, dsData.Tables(0).Rows(i)("PartNo"))
                                Else
                                    dsSeqNo = GetLABELNONew(ls_codeT + ls_codeB, ls_Period)
                                End If
                                If dsSeqNo.Tables(0).Rows.Count > 0 Then
                                    ls_Startno = dsSeqNo.Tables(0).Rows(0)("seqno")
                                Else
                                    ls_Startno = 0
                                End If
                                For x = 1 To dsData.Tables(0).Rows(i)("looping")
                                    If ls_codeT + ls_codeB = "HC" Then
                                        Dim a As String = "a"
                                    End If
                                    '------ NEW LABEL ------
                                    If ls_newLabelEx = False Then
                                        LabelNo = "00000" & ls_Startno + 1
                                        LabelNo = Trim(dsData.Tables(0).Rows(i)("LabelCode")) + Microsoft.VisualBasic.Right(LabelNo, 5)
                                    Else
                                        LabelNo = "0000000" & ls_Startno + 1
                                        LabelNo = ls_codeT + ls_codeB + Microsoft.VisualBasic.Right(LabelNo, 7)
                                    End If
                                    '------ NEW LABEL ------

                                    ls_SQL = "INSERT INTO PrintLabelExport " & vbCrLf & _
                                             " VALUES ( " & vbCrLf & _
                                             " '" & pPono & "', " & vbCrLf & _
                                             " '" & pAff & "', " & vbCrLf & _
                                             " '" & pSupp & "', " & vbCrLf & _
                                             " '" & dsData.Tables(0).Rows(i)("PartNo") & "', " & vbCrLf & _
                                             " '" & LabelNo & "', " & vbCrLf & _
                                             " getdate(), " & vbCrLf & _
                                             " '" & UserName & "', " & vbCrLf & _
                                             " '" & pOrderNo & "','','','' )"
                                    sqlCommNew = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                    sqlCommNew.ExecuteNonQuery()
                                    ls_Startno = ls_Startno + 1
                                Next
                            Next

                            sqlCommNew.Dispose()
                            sqlTran.Commit()
                        End Using
                        sqlConn.Close()
                    End Using
                End If
            End If
        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub

    Public Function SelectInsertBarcodeExport(ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String)
        Dim ls_sql As String = ""
        'MdlConn.ReadConnection()

        ls_sql = " select * from PrintLabelExport" & vbCrLf & _
                 " where PONO = '" & pPono & "'" & vbCrLf & _
                 " and AffiliateID = '" & pAff & "' and SupplierID = '" & pSupp & "' " & vbCrLf & _
                 " and OrderNo = '" & pOrderNo & "'"

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)

        Return ds

    End Function

    Public Function InsertBarcodeExport(ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String)
        Dim ls_sql As String = ""
        'MdlConn.ReadConnection()

        ls_sql = " Select distinct POM.Period, POD.PoNo,POD.OrderNo1, POD.AffiliateID, POD.SupplierID, POD.PartNo as PartNo, " & vbCrLf & _
                 " PUD.Week1 as qty, QtyBox = ISNULL(POD.POQtyBox,MPM.QtyBox), looping = CASE WHEN ISNULL(POD.POQtyBox,ISNULL(MPM.QtyBox,0)) = 0 THEN 0 ELSE Convert(numeric,PUD.Week1/ISNULL(POD.POQtyBox,MPM.QtyBox)) END, MSP.LabelCode " & vbCrLf & _
                 " From PO_Detail_Export POD INNER JOIN PO_DetailUpload_Export PUD with(nolock) " & vbCrLf & _
                 " ON POD.PONo = PUD.PONo " & vbCrLf & _
                 " AND POD.OrderNo1 = PUD.OrderNo1 " & vbCrLf & _
                 " AND POD.AffiliateID = PUD.AffiliateID " & vbCrLf & _
                 " AND POD.SupplierID = PUD.SupplierID " & vbCrLf & _
                 " AND POD.partNO = PUD.PartNo " & vbCrLf & _
                 " INNER JOIN MS_Parts MP ON MP.PartNo = POD.partNo " & vbCrLf & _
                 " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo " & vbCrLf & _
                 " AND MPM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                 " AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
                 " LEFT JOIN MS_Supplier MSP ON MSP.SupplierID = POD.SupplierID " & vbCrLf & _
                 " LEFT JOIN PO_Master_Export POM ON POM.PONo = POD.PONo and POM.OrderNo1 = POD.OrderNo1 and POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                 " AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                 " where POD.PONO = '" & pPono & "' and POD.OrderNO1 = '" & pOrderNo & "' " & vbCrLf & _
                 " and POD.AffiliateID = '" & pAff & "' and POD.SupplierID = '" & pSupp & "' " & vbCrLf

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)

        Return ds

    End Function

    Public Function GetLABELNO(ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pPartNo As String)
        Dim ls_sql As String = ""
        'MdlConn.ReadConnection()

        ls_sql = " select seqno = Convert(numeric,replace(max(labelno), left(max(labelno),1),'')) from PrintLabelExport with(nolock) " & vbCrLf & _
                 " where SupplierID = '" & pSupp & "' " & vbCrLf & _
                 " and PartNo = '" & Trim(pPartNo) & "'" & vbCrLf & _
                 " GROUP BY SupplierID, PartNo "

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds

    End Function

    Public Function GetLABELNONew(ByVal pcode As String, ByVal pperiod As String)
        Dim ls_sql As String = ""
        'MdlConn.ReadConnection()

        ls_sql = " select seqno = isnull(Convert(numeric,replace(max(labelno), left(max(labelno),2),'')),0) " & vbCrLf & _
                 " from PrintLabelExport PL with(nolock) " & vbCrLf & _
                 " --LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                 " --ON POM.PONo = PL.PONo and POM.OrderNo1 = PL.OrderNo and POM.AffiliateID = PL.AffiliateID and POM.SupplierID = PL.SupplierID " & vbCrLf & _
                 " where Left(LabelNo,2) = '" & pcode & "' " & vbCrLf & _
                 " --and POM.Period = '" & pperiod & "' "

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds

    End Function

    Private Function BidDataDeliveryConfirm(ByVal PONO As String, ByVal OrderNo As String, ByVal affiliateID As String, ByVal SupplierID As String)
        Dim ls_sql As String

        ls_sql = " select distinct " & vbCrLf & _
                " Partno = POD.PartNo, " & vbCrLf & _
                " PartName = MP.Partname, " & vbCrLf & _
                " UOM = MUC.Description, " & vbCrLf & _
                " MOQ = ISNULL(POD.POQtyBox,MPM.QtyBox), " & vbCrLf & _
                " OrderQty = POD.Week1,  " & vbCrLf & _
                " labelno = Rtrim(PL.Label1) + ' - ' + Rtrim(pl.Label2), " & vbCrLf & _
                " LabelNo1 = Rtrim(PL.Label1), LabelNo2 = Rtrim(PL.Label2), " & vbCrLf & _
                " SuppQty = POD.week1, TotalPOQty = POD.Week1 / ISNULL(POD.POQtyBox,MPM.QtyBox) " & vbCrLf & _
                " From PO_Detail_Export POD --LEFT JOIN PO_DetailUpload_Export PUD " & vbCrLf & _
                " --ON POD.PONo = PUD.PONo " & vbCrLf & _
                " --And POD.AffiliateID = PUD.AffiliateID " & vbCrLf & _
                " --And POD.SupplierID = PUD.SupplierID " & vbCrLf

        ls_sql = ls_sql + " --And POD.PartNo = PUD.PartNo " & vbCrLf & _
                " --And POD.OrderNo1 = PUD.OrderNo1 " & vbCrLf & _
                " --AND POD.ForwarderID = PUD.ForwarderID " & vbCrLf & _
                " LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo " & vbCrLf & _
                " AND MPM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                " AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
                " INNER JOIN (select pono, orderNo, affiliateID, SupplierID, PartNo," & vbCrLf & _
                "             Min(labelNo) as label1, Max(labelNo) as label2 from PrintLabelExport " & vbCrLf & _
                "             Group by pono, orderNo, affiliateID, SupplierID, PartNo) PL ON PL.PONo = POD.PONo   " & vbCrLf & _
                "         and PL.AffiliateID = POD.AffiliateID  AND PL.SupplierID = POD.SupplierID" & vbCrLf & _
                "         AND PL.PartNO = POD.PartNo  AND PL.orderNo = POD.OrderNo1" & vbCrLf & _
                " LEFT JOIN MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls" & vbCrLf & _
                " where POD.PONO = '" & Trim(PONO) & "' " & vbCrLf & _
                " AND POD.AffiliateID = '" & Trim(affiliateID) & "'" & vbCrLf & _
                " AND POD.SupplierID = '" & Trim(SupplierID) & "' " & vbCrLf & _
                " AND POD.OrderNo1 = '" & Trim(OrderNo) & "' "

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Private Function EmailToEmailCCKanban_Export(ByVal pAfffCode As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = " select 'AFF' as FLAG , kanbanCC =  Rtrim(kanbanCC) + ';' + Rtrim(KanbanTo),kanbanTo = '' from MS_emailAffiliate_Export where AffiliateID = '" & pAfffCode & "' " & vbCrLf & _
                     " union ALL " & vbCrLf & _
                     " select 'PASI' as FLAG , kanbanCC = AffiliatePOCC, kanbanTo = AffiliatePOTo from MS_EmailPasi_Export " & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     " select 'SUPP' as FLAG , kanbanCC, kanbanTo = '' from MS_EmailSupplier_Export where supplierID = '" & pSupplierID & "' "
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function BindDataLabelPrint(ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String)
        Dim ls_sql As String
        'MdlConn.ReadConnection()

        ls_sql = " select " & vbCrLf & _
                  " Partno = POD.PartNo,   " & vbCrLf & _
                  " Partname = MP.PartName,   " & vbCrLf & _
                  " labelno = Rtrim(PL.Label1) + ' - ' + Rtrim(pl.Label2),  " & vbCrLf & _
                  " Label1 = Rtrim(label1), " & vbCrLf & _
                  " Label2 = Rtrim(label2), " & vbCrLf & _
                  " orderno = POD.PONo,   " & vbCrLf & _
                  " uom = MU.Description,   " & vbCrLf & _
                  " qtybox = ISNULL(POD.POQtyBox,MPM.Qtybox),   " & vbCrLf & _
                  " qty = convert(char,PUD.Week1),   " & vbCrLf & _
                  " boxqty = Ceiling(PUD.week1 / ISNULL(POD.POQtyBox,MPM.Qtybox)),   " & vbCrLf

        ls_sql = ls_sql + " DestinationPort = isnull(MA.DestinationPort,''), " & vbCrLf & _
                          " DestinationPoint = isnull(DeliveryPoint,''), " & vbCrLf & _
                          " CustName = POD.AffiliateID, " & vbCrLf & _
                          " CustCode = AffiliateCode, " & vbCrLf & _
                          " ConsigneeCode = isnull(MA.ConsigneeCode,'') " & vbCrLf & _
                          " From PO_Detail_Export POD INNER JOIN PO_DetailUpload_Export PUD  " & vbCrLf & _
                          " ON POD.PONo = PUD.PONo   " & vbCrLf & _
                          " AND POD.AffiliateID = PUD.AffiliateID   " & vbCrLf & _
                          " AND POD.SupplierID = PUD.SupplierID   " & vbCrLf & _
                          " AND POD.partNO = PUD.PartNo   " & vbCrLf & _
                          " AND pod.fORWARDERid = PUD.ForwarderID " & vbCrLf & _
        " AND POD.OrderNo1 = PUD.OrderNo1 " & vbCrLf & _
                          " INNER JOIN PO_Master_Export POM   " & vbCrLf & _
                          " ON POM.Pono = POD.PONo   " & vbCrLf

        ls_sql = ls_sql + " and POM.AffiliateID = POD.AffiliateID   " & vbCrLf & _
                          " And POM.SupplierID = POD.SupplierID   " & vbCrLf & _
        " AND POM.OrderNo1 = POD.OrderNo1 " & vbCrLf & _
                          " INNER JOIN (select pono, orderNo, affiliateID, SupplierID, PartNo,  " & vbCrLf & _
                          " 			Min(labelNo) as label1, Max(labelNo) as label2 from PrintLabelExport " & vbCrLf & _
                          " 			Group by pono, orderNo, affiliateID, SupplierID, PartNo) PL ON PL.PONo = POD.PONo   " & vbCrLf & _
                          " and PL.AffiliateID = POD.AffiliateID  AND PL.SupplierID = POD.SupplierID   " & vbCrLf & _
                          " AND PL.PartNO = POD.PartNo   " & vbCrLf & _
                          " AND PL.OrderNo = POM.OrderNo1 " & vbCrLf & _
                          " INNER JOIN MS_Parts MP ON MP.PartNo = POD.partNo   " & vbCrLf & _
                          " LEFT JOIN MS_PARTMApping MPM ON MPM.PartNo = PL.PartNo and MPM.AffiliateID = PL.AffiliateID and MPM.SupplierID = PL.SupplierID  " & vbCrLf & _
                          " INNER JOIN ms_unitcls MU ON MU.UnitCls = MP.Unitcls   " & vbCrLf & _
                          " LEFT JOIN ms_affiliate MA ON MA.AffiliateID = POD.AffiliateID " & vbCrLf

        ls_sql = ls_sql + " LEFT JOIN ms_Forwarder MF ON MF.ForwarderID = POD.ForwarderID " & vbCrLf & _
                          " LEFT JOIN MS_Supplier MS ON MS.SupplierID = POD.SupplierID " & vbCrLf & _
                          " Where PL.POno = '" & Trim(pPono) & "'" & vbCrLf & _
                          " AND PL.OrderNo = '" & Trim(pOrderNo) & "' " & vbCrLf & _
                          " AND PL.AffiliateID = '" & Trim(pAff) & "' " & vbCrLf & _
                          " AND PL.SupplierID = '" & Trim(pSupp) & "'"

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Private Function sendEmailPASI_EXPORT(ByVal pStatus As String, ByVal pCaption As String, ByVal pSupplier1 As String, ByVal pAtt1 As String, ByVal pAtt2 As String, ByVal pPono As String, ByVal pOrderNo As String) As Boolean 'Link Affiliate Order Entry
        Try
            Dim TempFilePath As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCCPO_Export("", "PASI", pSupplier1)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPASI_EXPORT = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPASI_EXPORT = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "Send " & pCaption

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            'Dim receiptBCCEmail As String
            'If pSupplier1.Trim.ToUpper = "KMK" Then
            '    receiptBCCEmail = "Ariefrosyadi91@gmail.com;Ristian_a@yahoo.com;pasi.purchase@gmail.com"
            'Else
            '    receiptBCCEmail = "pasi.purchase@gmail.com"
            'End If
            'If receiptBCCEmail <> "" Then
            '    For Each recipientBCC In receiptBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If
            GetSettingEmail_Export("Kanban")
            'uf_GetNotification("11")
            'ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
            'mailMessage.Body = ls_Body
            If pPono.Trim = pOrderNo.Trim Then
                ls_Body = clsNotification.GetNotification("18", "", pPono.Trim)
            Else
                ls_Body = clsNotification.GetNotification("18", "", pPono.Trim & " Split (" & pOrderNo.Trim & ")")
            End If
            mailMessage.Body = ls_Body
            'mailMessage.Body = "Dear Sir/Madam,  This is notification for: " & pCaption
            Dim filename As String = TempFilePath & pAtt1
            If pAtt1 <> "" Then mailMessage.Attachments.Add(New Attachment(filename))
            Dim filename2 As String = TempFilePath & pAtt2
            If pAtt2 <> "" Then mailMessage.Attachments.Add(New Attachment(filename2))
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            'smtp.Host = "smtp.atisicloud.com"
            'smtp.Host = "mail.fast.net.id"

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
            'Delete the file
            'Kill(TempFilePath & TempFileName)
            sendEmailPASI_EXPORT = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send DN No: " & pPono & " to Supplier SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send DN No: " & pPono & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
            sendEmailPASI_EXPORT = False
        End Try
    End Function

    Private Function sendEmailPASI_EXPORTForwarder_Information(ByVal pSuratJalan As String, ByVal pAffiliate As String, ByVal pForwarder As String) As Boolean
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim dsEmail As New DataSet
            dsEmail = EmailSendForwarder_Export(pForwarder, pAffiliate)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    If fromEmail = "" Then
                        fromEmail = dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                    Else
                        fromEmail = fromEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") <> "PASI" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") <> "PASI" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                    End If
                End If
            Next
            fromEmail = Replace(fromEmail, " ", "")
            receiptEmail = Replace(receiptEmail, ",", ";")
            receiptCCEmail = Replace(receiptCCEmail, ",", ";")

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Information Forwarder Delivery Confirmation STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPASI_EXPORTForwarder_Information = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Information Forwarder Delivery Confirmation STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPASI_EXPORTForwarder_Information = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "Informasi Perubahan Untuk DN " & pSuratJalan

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            GetSettingEmail_Export("Kanban")

            ls_Body = clsNotification.GetNotification("34", "", , "", pSuratJalan)

            mailMessage.Body = ls_Body

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient

            smtp.Host = smtpClient
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            smtp.Port = portClient
            smtp.Send(mailMessage)

            sendEmailPASI_EXPORTForwarder_Information = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Information Forwarder Delivery Confirmation DN No: " & pSuratJalan & " SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Information Forwarder Delivery Confirmation DN No: " & pSuratJalan & " STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
            sendEmailPASI_EXPORTForwarder_Information = False
        End Try
    End Function

    Private Function EmailToEmailCCPO_Export(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                " select 'AFF' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailAffiliate_Export where AffiliateID='" & pAfffCode & "'" & vbCrLf & _
                " union all " & vbCrLf & _
                " --PASI TO -CC " & vbCrLf & _
                " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI_Export where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                " union all " & vbCrLf & _
                " --Supplier TO- CC " & vbCrLf & _
                " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail='' from ms_emailSupplier_Export where SupplierID='" & Trim(pSupplierID) & "'"
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function BindDataOrderConfirmation(ByVal PONO As String, ByVal OrderNo As String, ByVal affiliateID As String, ByVal SupplierID As String)
        Dim ls_sql As String
        'MdlConn.ReadConnection()

        ls_sql = " SELECT DISTINCT FWDID = isnull(POM.ForwarderID,''), SUPPID = isnull(POM.SupplierID,''), ConsigneeCode = isnull(MA.ConsigneeCode,''), period = CONVERT(char(7),POM.PERIOD), " & vbCrLf & _
                  " EmergencyCls, " & vbCrLf & _
                  " orderNO = POM.OrderNo1, " & vbCrLf & _
                  " PONo = POM.PONo, " & vbCrLf & _
                  " shipby = ISNULL(POM.ShipCls,''), " & vbCrLf & _
                  " AffCode = ISNULL(MA.ConsigneeCode,''), " & vbCrLf & _
                  " AffName = ISNULL(MA.AffiliateName,''), " & vbCrLf & _
                  " AffAdd = ISNULL(MA.ConsigneeAddress,''), " & vbCrLf & _
                  " SuppName = ISNULL(MS.SupplierName,''), " & vbCrLf & _
                  " SuppAdd = ISNULL(MS.ADDRESS,''), " & vbCrLf & _
                  " FwdName = ISNULL(MF.ForwarderName,''), " & vbCrLf & _
                  " FwdAdd = ISNULL(MF.Address,''), " & vbCrLf & _
                  " ETDVendor = convert(char(10),convert(datetime, ETDVendor1),120), " & vbCrLf & _
                  " ETDPort = convert(char(10),convert(datetime, ETDPort1),120), " & vbCrLf & _
                  " ETAPort = convert(char(10),convert(datetime, ETAPort1),120), " & vbCrLf & _
                  " ETAFactory = convert(char(10),convert(datetime, ETAFactory1),120), " & vbCrLf

        ls_sql = ls_sql + " PartNo = POD.PartNo, " & vbCrLf & _
                          " PartName = MP.PartName, " & vbCrLf & _
                          " UOM = ISNULL(MU.DESCRIPTION,''), " & vbCrLf & _
                          " QtyBox = ISNULL(POD.POQtyBox,MPM.Qtybox), " & vbCrLf & _
                          " Qty = POD.Week1, " & vbCrLf & _
                          " SupplierQty = POD.Week1, " & vbCrLf & _
                          " TotalBox = POD.Week1 / ISNULL(POD.POQtyBox,MPM.Qtybox), " & vbCrLf & _
                          " TotalNet = round((POD.Week1 / ISNULL(POD.POQtyBox,MPM.Qtybox)) * (NetWeight/1000),2), " & vbCrLf & _
                          " Volume = round(((Length/1000) * (Height/1000) * (Width/ 1000)) * (POD.Week1/ISNULL(POD.POQtyBox,MPM.Qtybox)),2) " & vbCrLf & _
                          " ,Length = round(((Length))/1000,2) " & vbCrLf & _
                          " ,Width = round(((Width))/1000,2) " & vbCrLf & _
                          " ,Height = round(((Height))/1000,2) " & vbCrLf & _
                          " ,Gross = round(((GrossWeight))/1000,2) " & vbCrLf & _
                          " , BoxPallet = isnull(BoxPallet,0) " & vbCrLf & _
                          " FROM PO_MASTER_EXPORT POM LEFT JOIN PO_DETAIL_EXPORT POD " & vbCrLf & _
                          " ON POM.PONO = POD.PONO " & vbCrLf

        ls_sql = ls_sql + " AND POM.ORDERNO1 = POD.ORDERNO1 " & vbCrLf & _
                          " AND POM.AFFILIATEID = POD.AFFILIATEID " & vbCrLf & _
                          " AND POM.SUPPLIERID = POD.SUPPLIERID " & vbCrLf & _
                          " --LEFT JOIN PO_DetailUpload_Export PUD " & vbCrLf & _
                          " --ON POD.AffiliateID = PUD.AffiliateID " & vbCrLf & _
                          " --And POD.SupplierID = PUD.SupplierID " & vbCrLf & _
                          " --And POD.PartNo = PUD.PartNo " & vbCrLf & _
                          " --AND POD.PONO = PUD.PONO	 " & vbCrLf & _
                          " --AND POD.OrderNo1 = PUD.OrderNo1 " & vbCrLf & _
                          " LEFT JOIN MS_AFFILIATE MA ON MA.AFFILIATEID = POM.AFFILIATEID " & vbCrLf & _
                          " LEFT JOIN MS_SUPPLIER MS ON MS.SUPPLIERID = POM.SUPPLIERID " & vbCrLf

        ls_sql = ls_sql + " LEFT JOIN MS_PARTS MP ON MP.PARTNO = POD.PARTNO " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo " & vbCrLf & _
                          " AND MPM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          " AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
                          " LEFT JOIN MS_FORWARDER MF ON MF.FORWARDERID = POM.FORWARDERID " & vbCrLf & _
                          " LEFT JOIN MS_UNITCLS MU ON MU.UNITCLS = MP.UNITCLS " & vbCrLf & _
                          " LEFT JOIN MS_SHIPCLS MSC ON MSC.SHIPCLS = POM.SHIPCLS " & vbCrLf & _
                          " where POD.PONO = '" & Trim(PONO) & "' " & vbCrLf & _
                          " AND POD.AffiliateID = '" & Trim(affiliateID) & "'" & vbCrLf & _
                          " AND POD.SupplierID = '" & Trim(SupplierID) & "' " & vbCrLf & _
                          " AND POD.OrderNO1 = '" & Trim(OrderNo) & "' " & vbCrLf

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Private Function sendEmailtoSupllierDeliveryConfirmation(ByVal pStatus As String, ByVal pnamaFile1 As String, ByVal pnamaFile2 As String, ByVal pPono As String, ByVal pFWD As String, ByVal pAFF As String, ByVal pOrderNo As String, ByVal pSupplier As String, ByVal pSj As String) As Boolean
        Try
            Dim TempFilePath As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim TempFileName1 As String
            Dim TempFileName2 As String

            TempFilePath = Trim(txtSaveAsDOM.Text)

            Dim pFile As String = ""

            TempFileName1 = pnamaFile1
            TempFileName2 = pnamaFile2

            TempFilePath = Trim(txtSaveAsDOM.Text)

            Dim dsEmail As New DataSet
            dsEmail = EmailSendForwarder_Export(pFWD, pAFF)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    If fromEmail = "" Then
                        fromEmail = dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                    Else
                        fromEmail = fromEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") <> "PASI" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") <> "PASI" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                    End If
                End If
            Next
            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PASI Final Approval STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailtoSupllierDeliveryConfirmation = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PASI Final Approval STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailtoSupllierDeliveryConfirmation = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)

            If pStatus = "RECEIVING" Then
                If pPono.Trim <> pOrderNo.Trim Then
                    'mailMessage.Subject = "[TRIAL] GOOD RECEIVING :" & Trim(pAFF) & "-" & Trim(pSupplier) & "-" & pPono & " Split (" & pOrderNo & ")"
                    mailMessage.Subject = "DN-" & Trim(pSupplier) & "-" & Trim(pPono) & " Split (" & pOrderNo & ")" & " Delivery Note From Supplier "
                Else
                    'mailMessage.Subject = "[TRIAL] GOOD RECEIVING :" & Trim(pAFF) & "-" & Trim(pSupplier) & "-" & pPono & ""
                    mailMessage.Subject = "DN-" & Trim(pSupplier) & "-" & Trim(pPono) & " Delivery Note From Supplier "
                End If
            ElseIf pStatus = "MOVING" Then
                mailMessage.Subject = "ML-" & Trim(pSupplier) & "-" & Trim(pPono) & " PO MOVING LIST "
            Else
                'mailMessage.Subject = "[TRIAL] PASI FINAL APPROVAL :" & Trim(pAFF) & "-" & Trim(pSupplier) & "-" & pPono & ""
                If pPono.Trim = pOrderNo.Trim Then
                    mailMessage.Subject = "PO-" & Trim(pSupplier) & "-" & Trim(pPono) & " PASI FINAL APPROVAL "
                Else
                    mailMessage.Subject = "PO-" & Trim(pSupplier) & "-" & Trim(pPono) & " Split (" & Trim(pOrderNo) & ") PASI FINAL APPROVAL "
                End If
            End If


            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If
            'Dim receiptBCCEmail As String = "pasi.purchase@gmail.com"
            'If receiptBCCEmail <> "" Then
            '    For Each recipientBCC In receiptBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If

            GetSettingEmail_Export("DELIVERY")
            'uf_GetNotification("40")
            'ls_Body = pLine1 & vbCr & pLine2 & vbCr & "Kanban No:" & KKanbanno1 & "," & KKanbanno2 & "," & KKanbanno3 & "," & KKanbanno4 & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
            'ls_Body = clsNotification.GetNotification("40", "", "", pPono)
            'If KstatusKanban = True Then
            '    mailMessage.Body = ls_Body
            'Else
            If pStatus = "RECEIVING" Then
                If Trim(pPono) <> Trim(pOrderNo) Then
                    ls_Body = clsNotification.GetNotification("20", "", pPono.Trim & " Split (" & pOrderNo.Trim & ")", , pSj.Trim)
                Else
                    ls_Body = clsNotification.GetNotification("20", "", pPono.Trim, , pSj.Trim)
                End If
            ElseIf pStatus = "MOVING" Then
                ls_Body = clsNotification.GetNotification("24", "", pPono.Trim)
            Else
                If pPono <> pOrderNo Then
                    ls_Body = clsNotification.GetNotification("19", "", pPono.Trim & " Split (" & pOrderNo.Trim & ")")
                Else
                    ls_Body = clsNotification.GetNotification("19", "", pPono.Trim)
                End If
            End If

            mailMessage.Body = ls_Body '"Dear Sir/Madam,  This is notification for: Receiving Forwarder Confirmation : (" & Trim(pPono) & ")"
            'End If

            Dim filename1 As String
            Dim fi1 As New FileInfo(TempFilePath & TempFileName1)
            If fi1.Exists Then
                filename1 = TempFilePath & TempFileName1
                mailMessage.Attachments.Add(New Attachment(filename1))
            End If

            If TempFileName1 <> "" Then
                Dim filename2 As String
                Dim fi2 As New FileInfo(TempFilePath & TempFileName2)
                If fi2.Exists Then
                    filename2 = TempFilePath & TempFileName2
                    mailMessage.Attachments.Add(New Attachment(filename2))
                End If
            End If

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
            sendEmailtoSupllierDeliveryConfirmation = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
        Catch ex As Exception
            sendEmailtoSupllierDeliveryConfirmation = False
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text

        End Try

    End Function

    Private Function EmailSendForwarder_Export(ByVal pFWD As String, ByVal pAff As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = " select Flag = 'FWD', KanbanCC = isnull(POExportCC,''), KanbanTO = isnull(POExportTo,''), KanbanFrom ='' From ms_emailForwarder where ForwarderID = '" & Trim(pFWD) & "'" & vbCrLf & _
                 " union ALL " & vbCrLf & _
                 " select Flag = 'PASI', kanbanCC = isnull(AffiliatePOCC,'') , kanbanTo = isnull(AffiliatePOTo,''), kanbanFrom = isnull(AffiliatePOTo,'') from MS_EmailPasi_Export  " & vbCrLf & _
                 " select Flag = 'AFF', kanbanCC = isnull(AffiliatePOCC,'') , kanbanTo = isnull(AffiliatePOTo,''), kanbanFrom = isnull(AffiliatePOTo,'') from MS_EmailAffiliate_Export where AffiliateID = '" & Trim(pAff) & "' " & vbCrLf
        'ls_SQL = " select Flag = 'AFF', kanbanCC = isnull(kanbanCC,'') ,kanbanTo = isnull(KanbanTo,''), kanbanFrom = '' from MS_emailAffiliate where AffiliateID = '" & pAfffCode & "' " & vbCrLf & _
        '         " union ALL " & vbCrLf & _
        '         " select Flag = 'PASI', kanbanCC = isnull(kanbanCC,'') , kanbanTo = isnull(KanbanTo,''), kanbanFrom = isnull(KanbanTo,'') from MS_EmailPasi  " & vbCrLf & _
        '         " UNION ALL " & vbCrLf & _
        '         " select Flag = 'SUPP', kanbanCC = isnull(kanbanCC,'') , kanbanTo = isnull(kanbanTo,''), KanbanFrom = '' from MS_EmailSupplier where supplierID = '" & pSupplierID & "' "
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Sub UpdateStatusPOExport(ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pOrderNo1 As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " update dbo.PO_Master_Export set FinalApprovalCls = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND PoNo = '" & pPoNo & "'" & vbCrLf & _
                         " AND OrderNo1 = '" & pOrderNo1 & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub

    Private Sub Excel_DeliveryConfirmationReplacement()
        Dim strFileSize As String = ""
        Dim ls_sql As String = ""

        Dim xlApp = New Excel.Application
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim sheetNumber2 As Integer = 3
        Dim i As Integer, xi As Integer
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        Dim NewFileCopy As String


        Dim ds As New DataSet
        Dim dsSplit As New DataSet
        Dim dsHeader As New DataSet
        Dim dsDetail As New DataSet
        Dim dsETAETD As New DataSet
        Dim dsDetailDelivery As New DataSet

        Dim ls_SJ As String = ""
        Dim ls_Supplier As String = ""
        Dim ls_supplierName As String = ""
        Dim ls_supplierAdd As String = ""
        Dim ls_delivery As String = ""
        Dim ls_Attn As String = ""
        Dim ls_Telp As String = ""
        Dim ls_DeliveryName As String = ""
        Dim ls_deliveryAdd As String = ""
        Dim ls_orderNo As String = ""
        Dim ls_orderNoSplit As String = ""
        Dim ls_ETDV As Date
        Dim ls_ETDP As Date
        Dim ls_ETAP As Date
        Dim ls_ETAF As Date
        Dim ls_Aff As String = ""
        Dim ls_Consignee As String = ""
        Dim ls_AFFName As String = ""
        Dim ls_AffADD As String
        Dim ls_PEriod As String = ""
        Dim ls_FinalApprovalCls As String = ""
        Dim ls_orderNoReff As String = ""
        Dim ls_sts As String = ""

        Dim i_loop As Long
        Dim pFilename_e As String = ""
        Dim pFileName_e2 As String = ""
        Dim adaRemaining As Boolean = False
        Dim dsR As New DataSet
        Dim ls_filter As String = "", ls_filterPart As String = ""
        Dim ls_StatusRemaining As String = ""
        Dim sDefect As String = ""
        Dim booSplit As Boolean = False

        Dim dsSeqNo As New DataSet

        Try
            ls_sql = " select  " & vbCrLf & _
                      " 	distinct a.SuratJalanNo, a.SupplierID, a.AffiliateID, a.PONo, a.OrderNo, b.ForwarderID, c.ETDPort1, c.ETAFactory1, " & vbCrLf & _
                      " 	c.ETAPort1, c.ETDVendor1, d.SupplierName, d.Address SUPPAddress, " & vbCrLf & _
                      " 	attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), c.Period,  " & vbCrLf & _
                      "    MF.ForwarderName, isnull(MF.Address,'')  FWDAddress, MA.AffiliateName as AFFName, isnull(MA.Address,'') AFFAddress " & vbCrLf & _
                      " from ReceiveForwarder_DetailBox a  " & vbCrLf & _
                      " left join ReceiveForwarder_Master b on a.SuratJalanNo = b.SuratJalanNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.PONo and a.OrderNo = b.OrderNo" & vbCrLf & _
                      " left join MS_Supplier d on a.SupplierID = d.SupplierID " & vbCrLf & _
                      " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = b.ForwarderID " & vbCrLf & _
                      " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = a.AffiliateID " & vbCrLf & _
                      " left join PO_Master_Export c on a.AffiliateID= c.AffiliateID and a.SupplierID = c.SupplierID and a.PONo = c.PONo and a.OrderNo = c.OrderNo1 "

            ls_sql = ls_sql + " 	  " & vbCrLf & _
                              " where StatusDefect = '1' and ISNULL(a.ExcelCls,1) = 1 " & vbCrLf & _
                              "  "

            ds = cls.uf_GetDataSet(ls_sql)

            Dim ls_file As String = ""
            For i_loop = 0 To ds.Tables(0).Rows.Count - 1
                pFilename_e = "" : pFileName_e2 = ""

                '================================DELIVERY CONFIRMATION==================================
                Dim fi As New FileInfo(Trim(txtAttachmentDOM.Text) & "\Template Customer Delivery Confirmation.xlsm")

                If Not fi.Exists Then
                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Customer Delivery Confirmation STOPPED, because File Excel isn't Found " & vbCrLf & _
                                    rtbProcess.Text
                    Exit Sub
                End If

                ls_SJ = Trim(ds.Tables(0).Rows(i_loop)("SuratJalanNo"))
                ls_orderNoSplit = Trim(ds.Tables(0).Rows(i_loop)("OrderNo"))
                ls_Supplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))
                ls_supplierName = Trim(ds.Tables(0).Rows(i_loop)("SupplierName"))
                ls_supplierAdd = Trim(ds.Tables(0).Rows(i_loop)("SuppAddress"))
                ls_delivery = Trim(ds.Tables(0).Rows(i_loop)("ForwarderID"))
                ls_DeliveryName = Trim(ds.Tables(0).Rows(i_loop)("ForwarderName"))
                ls_deliveryAdd = Trim(ds.Tables(0).Rows(i_loop)("FWDAddress"))
                ls_orderNo = Trim(ds.Tables(0).Rows(i_loop)("PONo"))
                ls_ETDV = Format((ds.Tables(0).Rows(i_loop)("ETDVendor1")), "yyyy-MM-dd")
                ls_ETDP = Format((ds.Tables(0).Rows(i_loop)("ETDPort1")), "yyyy-MM-dd")
                ls_ETAP = Format((ds.Tables(0).Rows(i_loop)("ETAPort1")), "yyyy-MM-dd")
                ls_ETAF = Format((ds.Tables(0).Rows(i_loop)("ETAFactory1")), "yyyy-MM-dd")
                ls_Aff = Trim(ds.Tables(0).Rows(i_loop)("AffiliateID"))
                'ls_Consignee = Trim(ds.Tables(0).Rows(i_loop)("Consignee"))
                ls_AFFName = Trim(ds.Tables(0).Rows(i_loop)("AFFName"))
                ls_AffADD = Trim(ds.Tables(0).Rows(i_loop)("AFFAddress"))
                ls_PEriod = Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy-MM")
                ls_Attn = Trim(ds.Tables(0).Rows(i_loop)("attn"))
                ls_Telp = Trim(ds.Tables(0).Rows(i_loop)("telp"))
                'ls_sts = Trim(ds.Tables(0).Rows(i_loop)("sts"))

                dsDetailDelivery = BidDataDeliveryConfirmReplacement(ls_orderNo, ls_orderNoSplit, ls_Aff, ls_Supplier, ls_SJ)

                If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                    Dim dsEmail As New DataSet
                    dsEmail = EmailToEmailCCKanban_Export(ls_Aff, ls_Supplier)
                    '1 CC Affiliate
                    '2 CC PASI
                    '3 CC & TO Supplier
                    For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(i)("KanbanCC")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("KanbanCC")
                        End If

                        If dsEmail.Tables(0).Rows(i)("KanbanTO") <> "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(i)("KanbanTO")
                        End If
                    Next

                    Dim k As Long
                    Dim dsAffiliate As New DataSet
                    dsAffiliate = Affiliate(Trim(ls_Aff))

                    Dim dsSupplier As New DataSet
                    dsSupplier = Supplier(Trim(ls_Supplier))

                    NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Template Customer Delivery Confirmation.xlsm"
                    ls_file = NewFileCopy
                    ExcelBook = xlApp.Workbooks.Open(ls_file)
                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                    ExcelSheet.Range("H2").Value = receiptEmail.Trim
                    ExcelSheet.Range("H3").Value = ls_Aff.Trim
                    ExcelSheet.Range("H4").Value = ls_delivery.Trim
                    ExcelSheet.Range("H5").Value = ls_Supplier.Trim

                    ExcelSheet.Range("AP11").Value = "NO"

                    ExcelSheet.Range("I11:X11").Value = ls_supplierName.Trim
                    ExcelSheet.Range("I12:X15").Value = ls_supplierAdd.Trim

                    ExcelSheet.Range("I19:X19").Value = ls_DeliveryName.Trim
                    ExcelSheet.Range("I20:X22").Value = ls_deliveryAdd.Trim
                    ExcelSheet.Range("I23:X23").Value = "ATTN : " & Trim(ls_Attn) & "     TELP : " & Trim(ls_Telp)

                    ExcelSheet.Range("AE19:AT19").Value = Trim(dsAffiliate.Tables(0).Rows(0)("ConsigneeName"))
                    ExcelSheet.Range("AE20:AT22").Value = Trim(dsAffiliate.Tables(0).Rows(0)("ConsigneeAddress"))

                    ExcelSheet.Range("AE11:AI11").Value = ls_PEriod

                    ExcelSheet.Range("AE13:AI13").Value = ls_orderNoSplit.Trim

                    'dsSeqNo = BidDataDeliveryConfirmReplacement(ls_orderNo, ls_orderNoSplit, ls_Aff, ls_Supplier, ls_SJ)

                    ExcelSheet.Range("I28:P28").Value = ls_SJ.Trim & "-R1"

                    If ls_orderNo <> ls_orderNoSplit Then
                        ExcelSheet.Range("AE15:AI15").Value = ls_orderNo.Trim
                    End If

                    ExcelSheet.Range("AE17:AI17").Value = ls_ETDV

                    k = 0
                    For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                        ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Merge()
                        ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Merge()
                        ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Merge()
                        ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Merge()
                        ExcelSheet.Range("W" & k + 34 & ": Y" & k + 34).Merge()
                        ExcelSheet.Range("Z" & k + 34 & ": AB" & k + 34).Merge()
                        ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Merge()
                        ExcelSheet.Range("AE" & k + 34 & ": AF" & k + 34).Merge()
                        ExcelSheet.Range("AG" & k + 34 & ": AJ" & k + 34).Merge()
                        ExcelSheet.Range("AK" & k + 34 & ": AN" & k + 34).Merge()
                        ExcelSheet.Range("AO" & k + 34 & ": AR" & k + 34).Merge()
                        ExcelSheet.Range("AS" & k + 34 & ": AV" & k + 34).Merge()

                        ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Value = k + 1
                        ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Value = ls_orderNo '
                        ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Partno"))
                        ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("PartName")
                        ExcelSheet.Range("W" & k + 34 & ": Y" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelno1"))
                        ExcelSheet.Range("Z" & k + 34 & ": AB" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelno2"))
                        ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("UOM")
                        ExcelSheet.Range("AE" & k + 34 & ": AF" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("MOQ")
                        ExcelSheet.Range("AG" & k + 34 & ": AJ" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                        ExcelSheet.Range("AG" & i + 34 & ": AJ" & i + 34).NumberFormat = "#,##0"
                        ExcelSheet.Range("AK" & k + 34 & ": AN" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                        ExcelSheet.Range("AK" & i + 34 & ": AN" & i + 34).NumberFormat = "#,##0"
                        ExcelSheet.Range("AO" & k + 34 & ": AR" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                        ExcelSheet.Range("AO" & i + 34 & ": AR" & i + 34).NumberFormat = "#,##0"
                        ExcelSheet.Range("AS" & k + 34 & ": AV" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalPOQty")
                        ExcelSheet.Range("AS" & i + 34 & ": AV" & i + 34).NumberFormat = "#,##0"
                        k = k + 1
                    Next

                    ExcelSheet.Range("B35").Interior.Color = Color.White
                    ExcelSheet.Range("B35").Font.Color = Color.Black
                    ExcelSheet.Range("B" & k + 34).Value = "E"
                    ExcelSheet.Range("B" & k + 34).Interior.Color = Color.Black
                    ExcelSheet.Range("B" & k + 34).Font.Color = Color.White

                    k = k - 1
                    DrawAllBorders(ExcelSheet.Range("B34" & ": AV" & k + 34))
                    ExcelSheet.Range("AM34" & ": AP" & k + 34).Interior.Color = Color.Yellow
                    ExcelSheet.Range("W34" & ": AB" & k + 34).Interior.Color = Color.Yellow

                    'Save ke Local
                    xlApp.DisplayAlerts = False

                    If ls_orderNo <> ls_orderNoSplit Then
                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\REPLACEMENT DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm")
                        pFilename_e = "\REPLACEMENT DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm"
                    Else
                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\REPLACEMENT DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm")
                        pFilename_e = "\REPLACEMENT DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm"
                    End If

                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
                '================================DELIVERY CONFIRMATION==================================

                If ls_orderNo <> ls_orderNoSplit Then
                    If sendEmailPASI_EXPORTReplacement("Delivery Confirmation", "Replacement Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier & " Split (" & Trim(ls_orderNoSplit) & ")", ls_Supplier, pFilename_e, pFileName_e2, ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                Else
                    If sendEmailPASI_EXPORTReplacement("Delivery Confirmation", "Replacement Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier, ls_Supplier, pFilename_e, pFileName_e2, ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                End If

                Call UpdateStatusPOExportReplacement(ls_Aff, ls_Supplier, ls_orderNo, ls_orderNoSplit, ls_SJ)

keluar:
                xlApp.Workbooks.Close()
                xlApp.Quit()
            Next
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation to Supplier STOPPED, because " & Err.Description & " " & vbCrLf & _
                                    rtbProcess.Text
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Function BidDataDeliveryConfirmReplacement(ByVal PONO As String, ByVal OrderNo As String, ByVal affiliateID As String, ByVal SupplierID As String, ByVal filter As String)
        Dim ls_sql As String
        'MdlConn.ReadConnection()

        ls_sql = " select a.PartNo, b.PartName,  " & vbCrLf & _
                  " 	a.Label1 as labelno1, a.Label2 as labelno2, UOM = c.Description, MOQ = ISNULL(PDE.POMOQ,d.MOQ), " & vbCrLf & _
                  " 	SuppQty = a.Box * ISNULL(PDE.POQtyBox,d.QtyBox), TotalPOQty = a.Box  " & vbCrLf & _
                  " from ReceiveForwarder_DetailBox a " & vbCrLf & _
                  " left join PO_Detail_Export PDE on a.PONo = PDE.PONo and a.AffiliateID = PDE.AffiliateID and a.SupplierID = PDE.SupplierID and a.PartNo = PDE.PartNo " & vbCrLf & _
                  " left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
                  " left join MS_UnitCls c on b.UnitCls = c.UnitCls " & vbCrLf & _
                  " left join MS_PartMapping d on a.PartNo = d.PartNo and d.SupplierID = a.SupplierID and d.AffiliateID = a.AffiliateID " & vbCrLf & _
                  " where a.SuratJalanNo = '" & filter & "' and a.AffiliateID = '" & affiliateID & "' and a.SupplierID = '" & SupplierID & "' and a.PONo = '" & PONO & "' and a.OrderNo = '" & OrderNo & "' and StatusDefect = '1'" & vbCrLf & _
                  "  "
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Private Function sendEmailPASI_EXPORTReplacement(ByVal pStatus As String, ByVal pCaption As String, ByVal pSupplier1 As String, ByVal pAtt1 As String, ByVal pAtt2 As String, ByVal pPono As String, ByVal pOrderNo As String) As Boolean 'Link Affiliate Order Entry
        Try
            Dim TempFilePath As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCCPO_Export("", "PASI", pSupplier1)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPASI_EXPORTReplacement = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPASI_EXPORTReplacement = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "Send " & pCaption

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            'Dim receiptBCCEmail As String
            'If pSupplier1.Trim.ToUpper = "KMK" Then
            '    receiptBCCEmail = "Ariefrosyadi91@gmail.com;Ristian_a@yahoo.com;pasi.purchase@gmail.com"
            'Else
            '    receiptBCCEmail = "pasi.purchase@gmail.com"
            'End If
            'If receiptBCCEmail <> "" Then
            '    For Each recipientBCC In receiptBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If
            GetSettingEmail_Export("Kanban")
            'uf_GetNotification("11")
            'ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
            'mailMessage.Body = ls_Body
            If pPono.Trim = pOrderNo.Trim Then
                ls_Body = clsNotification.GetNotification("18", "", pPono.Trim)
            Else
                ls_Body = clsNotification.GetNotification("18", "", pPono.Trim & " Split (" & pOrderNo.Trim & ")")
            End If
            mailMessage.Body = ls_Body
            'mailMessage.Body = "Dear Sir/Madam,  This is notification for: " & pCaption
            Dim filename As String = TempFilePath & pAtt1
            If pAtt1 <> "" Then mailMessage.Attachments.Add(New Attachment(filename))
            Dim filename2 As String = TempFilePath & pAtt2
            If pAtt2 <> "" Then mailMessage.Attachments.Add(New Attachment(filename2))
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            'smtp.Host = "smtp.atisicloud.com"
            'smtp.Host = "mail.fast.net.id"

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
            'Delete the file
            'Kill(TempFilePath & TempFileName)
            sendEmailPASI_EXPORTReplacement = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send DN No: " & pPono & " to Supplier SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send DN No: " & pPono & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
            sendEmailPASI_EXPORTReplacement = False
        End Try
    End Function

    Private Sub UpdateStatusPOExportReplacement(ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pOrderNo1 As String, ByVal pSuratJalanNO As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " update dbo.ReceiveForwarder_DetailBox set ExcelCls = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND PoNo = '" & pPoNo & "'" & vbCrLf & _
                         " AND OrderNo = '" & pOrderNo1 & "'" & vbCrLf & _
                         " AND SuratJalanNo = '" & pSuratJalanNO & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub

    Private Sub Excel_DeliveryConfirmationRemaining()
        Dim strFileSize As String = ""
        Dim ls_sql As String = ""

        Dim xlApp = New Excel.Application
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim sheetNumber2 As Integer = 3
        Dim i As Integer, xi As Integer
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        Dim NewFileCopy As String


        Dim ds As New DataSet
        Dim dsSplit As New DataSet
        Dim dsHeader As New DataSet
        Dim dsDetail As New DataSet
        Dim dsETAETD As New DataSet
        Dim dsDetailDelivery As New DataSet

        Dim ls_SJ As String = ""
        Dim ls_Supplier As String = ""
        Dim ls_supplierName As String = ""
        Dim ls_supplierAdd As String = ""
        Dim ls_delivery As String = ""
        Dim ls_Attn As String = ""
        Dim ls_Telp As String = ""
        Dim ls_DeliveryName As String = ""
        Dim ls_deliveryAdd As String = ""
        Dim ls_orderNo As String = ""
        Dim ls_orderNoSplit As String = ""
        Dim ls_ETDV As Date
        Dim ls_ETDP As Date
        Dim ls_ETAP As Date
        Dim ls_ETAF As Date
        Dim ls_Aff As String = ""
        Dim ls_Consignee As String = ""
        Dim ls_AFFName As String = ""
        Dim ls_AffADD As String
        Dim ls_PEriod As String = ""
        Dim ls_FinalApprovalCls As String = ""
        Dim ls_orderNoReff As String = ""
        Dim ls_sts As String = ""
        Dim ls_Commercial As String = ""

        Dim i_loop As Long
        Dim pFilename_e As String = ""
        Dim pFileName_e2 As String = ""
        Dim adaRemaining As Boolean = False
        Dim dsR As New DataSet
        Dim ls_filter As String = "", ls_filterPart As String = ""
        Dim ls_StatusRemaining As String = ""
        Dim sDefect As String = ""
        Dim booSplit As Boolean = False

        Dim dsSeqNo As New DataSet

        Try
            ls_sql = "SELECT DISTINCT /*b.SuratJalanNo,*/ b.SupplierID, b.AffiliateID, b.PONo, b.OrderNo, b.ForwarderID, " & vbCrLf & _
                "c.ETDPort1, c.ETAFactory1, c.ETAPort1, c.ETDVendor1, d.SupplierName, d.Address SUPPAddress, " & vbCrLf & _
                "attn = ISNULL(MF.Attn,''), telp = ISNULL(MF.MobilePhone,''), c.Period, " & vbCrLf & _
                "MF.ForwarderName, ISNULL(MF.Address,'') FWDAddress, MA.AffiliateName as AFFName, ISNULL(MA.Address,'') AFFAddress, PME.CommercialCls " & vbCrLf & _
                "FROM PrintLabelExport a " & vbCrLf & _
                "INNER JOIN ReceiveForwarder_Master b ON a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.PONo = b.PONo AND a.OrderNo = b.OrderNo " & vbCrLf & _
                "LEFT JOIN PO_Master_Export PME ON a.PONo = PME.PONo And a.AffiliateID = PME.AffiliateID And a.SupplierID = PME.SupplierID And a.OrderNo = PME.OrderNo1" & vbCrLf & _
                "LEFT JOIN MS_Supplier d ON a.SupplierID = d.SupplierID " & vbCrLf & _
                "LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = b.ForwarderID " & vbCrLf & _
                "LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = a.AffiliateID " & vbCrLf & _
                "LEFT JOIN PO_Master_Export c ON a.AffiliateID= c.AffiliateID and a.SupplierID = c.SupplierID and a.PONo = c.PONo and a.OrderNo = c.OrderNo1 " & vbCrLf & _
                "WHERE a.StatusRemaining = '1' And a.SuratJalanNo_FWD = ''"

            ds = cls.uf_GetDataSet(ls_sql)

            Dim ls_file As String = ""
            For i_loop = 0 To ds.Tables(0).Rows.Count - 1
                pFilename_e = "" : pFileName_e2 = ""

                '================================DELIVERY CONFIRMATION==================================
                Dim fi As New FileInfo(Trim(txtAttachmentDOM.Text) & "\Template Customer Delivery Confirmation.xlsm")

                If Not fi.Exists Then
                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Customer Delivery Confirmation STOPPED, because File Excel isn't Found " & vbCrLf & _
                                    rtbProcess.Text
                    Exit Sub
                End If

                'ls_SJ = Trim(ds.Tables(0).Rows(i_loop)("SuratJalanNo"))
                ls_orderNoSplit = Trim(ds.Tables(0).Rows(i_loop)("OrderNo"))
                ls_Supplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))
                ls_supplierName = Trim(ds.Tables(0).Rows(i_loop)("SupplierName"))
                ls_supplierAdd = Trim(ds.Tables(0).Rows(i_loop)("SuppAddress"))
                ls_delivery = Trim(ds.Tables(0).Rows(i_loop)("ForwarderID"))
                ls_DeliveryName = Trim(ds.Tables(0).Rows(i_loop)("ForwarderName"))
                ls_deliveryAdd = Trim(ds.Tables(0).Rows(i_loop)("FWDAddress"))
                ls_orderNo = Trim(ds.Tables(0).Rows(i_loop)("PONo"))
                ls_ETDV = Format((ds.Tables(0).Rows(i_loop)("ETDVendor1")), "yyyy-MM-dd")
                ls_ETDP = Format((ds.Tables(0).Rows(i_loop)("ETDPort1")), "yyyy-MM-dd")
                ls_ETAP = Format((ds.Tables(0).Rows(i_loop)("ETAPort1")), "yyyy-MM-dd")
                ls_ETAF = Format((ds.Tables(0).Rows(i_loop)("ETAFactory1")), "yyyy-MM-dd")
                ls_Aff = Trim(ds.Tables(0).Rows(i_loop)("AffiliateID"))
                ls_AFFName = Trim(ds.Tables(0).Rows(i_loop)("AFFName"))
                ls_AffADD = Trim(ds.Tables(0).Rows(i_loop)("AFFAddress"))
                ls_PEriod = Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy-MM")
                ls_Attn = Trim(ds.Tables(0).Rows(i_loop)("attn"))
                ls_Telp = Trim(ds.Tables(0).Rows(i_loop)("telp"))
                ls_Commercial = Trim(ds.Tables(0).Rows(i_loop)("CommercialCls"))

                dsDetailDelivery = BidDataDeliveryConfirmRemaining(ls_orderNo, ls_orderNoSplit, ls_Aff, ls_Supplier, ls_SJ)

                If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                    Dim dsEmail As New DataSet
                    dsEmail = EmailToEmailCCKanban_Export(ls_Aff, ls_Supplier)
                    '1 CC Affiliate
                    '2 CC PASI
                    '3 CC & TO Supplier
                    For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(i)("KanbanCC")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("KanbanCC")
                        End If

                        If dsEmail.Tables(0).Rows(i)("KanbanTO") <> "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(i)("KanbanTO")
                        End If
                    Next

                    Dim k As Long
                    Dim dsAffiliate As New DataSet
                    dsAffiliate = Affiliate(Trim(ls_Aff))

                    Dim dsSupplier As New DataSet
                    dsSupplier = Supplier(Trim(ls_Supplier))

                    NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Template Customer Delivery Confirmation.xlsm"
                    ls_file = NewFileCopy
                    ExcelBook = xlApp.Workbooks.Open(ls_file)
                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                    ExcelSheet.Range("H2").Value = receiptEmail.Trim
                    ExcelSheet.Range("H3").Value = ls_Aff.Trim
                    ExcelSheet.Range("H4").Value = ls_delivery.Trim
                    ExcelSheet.Range("H5").Value = ls_Supplier.Trim

                    ExcelSheet.Range("A7").Value = "REMAINING DELIVERY NOTE (EXPORT)"

                    If ls_Commercial = "0" Then
                        ExcelSheet.Range("AP11").Value = "NO"
                    Else
                        ExcelSheet.Range("AP11").Value = "YES"
                    End If


                    ExcelSheet.Range("I11:X11").Value = ls_supplierName.Trim
                    ExcelSheet.Range("I12:X15").Value = ls_supplierAdd.Trim

                    ExcelSheet.Range("I19:X19").Value = ls_DeliveryName.Trim
                    ExcelSheet.Range("I20:X22").Value = ls_deliveryAdd.Trim
                    ExcelSheet.Range("I23:X23").Value = "ATTN : " & Trim(ls_Attn) & "     TELP : " & Trim(ls_Telp)

                    ExcelSheet.Range("AE19:AT19").Value = Trim(dsAffiliate.Tables(0).Rows(0)("ConsigneeName"))
                    ExcelSheet.Range("AE20:AT22").Value = Trim(dsAffiliate.Tables(0).Rows(0)("ConsigneeAddress"))

                    ExcelSheet.Range("AE11:AI11").Value = ls_PEriod

                    ExcelSheet.Range("AE13:AI13").Value = ls_orderNoSplit.Trim

                    If ls_orderNo <> ls_orderNoSplit Then
                        ExcelSheet.Range("AE15:AI15").Value = ls_orderNo.Trim
                    End If

                    ExcelSheet.Range("AE17:AI17").Value = ls_ETDV

                    k = 0
                    For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                        ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Merge()
                        ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Merge()
                        ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Merge()
                        ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Merge()
                        ExcelSheet.Range("W" & k + 34 & ": Y" & k + 34).Merge()
                        ExcelSheet.Range("Z" & k + 34 & ": AB" & k + 34).Merge()
                        ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Merge()
                        ExcelSheet.Range("AE" & k + 34 & ": AF" & k + 34).Merge()
                        ExcelSheet.Range("AG" & k + 34 & ": AJ" & k + 34).Merge()
                        ExcelSheet.Range("AK" & k + 34 & ": AN" & k + 34).Merge()
                        ExcelSheet.Range("AO" & k + 34 & ": AR" & k + 34).Merge()
                        ExcelSheet.Range("AS" & k + 34 & ": AV" & k + 34).Merge()

                        ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Value = k + 1
                        ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Value = ls_orderNo '
                        ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Partno"))
                        ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("PartName")
                        ExcelSheet.Range("W" & k + 34 & ": Y" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("LabelNo1"))
                        ExcelSheet.Range("Z" & k + 34 & ": AB" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("LabelNo2"))
                        ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("UOM")
                        ExcelSheet.Range("AE" & k + 34 & ": AF" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("MOQ")
                        ExcelSheet.Range("AG" & k + 34 & ": AJ" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                        ExcelSheet.Range("AG" & i + 34 & ": AJ" & i + 34).NumberFormat = "#,##0"
                        ExcelSheet.Range("AK" & k + 34 & ": AN" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("OrderQty")
                        ExcelSheet.Range("AK" & i + 34 & ": AN" & i + 34).NumberFormat = "#,##0"
                        ExcelSheet.Range("AO" & k + 34 & ": AR" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                        ExcelSheet.Range("AO" & i + 34 & ": AR" & i + 34).NumberFormat = "#,##0"
                        ExcelSheet.Range("AS" & k + 34 & ": AV" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalPOQty")
                        ExcelSheet.Range("AS" & i + 34 & ": AV" & i + 34).NumberFormat = "#,##0"
                        k = k + 1
                    Next

                    ExcelSheet.Range("B35").Interior.Color = Color.White
                    ExcelSheet.Range("B35").Font.Color = Color.Black
                    ExcelSheet.Range("B" & k + 34).Value = "E"
                    ExcelSheet.Range("B" & k + 34).Interior.Color = Color.Black
                    ExcelSheet.Range("B" & k + 34).Font.Color = Color.White

                    k = k - 1
                    DrawAllBorders(ExcelSheet.Range("B34" & ": AV" & k + 34))
                    ExcelSheet.Range("AM34" & ": AP" & k + 34).Interior.Color = Color.Yellow
                    ExcelSheet.Range("W34" & ": AB" & k + 34).Interior.Color = Color.Yellow

                    'Save ke Local
                    xlApp.DisplayAlerts = False

                    If ls_orderNo <> ls_orderNoSplit Then
                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\REMAINING DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm")
                        pFilename_e = "\REMAINING DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNoSplit) & ")-" & Trim(ls_Supplier) & ".xlsm"
                    Else
                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\REMAINING DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm")
                        pFilename_e = "\REMAINING DELIVERY CONFIRMATION-" & Trim(ls_orderNo) & "-" & Trim(ls_Supplier) & ".xlsm"
                    End If

                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If
                '================================DELIVERY CONFIRMATION==================================

                If ls_orderNo <> ls_orderNoSplit Then
                    If sendEmailPASI_EXPORTRemaining("Delivery Confirmation", "Remaining Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier & " Split (" & Trim(ls_orderNoSplit) & ")", ls_Supplier, pFilename_e, pFileName_e2, ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                Else
                    If sendEmailPASI_EXPORTRemaining("Delivery Confirmation", "Remaining Delivery Confirmation : " & Trim(ls_orderNo) & "-" & ls_Supplier, ls_Supplier, pFilename_e, pFileName_e2, ls_orderNo, ls_orderNoSplit) = False Then GoTo keluar
                End If

                Call UpdateStatusPOExportRemaining(ls_Aff, ls_Supplier, ls_orderNo, ls_orderNoSplit, ls_SJ)

keluar:
                xlApp.Workbooks.Close()
                xlApp.Quit()
            Next
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation to Supplier STOPPED, because " & Err.Description & " " & vbCrLf & _
                                    rtbProcess.Text
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Function BidDataDeliveryConfirmRemaining(ByVal PONO As String, ByVal OrderNo As String, ByVal affiliateID As String, ByVal SupplierID As String, ByVal filter As String)
        Dim ls_sql As String
        'MdlConn.ReadConnection()

        'ls_sql = " select a.PartNo, b.PartName,  " & vbCrLf & _
        '          " 	a.Label1 as labelno1, a.Label2 as labelno2, UOM = c.Description, d.MOQ, " & vbCrLf & _
        '          " 	SuppQty = a.Box * d.QtyBox, TotalPOQty = a.Box  " & vbCrLf & _
        '          " from ReceiveForwarder_DetailBox a " & vbCrLf & _
        '          " left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
        '          " left join MS_UnitCls c on b.UnitCls = c.UnitCls " & vbCrLf & _
        '          " left join MS_PartMapping d on a.PartNo = d.PartNo and d.SupplierID = a.SupplierID and d.AffiliateID = a.AffiliateID " & vbCrLf & _
        '          " where a.SuratJalanNo = '" & filter & "' and a.AffiliateID = '" & affiliateID & "' and a.SupplierID = '" & SupplierID & "' and a.PONo = '" & PONO & "' and a.OrderNo = '" & OrderNo & "' and StatusDefect = '1'" & vbCrLf & _
        '          "  "
        ls_sql = "  select distinct  " & vbCrLf & _
                  "  Partno = POD.PartNo,  " & vbCrLf & _
                  "  PartName = MP.Partname,  " & vbCrLf & _
                  "  UOM = MUC.Description,  " & vbCrLf & _
                  "  MOQ = ISNULL(POD.POQtyBox,MPM.QtyBox),  " & vbCrLf & _
                  "  OrderQty = POD.Week1,   " & vbCrLf & _
                  "  labelno = Rtrim(PL.Label1) + ' - ' + Rtrim(pl.Label2),  " & vbCrLf & _
                  "  LabelNo1 = Rtrim(PL.Label1), LabelNo2 = Rtrim(PL.Label2),  " & vbCrLf & _
                  "  SuppQty = POD.week1 - ISNULL(RD.GoodRecQty,0),  " & vbCrLf & _
                  "  TotalPOQty = (POD.Week1 - ISNULL(RD.GoodRecQty,0)) / ISNULL(POD.POQtyBox,MPM.QtyBox)  " & vbCrLf & _
                  "  From PO_Detail_Export POD  " & vbCrLf

        ls_sql = ls_sql + "  LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                          "  LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo  " & vbCrLf & _
                          "  AND MPM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "  AND MPM.SupplierID = POD.SupplierID  " & vbCrLf & _
                          "  INNER JOIN (select pono, orderNo, affiliateID, SupplierID, PartNo, " & vbCrLf & _
                          "              Min(labelNo) as label1, Max(labelNo) as label2, statusRemaining from PrintLabelExport  " & vbCrLf & _
                          "              Group by pono, orderNo, affiliateID, SupplierID, PartNo, statusRemaining) PL ON PL.PONo = POD.PONo    " & vbCrLf & _
                          "          and PL.AffiliateID = POD.AffiliateID  AND PL.SupplierID = POD.SupplierID " & vbCrLf & _
                          "          AND PL.PartNO = POD.PartNo  AND PL.orderNo = POD.OrderNo1 " & vbCrLf & _
                          "  LEFT JOIN MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_Detail RD  " & vbCrLf

        ls_sql = ls_sql + " 		ON POD.PONo = RD.PONo " & vbCrLf & _
                          " 	   AND POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                          " 	   AND POD.SupplierID = RD.SupplierID " & vbCrLf & _
                          " 	   AND POD.PartNo = RD.PartNo " & vbCrLf & _
                          "  " & vbCrLf & _
                          "  where POD.PONO = '" & PONO & "'  " & vbCrLf & _
                          "  AND POD.AffiliateID = '" & affiliateID & "' " & vbCrLf & _
                          "  AND POD.SupplierID = '" & SupplierID & "'  " & vbCrLf & _
                          "  AND POD.OrderNo1 = '" & OrderNo & "'  " & vbCrLf & _
                          "  AND PL.statusRemaining = '1' " & vbCrLf

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Private Function sendEmailPASI_EXPORTRemaining(ByVal pStatus As String, ByVal pCaption As String, ByVal pSupplier1 As String, ByVal pAtt1 As String, ByVal pAtt2 As String, ByVal pPono As String, ByVal pOrderNo As String) As Boolean 'Link Affiliate Order Entry
        Try
            Dim TempFilePath As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCCPO_Export("", "PASI", pSupplier1)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPASI_EXPORTRemaining = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Confirmation STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPASI_EXPORTRemaining = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "Send " & pCaption

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            'Dim receiptBCCEmail As String
            'If pSupplier1.Trim.ToUpper = "KMK" Then
            '    receiptBCCEmail = "Ariefrosyadi91@gmail.com;Ristian_a@yahoo.com;pasi.purchase@gmail.com"
            'Else
            '    receiptBCCEmail = "pasi.purchase@gmail.com"
            'End If
            'If receiptBCCEmail <> "" Then
            '    For Each recipientBCC In receiptBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If
            GetSettingEmail_Export("Kanban")
            'uf_GetNotification("11")
            'ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
            'mailMessage.Body = ls_Body
            If pPono.Trim = pOrderNo.Trim Then
                ls_Body = clsNotification.GetNotification("18", "", pPono.Trim)
            Else
                ls_Body = clsNotification.GetNotification("18", "", pPono.Trim & " Split (" & pOrderNo.Trim & ")")
            End If
            mailMessage.Body = ls_Body
            'mailMessage.Body = "Dear Sir/Madam,  This is notification for: " & pCaption
            Dim filename As String = TempFilePath & pAtt1
            If pAtt1 <> "" Then mailMessage.Attachments.Add(New Attachment(filename))
            Dim filename2 As String = TempFilePath & pAtt2
            If pAtt2 <> "" Then mailMessage.Attachments.Add(New Attachment(filename2))
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            'smtp.Host = "smtp.atisicloud.com"
            'smtp.Host = "mail.fast.net.id"

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
            'Delete the file
            'Kill(TempFilePath & TempFileName)
            sendEmailPASI_EXPORTRemaining = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send DN No: " & pPono & " to Supplier SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send DN No: " & pPono & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
            sendEmailPASI_EXPORTRemaining = False
        End Try
    End Function

    Private Sub UpdateStatusPOExportRemaining(ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pOrderNo1 As String, ByVal pSuratJalanNO As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                'ls_SQL = " update dbo.ReceiveForwarder_DetailBox set ExcelCls = '2' " & vbCrLf & _
                '         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                '         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                '         " AND PoNo = '" & pPoNo & "'" & vbCrLf & _
                '         " AND OrderNo = '" & pOrderNo1 & "'" & vbCrLf & _
                '         " AND SuratJalanNo = '" & pSuratJalanNO & "'"
                ls_SQL = " UPDATE dbo.PrintLabelExport set statusRemaining = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND PoNo = '" & pPoNo & "'" & vbCrLf & _
                         " AND OrderNo = '" & pOrderNo1 & "'" & vbCrLf & _
                         " AND SuratJalanNo_FWD = ''"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub

    Private Sub Excel_DeliveryToForwarder()
        'On Error GoTo ErrHandler
        Dim strFileSize As String = ""
        Dim ls_sql As String = ""

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        'Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        ' Dim fileTocopy As String
        Dim NewFileCopy As String
        'Dim NewFileCopyas As String

        Dim KTime1 As String = ""
        Dim KTime2 As String = ""
        Dim KTime3 As String = ""
        Dim KTime4 As String = ""
        Dim pNamaFile1 As String = ""
        Dim pNamaFile2 As String = ""

        'Dim jkanbanno As String
        'Dim jQty As Long
        'Dim jQtyBox As Long
        'Dim jQtyPallet As Long

        Dim ds As New DataSet
        Dim dsSplit As New DataSet
        Dim dsHeader As New DataSet
        Dim dsDetail As New DataSet
        Dim dsETAETD As New DataSet
        Dim dsDetailDelivery As New DataSet
        Dim ls_commercial As String = ""

        Dim ls_SJ As String = ""
        Dim ls_Supplier As String = ""
        Dim ls_supplierName As String = ""
        Dim ls_supplierAdd As String = ""
        Dim ls_delivery As String = ""
        Dim ls_DeliveryName As String = ""
        Dim ls_deliveryAdd As String = ""
        Dim ls_orderNo As String = ""
        Dim ls_orderNo1 As String = ""
        Dim ls_ETDV As Date
        Dim ls_ETDP As Date
        Dim ls_ETAP As Date
        Dim ls_ETAF As Date
        Dim ls_Aff As String = ""
        Dim ls_AFFName As String = ""
        Dim ls_AffADD As String
        Dim ls_Attn As String = ""
        Dim ls_telp As String = ""
        Dim ls_Consignee As String = ""
        Dim ls_ConsName As String = ""
        Dim ls_ConsigneAdd As String = ""
        Dim i_loop As Long
        Dim xlApp = New Excel.Application
        Dim ls_ExcelCls As String = ""
        Dim booSplit As Boolean
        Dim dPeriod As Date

        Try
            ls_sql = " select distinct TOP 1 Consignee = isnull(MA.ConsigneeCode,''), attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), period = PME.Period, DOM.SuratJalanNo, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') + ' ' + isnull(MS.City,'') + ' ' + isnull(MS.Postalcode,'') SUPPAddress,  " & vbCrLf & _
                     " PME.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  + ' ' + isnull(MF.City,'') + ' ' + isnull(MF.PostalCode,'') as FWDAddress, ETDVendor1 as ETDVendor, ETDPort1 as ETDPort, ETAPort1 as ETAPort, ETAFactory1 as ETAFactory, " & vbCrLf & _
                     " PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'')  + ' ' + isnull(MA.City,'') + ' ' + isnull(MA.PostalCode,'') as AFFAddress, ISNULL(DOM.MovingList,0) MovingList, ConsigneName = isnull(MA.ConsigneeName,''), ConsigneeAdd = Rtrim(Isnull(MA.ConsigneeAddress,'')), isnull(DOM.ExcelCls,0) ExcelCls, ISNULL(DOM.SplitReffPONo, '') SplitReffPONo, ISNULL(DOM.CommercialCls,'1') CommercialCls " & vbCrLf & _
                     " from DOSUpplier_Master_Export DOM " & vbCrLf & _
                     " INNER JOIN PO_Master_Export PME ON DOM.PONo = PME.PONo and DOM.SupplierID = PME.SupplierID and DOM.AffiliateID = PME.AffiliateID AND DOM.OrderNo = PME.OrderNo1" & vbCrLf & _
                     " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                     " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID " & vbCrLf & _
                     " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                     " where isnull(DOM.ExcelCls,0) IN ('1', '3') and isnull(DOM.PONO,'') <> '' and DOM.SuratJalanno <> '' "

            'MdlConn.ReadConnection()
            ds = cls.uf_GetDataSet(ls_sql)

            Dim ls_file As String

            For i_loop = 0 To ds.Tables(0).Rows.Count - 1
                pNamaFile1 = ""
                pNamaFile2 = ""

                '=======================================Delivery Instruction To Forwarder===========================================                    
                Dim fi3 As New FileInfo(Trim(txtAttachmentDOM.Text) & "\TEMPLATE DELIVERY NOTE FORWARDER_EXPORT.xlsm")

                If Not fi3.Exists Then
                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Note STOPPED, because File Excel isn't Found " & vbCrLf & _
                                    rtbProcess.Text
                    Exit Sub
                End If

                ls_ExcelCls = Trim(ds.Tables(0).Rows(i_loop)("ExcelCls"))
                If ls_ExcelCls = "3" Then
                    booSplit = True

                    ls_sql = " select distinct Consignee = isnull(MA.ConsigneeCode,''), attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), period = PME.Period, DOM.SuratJalanNo, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') + ' ' + isnull(MS.City,'') + ' ' + isnull(MS.Postalcode,'') SUPPAddress,  " & vbCrLf & _
                             " PME.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  + ' ' + isnull(MF.City,'') + ' ' + isnull(MF.PostalCode,'') as FWDAddress, ETDVendor1 as ETDVendor, ETDPort1 as ETDPort, ETAPort1 as ETAPort, ETAFactory1 as ETAFactory, " & vbCrLf & _
                             " PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'')  + ' ' + isnull(MA.City,'') + ' ' + isnull(MA.PostalCode,'') as AFFAddress, ISNULL(DOM.MovingList,0) MovingList, ConsigneName = isnull(MA.ConsigneeName,''), ConsigneeAdd = Rtrim(Isnull(MA.ConsigneeAddress,'')), isnull(DOM.ExcelCls,0) ExcelCls, ISNULL(DOM.SplitReffPONo, '') SplitReffPONo from  " & vbCrLf & _
                             " DOSUpplier_Master_Export DOM LEFT JOIN PO_Master_Export PME  " & vbCrLf & _
                             " ON  DOM.PONo = PME.PONo and DOM.SupplierID = PME.SupplierID and DOM.AffiliateID = PME.AffiliateID AND DOM.OrderNo = PME.OrderNo1" & vbCrLf & _
                             " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                             " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID " & vbCrLf & _
                             " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                             " WHERE DOM.PONo = '" & Trim(ds.Tables(0).Rows(i_loop)("PONo")) & "' " & vbCrLf & _
                             " AND DOM.AffiliateID = '" & Trim(ds.Tables(0).Rows(i_loop)("AffiliateID")) & "' " & vbCrLf & _
                             " AND DOM.SupplierID = '" & Trim(ds.Tables(0).Rows(i_loop)("SupplierID")) & "' " & vbCrLf & _
                             " AND DOM.OrderNo = '" & Trim(ds.Tables(0).Rows(i_loop)("SplitReffPONo")) & "' "

                    dsSplit = cls.uf_GetDataSet(ls_sql)

                    ls_SJ = Trim(dsSplit.Tables(0).Rows(0)("SuratJalanNo"))
                    ls_Supplier = Trim(dsSplit.Tables(0).Rows(0)("supplierID"))
                    ls_supplierName = Trim(dsSplit.Tables(0).Rows(0)("suppliername"))
                    ls_supplierAdd = Trim(dsSplit.Tables(0).Rows(0)("SuppAddress"))
                    ls_delivery = Trim(dsSplit.Tables(0).Rows(0)("ForwarderID"))
                    ls_DeliveryName = Trim(dsSplit.Tables(0).Rows(0)("ForwarderName"))
                    ls_deliveryAdd = Trim(dsSplit.Tables(0).Rows(0)("FWDAddress"))
                    ls_orderNo = Trim(dsSplit.Tables(0).Rows(0)("PONo"))
                    ls_orderNo1 = Trim(dsSplit.Tables(0).Rows(0)("OrderNo"))
                    ls_ETDV = Format((dsSplit.Tables(0).Rows(0)("ETDVendor")), "yyyy-MM-dd")
                    ls_ETDP = Format((dsSplit.Tables(0).Rows(0)("ETDPort")), "yyyy-MM-dd")
                    ls_ETAP = Format((dsSplit.Tables(0).Rows(0)("ETAPort")), "yyyy-MM-dd")
                    ls_ETAF = Format((dsSplit.Tables(0).Rows(0)("ETAFactory")), "yyyy-MM-dd")
                    ls_Aff = Trim(dsSplit.Tables(0).Rows(0)("AFF"))
                    ls_AFFName = Trim(dsSplit.Tables(0).Rows(0)("AFFName"))
                    ls_AffADD = Trim(dsSplit.Tables(0).Rows(0)("AFFAddress"))
                    ls_Attn = Trim(dsSplit.Tables(0).Rows(0)("attn"))
                    ls_telp = Trim(dsSplit.Tables(0).Rows(0)("telp"))
                    ls_Consignee = Trim(dsSplit.Tables(0).Rows(0)("Consignee"))
                    ls_ConsName = Trim(dsSplit.Tables(0).Rows(0)("Consignename"))
                    ls_ConsigneAdd = Trim(dsSplit.Tables(0).Rows(0)("ConsigneeAdd"))
                    dPeriod = ds.Tables(0).Rows(0)("Period")
                Else
Split:
                    booSplit = False

                    ls_SJ = Trim(ds.Tables(0).Rows(i_loop)("SuratJalanNo"))
                    ls_Supplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))
                    ls_supplierName = Trim(ds.Tables(0).Rows(i_loop)("suppliername"))
                    ls_supplierAdd = Trim(ds.Tables(0).Rows(i_loop)("SuppAddress"))
                    ls_delivery = Trim(ds.Tables(0).Rows(i_loop)("ForwarderID"))
                    ls_DeliveryName = Trim(ds.Tables(0).Rows(i_loop)("ForwarderName"))
                    ls_deliveryAdd = Trim(ds.Tables(0).Rows(i_loop)("FWDAddress"))
                    ls_orderNo = Trim(ds.Tables(0).Rows(i_loop)("PONo"))
                    ls_orderNo1 = Trim(ds.Tables(0).Rows(i_loop)("OrderNo"))
                    ls_ETDV = Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "yyyy-MM-dd")
                    ls_ETDP = Format((ds.Tables(0).Rows(i_loop)("ETDPort")), "yyyy-MM-dd")
                    ls_ETAP = Format((ds.Tables(0).Rows(i_loop)("ETAPort")), "yyyy-MM-dd")
                    ls_ETAF = Format((ds.Tables(0).Rows(i_loop)("ETAFactory")), "yyyy-MM-dd")
                    ls_Aff = Trim(ds.Tables(0).Rows(i_loop)("AFF"))
                    ls_AFFName = Trim(ds.Tables(0).Rows(i_loop)("AFFName"))
                    ls_AffADD = Trim(ds.Tables(0).Rows(i_loop)("AFFAddress"))
                    ls_Attn = Trim(ds.Tables(0).Rows(i_loop)("attn"))
                    ls_telp = Trim(ds.Tables(0).Rows(i_loop)("telp"))
                    ls_Consignee = Trim(ds.Tables(0).Rows(i_loop)("Consignee"))
                    ls_ConsName = Trim(ds.Tables(0).Rows(i_loop)("Consignename"))
                    ls_ConsigneAdd = Trim(ds.Tables(0).Rows(i_loop)("ConsigneeAdd"))
                    dPeriod = ds.Tables(0).Rows(i_loop)("Period")
                End If

                dsDetailDelivery = BindDataDeliveryInstruction(ls_SJ, ls_orderNo, ls_orderNo1, ls_Aff, ls_Supplier)
                If ds.Tables(0).Rows(i_loop)("CommercialCls") = "1" Then
                    ls_commercial = "YES"
                Else
                    ls_commercial = "NO"
                End If

                If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                    Dim dsEmail As New DataSet
                    dsEmail = EmailToEmailCCKanban_Export(ls_Aff, ls_Supplier)
                    '1 CC Affiliate
                    '2 CC PASI
                    '3 CC & TO Supplier
                    For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(i)("KanbanCC")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("KanbanCC")
                        End If

                        If dsEmail.Tables(0).Rows(i)("KanbanTO") <> "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(i)("KanbanTO")
                        End If
                    Next

                    Dim k As Long
                    Dim dsAffiliate As New DataSet
                    dsAffiliate = Affiliate(Trim(ls_Aff))

                    Dim dsSupplier As New DataSet
                    dsSupplier = Supplier(Trim(ls_Supplier))

                    Dim status As Boolean
                    status = True

                    If dsDetailDelivery.Tables(0).Rows.Count = 0 Then
                        status = False
                    Else
                        status = True
                    End If

                    If status = True Then
                        NewFileCopy = Trim(txtAttachmentDOM.Text) & "\TEMPLATE DELIVERY NOTE FORWARDER_EXPORT.xlsm"
                        ls_file = NewFileCopy
                        ExcelBook = xlApp.Workbooks.Open(ls_file)
                        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                        ExcelSheet.Range("H2").Value = receiptEmail.Trim
                        ExcelSheet.Range("H3").Value = ls_Consignee.Trim
                        ExcelSheet.Range("S1").Value = ls_Aff.Trim
                        ExcelSheet.Range("S1").Font.Color = Color.White
                        ExcelSheet.Range("H4").Value = ls_delivery.Trim
                        ExcelSheet.Range("H5").Value = ls_Supplier.Trim

                        ExcelSheet.Range("I11:X11").Merge()
                        ExcelSheet.Range("I11:X11").Value = ls_supplierName.Trim
                        ExcelSheet.Range("I12:X15").Merge()
                        ExcelSheet.Range("I12:X15").Value = ls_supplierAdd.Trim
                        ExcelSheet.Range("I19:X19").Merge()
                        ExcelSheet.Range("I19:X19").Value = ls_DeliveryName.Trim
                        ExcelSheet.Range("I20:X22").Merge()
                        ExcelSheet.Range("I20:X22").Value = ls_deliveryAdd.Trim
                        ExcelSheet.Range("I28:P28").Merge()
                        ExcelSheet.Range("I28:P28").Value = ls_SJ.Trim
                        ExcelSheet.Range("I23:X23").Value = "ATTN : " & Trim(ls_Attn) & "     TELP : " & Trim(ls_telp)
                        ExcelSheet.Range("AE17:AE17").Value = ls_commercial.Trim
                        ExcelSheet.Range("AE13:AE13").Value = ls_orderNo1.Trim
                        ExcelSheet.Range("AE15:AE15").Value = ls_orderNo.Trim

                        ExcelSheet.Range("AE11:AE11").Value = Format(dPeriod, "yyyy-MM-dd")

                        ExcelSheet.Range("AP11:AT11").Merge()
                        ExcelSheet.Range("AP11:AT11").Value = ls_ETDV
                        ExcelSheet.Range("AP13:AT13").Merge()
                        ExcelSheet.Range("AP13:AT13").Value = ls_ETDP
                        ExcelSheet.Range("AP15:AT15").Merge()
                        ExcelSheet.Range("AP15:AT15").Value = ls_ETAP
                        ExcelSheet.Range("AP17:AT17").Merge()
                        ExcelSheet.Range("AP17:AT17").Value = ls_ETAF

                        ExcelSheet.Range("AE19:AT19").Merge()
                        ExcelSheet.Range("AE19:AT19").Value = ls_ConsName.Trim

                        ExcelSheet.Range("AE20:AT20").Merge()
                        ExcelSheet.Range("AE20:AT20").Value = ls_ConsigneAdd.Trim
                        k = 0

                        For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                            'For i = 0 To 3
                            k = k
                            Dim newKanbanNo As String = ""

                            If RecExNew = False Then
                                ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Merge()
                                ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Merge()
                                ExcelSheet.Range("i" & k + 34 & ": Q" & k + 34).Merge()
                                ExcelSheet.Range("R" & k + 34 & ": U" & k + 34).Merge()
                                ExcelSheet.Range("V" & k + 34 & ": W" & k + 34).Merge()
                                ExcelSheet.Range("X" & k + 34 & ": Y" & k + 34).Merge()
                                ExcelSheet.Range("Z" & k + 34 & ": AC" & k + 34).Merge()
                                ExcelSheet.Range("AD" & k + 34 & ": AG" & k + 34).Merge()
                                ExcelSheet.Range("AH" & k + 34 & ": AK" & k + 34).Merge()
                                ExcelSheet.Range("AL" & k + 34 & ": AO" & k + 34).Merge()

                                ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Value = k + 1
                                ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partno")).Trim
                                ExcelSheet.Range("i" & k + 34 & ": Q" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partname"))
                                ExcelSheet.Range("R" & k + 34 & ": U" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo1")) + "-" + Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo2"))
                                ExcelSheet.Range("V" & k + 34 & ": W" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("uom")
                                ExcelSheet.Range("X" & k + 34 & ": Y" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("qtybox")
                                ExcelSheet.Range("Z" & k + 34 & ": AC" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("qty")

                                ExcelSheet.Range("AD" & k + 34 & ": AP" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("DeliveryQty")
                                ExcelSheet.Range("AH" & k + 34 & ": AK" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("BoxQty")
                                ExcelSheet.Range("AL" & k + 34 & ": AO" & k + 34).Value = 0
                            Else
                                ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Merge()
                                ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Merge()
                                ExcelSheet.Range("i" & k + 34 & ": Q" & k + 34).Merge()
                                ExcelSheet.Range("R" & k + 34 & ": U" & k + 34).Merge()
                                ExcelSheet.Range("V" & k + 34 & ": Y" & k + 34).Merge()
                                ExcelSheet.Range("Z" & k + 34 & ": AA" & k + 34).Merge()
                                ExcelSheet.Range("AB" & k + 34 & ": AC" & k + 34).Merge()
                                ExcelSheet.Range("AD" & k + 34 & ": AG" & k + 34).Merge()
                                ExcelSheet.Range("AH" & k + 34 & ": AK" & k + 34).Merge()
                                ExcelSheet.Range("AL" & k + 34 & ": AO" & k + 34).Merge()
                                ExcelSheet.Range("AP" & k + 34 & ": AS" & k + 34).Merge()

                                ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Value = k + 1
                                ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partno")).Trim
                                ExcelSheet.Range("i" & k + 34 & ": Q" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partname"))
                                ExcelSheet.Range("R" & k + 34 & ": U" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo1") & "").Trim
                                ExcelSheet.Range("V" & k + 34 & ": Y" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo2") & "").Trim
                                ExcelSheet.Range("R" & k + 34).Interior.Color = Color.Yellow
                                ExcelSheet.Range("V" & k + 34).Interior.Color = Color.Yellow
                                ExcelSheet.Range("Z" & k + 34 & ": AA" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("uom")
                                ExcelSheet.Range("AB" & k + 34 & ": AC" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("qtybox")
                                ExcelSheet.Range("AD" & k + 34 & ": AG" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("qty")
                                ExcelSheet.Range("AD" & k + 34 & ": AG" & k + 34).NumberFormat = "#,##0"

                                ExcelSheet.Range("AH" & k + 34 & ": AK" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("DeliveryQty")
                                ExcelSheet.Range("AH" & k + 34 & ": AK" & k + 34).NumberFormat = "#,##0"
                                ExcelSheet.Range("AL" & k + 34 & ": AO" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("BoxQty")
                                ExcelSheet.Range("AL" & k + 34 & ": AO" & k + 34).NumberFormat = "#,##0"
                                ExcelSheet.Range("AP" & k + 34 & ": AS" & k + 34).Value = 0
                                ExcelSheet.Range("AP" & k + 34).Font.Color = Color.Black
                            End If

                            k = k + 1
                        Next
                        ExcelSheet.Range("B35").Interior.Color = Color.White
                        ExcelSheet.Range("B35").Font.Color = Color.Black
                        ExcelSheet.Range("B" & k + 34).Value = "E"
                        ExcelSheet.Range("B" & k + 34).Interior.Color = Color.Black
                        ExcelSheet.Range("B" & k + 34).Font.Color = Color.White

                        If RecExNew = False Then
                            ExcelSheet.Range("AH34" & ": AO" & k + 33).Interior.Color = Color.Yellow
                            DrawAllBorders(ExcelSheet.Range("B34" & ": AO" & k + 33))
                        Else
                            ExcelSheet.Range("AL34" & ": AS" & k + 33).Interior.Color = Color.Yellow
                            DrawAllBorders(ExcelSheet.Range("B34" & ": AS" & k + 33))
                        End If

                        'Save ke Local
                        xlApp.DisplayAlerts = False

                        If pNamaFile1 = "" Then
                            If ls_orderNo.Trim = ls_orderNo1.Trim Then
                                ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & ".xlsm")
                                pNamaFile1 = "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & ".xlsm"
                            Else
                                ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNo1) & ")" & ".xlsm")
                                pNamaFile1 = "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNo1) & ")" & ".xlsm"
                            End If
                        Else
                            If ls_orderNo.Trim = ls_orderNo1.Trim Then
                                ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & ".xlsm")
                                pNamaFile2 = "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & ".xlsm"
                            Else
                                ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNo1) & ")" & ".xlsm")
                                pNamaFile2 = "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNo1) & ")" & ".xlsm"
                            End If
                        End If

                        xlApp.Workbooks.Close()
                        xlApp.Quit()

                        If booSplit Then GoTo Split
                    End If
                    '---------------------------------------excel---------------------------------------'
                    If pNamaFile1 = "" Then GoTo keluar
                    If sendEmailtoSupllierDeliveryConfirmation("RECEIVING", pNamaFile1, pNamaFile2, Trim(ls_orderNo), ls_delivery, Trim(ls_Aff), Trim(ls_orderNo1), Trim(ls_Supplier), Trim(ls_SJ)) = False Then GoTo keluar
                    Call UpdateStatusDOExport(ls_Aff, ls_Supplier, ls_orderNo, ls_SJ, ls_orderNo1)
                End If
                '=======================================Delivery Instruction To Forwarder===========================================

keluar:
                xlApp.Workbooks.Close()
                xlApp.Quit()
            Next
            Exit Sub

        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery to Forwarder STOPPED, because " & Err.Description & " " & vbCrLf & _
                                    rtbProcess.Text
            xlApp.Workbooks.Close()
            xlApp.Quit()
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Function BindDataDeliveryInstruction(ByVal pSJ As String, ByVal PONO As String, ByVal ls_orderNo As String, ByVal Aff As String, ByVal Supp As String)
        Dim ls_sql As String
        'MdlConn.ReadConnection()

        If statusPartial = False Then
            ls_sql = " select  distinct " & vbCrLf & _
                      " orderno = POD.PONo, " & vbCrLf & _
                      " Partno = POD.PartNo, " & vbCrLf & _
                      " Partname = MP.PartName, " & vbCrLf & _
                      " labelno1 = Rtrim(PL1.LabelNo), " & vbCrLf & _
                      " labelno2 = Rtrim(PL2.LabelNo), " & vbCrLf & _
                      " uom = MU.Description, " & vbCrLf & _
                      " qtybox = ISNULL(POD.POQtyBox,MPM.QtyBox), " & vbCrLf & _
                      " qty = convert(char,DOD.DOQty), " & vbCrLf & _
                      " remaining = convert(char,POD.Week1), " & vbCrLf & _
                      " DeliveryQty = convert(char,DOD.DOQty), " & vbCrLf & _
                      " boxqty = Ceiling(DOD.DOQty / ISNULL(POD.POQtyBox,MPM.QtyBox)), "

            ls_sql = ls_sql + " weight = NetWeight, " & vbCrLf & _
                              " barcode = convert(char(25),'') + Convert(char(20),POD.AFfiliateID) + convert(char(20), POD.Pono) + Convert(char(25), POD.PartNo) +  convert(char,DOD.DOQty) " & vbCrLf & _
                              " From PO_Detail_Export POD " & vbCrLf & _
                              " INNER JOIN PO_Master_Export POM " & vbCrLf & _
                              " ON POM.Pono = POD.PONo and POM.OrderNo1 = POD.OrderNo1 " & vbCrLf & _
                              " and POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " And POM.SupplierID = POD.SupplierID "

            ls_sql = ls_sql + " LEFT JOIN (select POno, AffiliateID, SupplierID, PartNo, min(labelNo) as labelno from PrintLabelExport group by POno, AffiliateID, SupplierID, PartNo)PL1 " & vbCrLf & _
                              " ON PL1.PONo = POD.PONo " & vbCrLf & _
                              " and PL1.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " AND PL1.SupplierID = POD.SupplierID " & vbCrLf & _
                              " AND PL1.PartNO = POD.PartNo " & vbCrLf & _
                              " LEFT JOIN (select POno, AffiliateID, SupplierID, PartNo, Max(labelNo) as labelno from PrintLabelExport group by POno, AffiliateID, SupplierID, PartNo)PL2 " & vbCrLf & _
                              " ON PL2.PONo = POD.PONo " & vbCrLf & _
                              " and PL2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " AND PL2.SupplierID = POD.SupplierID " & vbCrLf & _
                              " AND PL2.PartNO = POD.PartNo " & vbCrLf & _
                              " INNER JOIN MS_Parts MP ON MP.PartNo = POD.partNo " & vbCrLf & _
                              " LEFT JOIN MS_PartMapping MPM On MPM.PartNo = POD.PartNo " & vbCrLf & _
                              " AND MPM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " AND MPM.SupplierID = POD.SupplierID " & vbCrLf

            ls_sql = ls_sql + " INNER JOIN ms_unitcls MU ON MU.UnitCls = MP.Unitcls " & vbCrLf & _
                              " LEFT JOIN DOSupplier_Master_Export DOM ON DOM.SupplierID = POM.SupplierID and DOM.AffiliateID = POM.AffiliateID" & vbCrLf & _
                              " AND DOM.PONo = POM.PONo and DOM.ORderNo = POM.OrderNo1 " & vbCrLf & _
                              " INNER JOIN DOSupplier_Detail_Export DOD ON DOD.SuratJalanNo = DOM.SuratJalanNo and DOD.AffiliateID = DOM.AffiliateID " & vbCrLf & _
                              "and DOD.SupplierID = DOM.SupplierID and DOD.PONo = DOM.POno and DOD.OrderNo = DOM.OrderNo and DOD.PartNo = POD.PartNo " & vbCrLf & _
                              " Where DOM.SuratJalanno = '" & Trim(pSJ) & "'" & vbCrLf & _
                              " AND POD.OrderNo1 = '" & Trim(ls_orderNo) & "'" & vbCrLf & _
                              " AND POD.AffiliateID = '" & Trim(Aff) & "'" & vbCrLf & _
                              " AND POD.SupplierID = '" & Trim(Supp) & "'"
        Else
            ls_sql = "  select  distinct  " & vbCrLf & _
                  "  orderno = POD.PONo,  " & vbCrLf & _
                  "  Partno = POD.PartNo,  " & vbCrLf & _
                  "  Partname = MP.PartName,  " & vbCrLf & _
                  "  labelno1 = Rtrim(PL1.LabelNo),  " & vbCrLf & _
                  "  labelno2 = Rtrim(PL2.LabelNo),  " & vbCrLf & _
                  "  uom = MU.Description,  " & vbCrLf & _
                  "  qtybox = ISNULL(POD.POQtyBox,MPM.QtyBox),  " & vbCrLf & _
                  "  qty = convert(char,DOD.DOQty),  " & vbCrLf & _
                  "  remaining = convert(char,POD.Week1),  " & vbCrLf & _
                  "  DeliveryQty = convert(char,DOD.DOQty),  " & vbCrLf & _
                  "  boxqty = Ceiling(DOD.DOQty / ISNULL(POD.POQtyBox,MPM.QtyBox)),  weight = NetWeight,  "

            ls_sql = ls_sql + "  barcode = convert(char(25),'') + Convert(char(20),POD.AFfiliateID) + convert(char(20), POD.Pono) + Convert(char(25), POD.PartNo) +  convert(char,DOD.DOQty)  " & vbCrLf & _
                              "  From PO_Detail_Export POD  " & vbCrLf & _
                              "  INNER JOIN PO_Master_Export POM  " & vbCrLf & _
                              "  ON POM.Pono = POD.PONo and POM.OrderNo1 = POD.OrderNo1  " & vbCrLf & _
                              "  and POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                              "  And POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "  LEFT JOIN DOSupplier_Master_Export DOM ON DOM.SupplierID = POM.SupplierID and DOM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "  AND DOM.PONo = POM.PONo and DOM.ORderNo = POM.OrderNo1  " & vbCrLf & _
                              "  INNER JOIN DOSupplier_Detail_Export DOD ON DOD.SuratJalanNo = DOM.SuratJalanNo and DOD.AffiliateID = DOM.AffiliateID  " & vbCrLf & _
                              " and DOD.SupplierID = DOM.SupplierID and DOD.PONo = DOM.POno and DOD.OrderNo = DOM.OrderNo and DOD.PartNo = POD.PartNo  " & vbCrLf & _
                              "  LEFT JOIN (select OrderNo,SuratJalanno,POno, AffiliateID, SupplierID, PartNo, min(BoxNo) as labelno, SeqNo from DOSupplier_DetailBox_Export group by OrderNo,SuratJalanno,POno, AffiliateID, SupplierID, PartNo, SeqNo)PL1  "

            ls_sql = ls_sql + "  ON PL1.PONo = DOD.PONo  " & vbCrLf & _
                              "  and PL1.AffiliateID = DOD.AffiliateID  " & vbCrLf & _
                              "  AND PL1.SupplierID = DOD.SupplierID  " & vbCrLf & _
                              "  AND PL1.PartNO = DOD.PartNo  " & vbCrLf & _
                              "  AND PL1.SuratJalanno = DOD.SuratJalanno " & vbCrLf & _
                              "  AND PL1.OrderNo = DOD.OrderNo " & vbCrLf & _
                              "  AND PL1.SeqNo = DOD.SeqNo " & vbCrLf & _
                              "  LEFT JOIN (select OrderNo,SuratJalanno,POno, AffiliateID, SupplierID, PartNo, Max(BoxNo) as labelno, SeqNo from DOSupplier_DetailBox_Export group by OrderNo,SuratJalanno,POno, AffiliateID, SupplierID, PartNo, SeqNo)PL2  " & vbCrLf & _
                              "  ON PL2.PONo = DOD.PONo  " & vbCrLf & _
                              "  and PL2.AffiliateID = DOD.AffiliateID  " & vbCrLf & _
                              "  AND PL2.SupplierID = DOD.SupplierID  " & vbCrLf & _
                              "  AND PL2.PartNO = DOD.PartNo  " & vbCrLf & _
                              "  AND PL2.SuratJalanno = DOD.SuratJalanno " & vbCrLf & _
                              "  AND PL2.OrderNo = DOD.OrderNo " & vbCrLf & _
                              "  AND PL2.SeqNo = DOD.SeqNo " & vbCrLf

            ls_sql = ls_sql + "  INNER JOIN MS_Parts MP ON MP.PartNo = POD.partNo  " & vbCrLf & _
                              "  LEFT JOIN MS_PartMapping MPM On MPM.PartNo = POD.PartNo  " & vbCrLf & _
                              "  AND MPM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                              "  AND MPM.SupplierID = POD.SupplierID  " & vbCrLf & _
                              "  INNER JOIN ms_unitcls MU ON MU.UnitCls = MP.Unitcls  " & vbCrLf & _
                              " Where (Pl1.labelno is not null and pl2.labelno is not null) and DOM.SuratJalanno = '" & Trim(pSJ) & "'" & vbCrLf & _
                              " AND POD.OrderNo1 = '" & Trim(ls_orderNo) & "'" & vbCrLf & _
                              " AND POD.AffiliateID = '" & Trim(Aff) & "'" & vbCrLf & _
                              " AND POD.SupplierID = '" & Trim(Supp) & "'"
        End If

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Private Sub UpdateStatusDOExport(ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pSJ As String, ByVal pOrderNo As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " update dbo.DOSupplier_Master_Export set excelcls = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND SuratJalanno = '" & pSJ & "'" & vbCrLf & _
                         " AND PONo = '" & pPoNo & "'" & vbCrLf & _
                         " AND OrderNo = '" & pOrderNo & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub

    Private Sub pGetExcelTally()
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        'copy file from server to local
        Dim NewFileCopy As String

        'MdlConn.ReadConnection()
        Dim ls_SQL As String = ""
        Dim ds As New DataSet
        Dim dsDetail As New DataSet
        Dim sh_Affiliate As String = "", sh_ForwarderID As String, sh_Shippingno As String, sh_Consignee As String
        Dim xlApp = New Excel.Application

        Try
            ls_SQL = " Select distinct TotalCtn = SUM(SHD.BoxQty), Consignee = isnull(ConsigneeCode,''), SHM.AffiliateID, SHM.ForwarderID, SHM.ShippingInstructionNo, ETDPort = Convert(Char(12), convert(Datetime, isnull(SHM.ETDPort,'')),106), DestinationPort From ShippingInstruction_Master SHM " & vbCrLf & _
                  " LEFT JOIN ShippingInstruction_Detail SHD ON SHM.AffiliateID = SHD.AffiliateID " & vbCrLf & _
                  " AND SHM.ForwarderID = SHD.ForwarderID " & vbCrLf & _
                  " AND SHM.ShippingInstructionNo = SHD.ShippingInstructionNo " & vbCrLf & _
                  " LEFT JOIN ReceiveForwarder_Detail RD ON RD.AffiliateID = SHM.AffiliateID " & vbCrLf & _
                  " AND SHD.SupplierID = RD.SupplierID " & vbCrLf & _
                  " AND RD.PartNo = SHD.PartNo " & vbCrLf & _
                  " and RD.OrderNo = SHD.OrderNo  " & vbCrLf & _
                  " AND RD.SuratJalanNo = SHD.SuratJalanno " & vbCrLf & _
                  " AND SHD.SupplierID = RD.SupplierID " & vbCrLf & _
                  " LEFT JOIN PO_Master_Export POM ON POM.PONo = RD.PONO and POM.OrderNo1 = RD.OrderNo " & vbCrLf & _
                  " AND POM.AffiliateID = RD.AffiliateID and RD.SupplierID = POM.SupplierID " & vbCrLf & _
                  " LEFT JOIN MS_Affiliate MA ON POM.AffiliateID = MA.AffiliateID "

            ls_SQL = ls_SQL + "  where isnull(SHM.ExcelCls,'') = '1'  "
            ls_SQL = ls_SQL + " Group by ConsigneeCode, SHM.AffiliateID, SHM.ForwarderID, SHM.ShippingInstructionNo, SHM.ETDPort, DestinationPort  "


            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pDate = Now
                    sh_Affiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    sh_ForwarderID = Trim(ds.Tables(0).Rows(xi)("ForwarderID"))
                    sh_Consignee = Trim(ds.Tables(0).Rows(xi)("Consignee"))
                    sh_Shippingno = Trim(ds.Tables(0).Rows(xi)("ShippingInstructionNo"))

                    dsDetail = getTallyDetail(sh_Shippingno, sh_Affiliate, sh_ForwarderID)

                    'Create Excel File
                    Dim fi As New FileInfo(Trim(txtAttachmentDOM.Text) & "\Template Tally.xlsm") 'File dari Local
                    If Not fi.Exists Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because file Excel isn't Found" & vbCrLf & _
                                        rtbProcess.Text
                        Exit Sub
                    End If

                    NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Template Tally.xlsm"
                    Dim ls_file As String = NewFileCopy
                    ExcelBook = xlApp.Workbooks.Open(ls_file)
                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                    Dim dsEmail As New DataSet
                    dsEmail = EmailSH(sh_Affiliate, "PASI", sh_ForwarderID)
                    For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(i)("SupplierDeliverycc")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("SupplierDeliverycc")
                        End If
                        If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                            fromEmail = dsEmail.Tables(0).Rows(i)("toEmail")
                        End If
                        If receiptEmail = "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(i)("SupplierDeliveryTo")
                        Else
                            receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(i)("SupplierDeliveryTo")
                        End If
                    Next

                    receiptCCEmail = Replace(receiptCCEmail, ",", ";")
                    receiptEmail = Replace(receiptEmail, ",", ";")

                    ExcelSheet.Range("H1").Value = "TALLY"
                    ExcelSheet.Range("H2").Value = fromEmail
                    ExcelSheet.Range("H3").Value = Trim(sh_Consignee)
                    ExcelSheet.Range("H4").Value = Trim(sh_ForwarderID)
                    ExcelSheet.Range("I8").Value = Trim(sh_Shippingno)
                    ExcelSheet.Range("AA12").Value = ds.Tables(0).Rows(xi)("ETDPort")
                    ExcelSheet.Range("AA16").Value = ds.Tables(0).Rows(xi)("DestinationPort")
                    ExcelSheet.Range("I18").Value = ds.Tables(0).Rows(xi)("TotalCtn")
                    ExcelSheet.Range("S1").Value = Trim(sh_Affiliate)
                    ExcelSheet.Range("S1").Font.Color = Color.White

                    ExcelSheet.Range("Y2").Value = ""

                    If dsDetail.Tables(0).Rows.Count > 0 Then
                        For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                            'Header
                            ExcelSheet.Range("B" & i + 23 & ": C" & i + 23).Merge()
                            ExcelSheet.Range("B" & i + 23 & ": C" & i + 23).Value = i + 1
                            ExcelSheet.Range("B" & i + 23 & ": C" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                            ExcelSheet.Range("D" & i + 23 & ": H" & i + 23).Merge()
                            ExcelSheet.Range("D" & i + 23 & ": H" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("D" & i + 23 & ": H" & i + 23).Interior.Color = ColorYellow

                            ExcelSheet.Range("I" & i + 23 & ": N" & i + 23).Merge()
                            ExcelSheet.Range("I" & i + 23 & ": N" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("OrderNo"))
                            ExcelSheet.Range("I" & i + 23 & ": N" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("O" & i + 23 & ": U" & i + 23).Merge()
                            ExcelSheet.Range("O" & i + 23 & ": U" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                            ExcelSheet.Range("O" & i + 23 & ": U" & i + 23).Merge()

                            ExcelSheet.Range("V" & i + 23 & ": AD" & i + 23).Merge()
                            ExcelSheet.Range("V" & i + 23 & ": AD" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName"))
                            ExcelSheet.Range("V" & i + 23 & ": AD" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("AE" & i + 23 & ": AJ" & i + 23).Merge()
                            ExcelSheet.Range("AE" & i + 23 & ": AJ" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("CaseNo1"))
                            ExcelSheet.Range("AE" & i + 23 & ": AJ" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("AK" & i + 23 & ": AP" & i + 23).Merge()
                            ExcelSheet.Range("AK" & i + 23 & ": AP" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("CaseNo2"))
                            ExcelSheet.Range("AK" & i + 23 & ": AP" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("AQ" & i + 23 & ": AS" & i + 23).Merge()
                            'ExcelSheet.Range("AK" & i + 23 & ": AM" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("Length"))
                            ExcelSheet.Range("AQ" & i + 23 & ": AS" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AQ" & i + 23 & ": AS" & i + 23).NumberFormat = "#,##0.00"

                            ExcelSheet.Range("AT" & i + 23 & ": AV" & i + 23).Merge()
                            'ExcelSheet.Range("AN" & i + 23 & ": AP" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("Width"))
                            ExcelSheet.Range("AT" & i + 23 & ": AV" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AT" & i + 23 & ": AV" & i + 23).NumberFormat = "#,##0.00"

                            ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).Merge()
                            'ExcelSheet.Range("AQ" & i + 23 & ": AS" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("Height"))
                            ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).NumberFormat = "#,##0.00"

                            ExcelSheet.Range("D" & i + 23 & ": H" & i + 23).Merge()
                            ExcelSheet.Range("AT" & i + 23 & ": AV" & i + 23).Merge()

                            ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).Merge()
                            ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).Interior.Color = ColorYellow
                            ExcelSheet.Range("BC" & i + 23 & ": BE" & i + 23).Interior.Color = ColorYellow
                            ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AE" & i + 23 & ": AY" & i + 23).Interior.Color = ColorYellow

                            ExcelSheet.Range("BC" & i + 23 & ": BE" & i + 23).Merge()
                            ExcelSheet.Range("AZ" & i + 23 & ": BB" & i + 23).Merge()
                            ExcelSheet.Range("AZ" & i + 23 & ": BB" & i + 23).FormulaR1C1 = "=RC[-9]*RC[-6]*RC[-3]" '"=((RC[-9]*RC[-6]*RC[-3])/6000)/1000"
                            ExcelSheet.Range("AZ" & i + 23 & ": BB" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AZ" & i + 23 & ": BB" & i + 23).NumberFormat = "#,##0.00"
                            '=((AK23*AN23*AQ23)/6000)/1000
                            DrawAllBorders(ExcelSheet.Range("B" & i + 23 & ": BE" & i + 23))
                            'ExcelSheet.Range("B" & i + 23 & ": AD" & i + 37).Interior.Color = RGB(217, 217, 217)
                        Next
                    End If

                    ExcelSheet.Range("B" & i + 23).Value = "E"
                    ExcelSheet.Range("B" & i + 23).Interior.Color = Color.Black
                    ExcelSheet.Range("B" & i + 23).Font.Color = Color.White
                    ExcelSheet.Range("B24").Font.Color = Color.Black
                    ExcelSheet.Range("B24").Interior.Color = Color.White


                    xlApp.DisplayAlerts = False

                    Dim temp_Filename As String = "Tally Data " & Trim(sh_Affiliate) & "-" & Trim(sh_ForwarderID) & "-" & Trim(sh_Shippingno) & ".xlsm"
                    ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\" & temp_Filename)
                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & temp_Filename & " OK. " & vbCrLf & _
                    rtbProcess.Text

                    If sendEmailTallyToForwarder(temp_Filename, sh_Shippingno, sh_Affiliate, sh_ForwarderID) = False Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & temp_Filename & " NG. " & vbCrLf & _
                        rtbProcess.Text
                        GoTo keluar
                    Else
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & temp_Filename & " OK. " & vbCrLf & _
                        rtbProcess.Text
                    End If


                    Call UpdateTallyCls(True, sh_Shippingno, sh_Affiliate, sh_ForwarderID)

                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    Thread.Sleep(500)
keluar:
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                Next
            Else
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because there is nothing PO to send " & vbCrLf & _
                    rtbProcess.Text
            End If
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Function getTallyDetail(ByVal shShippingNo As String, ByVal shAffiliate As String, ByVal shForwarder As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = " select distinct " & vbCrLf & _
                  " InvoiceNo = SHM.ShippingInstructionNo, " & vbCrLf & _
                  " OrderNo = SHD.OrderNo, " & vbCrLf & _
                  " PartNo = SHD.PartNo, " & vbCrLf & _
                  " PartName = MP.PartGroupName, " & vbCrLf & _
                  " CaseNo1 = Label1,  " & vbCrLf & _
                  " CaseNo2 = Label2,  " & vbCrLf &
                  " Length = length, " & vbCrLf & _
                  " Width = Width, " & vbCrLf & _
                  " Height = Height, " & vbCrLf & _
                  " ForwarderID = SHM.ForwarderID " & vbCrLf & _
                  " From ShippingInstruction_master SHM  "

        ls_SQL = ls_SQL + " LEFT JOIN ShippingInstruction_Detail SHD " & vbCrLf & _
                          " ON SHM.ShippingInstructionNo = SHD.ShippingInstructionNo " & vbCrLf & _
                          " AND SHM.AffiliateID = SHD.AffiliateID " & vbCrLf & _
                          " AND SHM.ForwarderID = SHD.ForwarderID " & vbCrLf & _
                          " LEFT JOIN ReceiveForwarder_DetailBox RB " & vbCrLf & _
                          " ON RB.SuratJalanNo = SHD.SuratJalanNo  " & vbCrLf & _
                          "AND RB.AffiliateID = SHD.AffiliateID " & vbCrLf & _
                          "AND RB.SupplierID = SHD.SupplierID " & vbCrLf & _
                          "AND RB.OrderNo = SHD.OrderNo " & vbCrLf & _
                          "AND RB.PartNo = SHD.PartNo " & vbCrLf & _
                          "AND StatusDefect = '0' " & vbCrLf

        ls_SQL = ls_SQL + " LEFT JOIN MS_Parts MP ON MP.Partno = SHD.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SHD.PartNo " & vbCrLf & _
                          " AND MPM.AffiliateID = SHM.AffiliateID " & vbCrLf & _
                          " AND MPM.SupplierID = SHD.SupplierID " & vbCrLf & _
                          " Where SHM.ShippingInstructionNo = '" & Trim(shShippingNo) & "' " & vbCrLf & _
                          " AND SHM.AffiliateID = '" & Trim(shAffiliate) & "' " & vbCrLf & _
                          " AND SHM.ForwarderID = '" & Trim(shForwarder) & "' "

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Private Function EmailSH(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pforwarderID As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = " select distinct 'PASI' flag,SupplierDeliveryTo,SupplierDeliveryCC,toEmail = SupplierDeliveryTo  from ms_emailPASI_Export where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                 " Union ALL " & vbCrLf & _
                 "select distinct 'FWD' flag,SupplierDeliveryTo,SupplierDeliveryCC,toEmail = SupplierDeliveryTo from ms_emailForwarder where ForwarderID = '" & Trim(pforwarderID) & "'"
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function sendEmailTallyToForwarder(ByVal pFileName As String, ByVal shshipno As String, ByVal shaffiliate As String, ByVal shforwarder As String) As Boolean
        Try
            Dim TempFilePath As String
            Dim TempFileName As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)
            TempFileName = "\" & pFileName

            Dim dsEmail As New DataSet
            dsEmail = EmailSH(shaffiliate, "PASI", shforwarder)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "FWD" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("SupplierDeliveryTo")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("SupplierDeliveryTo")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "FWD" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("SupplierDeliverycc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("SupplierDeliverycc")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            'receiptCCEmail = "pasi-opa02@pemi.co.id;edi@tos.co.id"
            'receiptEmail = "pasi-opa02@pemi.co.id;edi@tos.co.id"

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailTallyToForwarder = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailTallyToForwarder = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            '            mailMessage.Subject = "[TRIAL] TALLY DATA: " & shaffiliate.Trim & "-" & shshipno.Trim & ""
            mailMessage.Subject = "TA-" & shaffiliate.Trim & "-" & shshipno.Trim & " Tally Data "
            'TA-AFF-Invoice No. Tally Data [TRIAL]
            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If
            'Dim receiptBCCEmail As String = "pasi.purchase@gmail.com"
            'If receiptBCCEmail <> "" Then
            '    For Each recipientBCC In receiptBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If
            GetSettingEmail_Export("PO")
            ls_Body = clsNotification.GetNotification("21", "", shshipno.Trim)
            mailMessage.Body = ls_Body

            Dim filename As String = TempFilePath & TempFileName
            mailMessage.Attachments.Add(New Attachment(filename))
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)

            sendEmailTallyToForwarder = True
            'Delete the file
            'Kill(TempFilePath & TempFileName)
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO No: " & pPONo & " to Supplier SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
            Exit Function
        Catch ex As Exception
            sendEmailTallyToForwarder = False
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO No: " & pPONo & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text

        End Try

    End Function

    Private Sub UpdateTallyCls(ByVal pIsNewData As Boolean, _
                         Optional ByVal pShippno As String = "", _
                         Optional ByVal pAFF As String = "", _
                         Optional ByVal pFWD As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & vbCrLf & _
                      " SET ExcelCls='2'" & vbCrLf & _
                      " WHERE ShippingInstructionNo='" & pShippno & "'  " & vbCrLf & _
                      " AND AffiliateID='" & pAFF & "' " & vbCrLf & _
                      " AND ForwarderID='" & pFWD & "' " & vbCrLf
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
        End Try
    End Sub

    Private Sub pGetPDFTally()
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""
        Dim pFilePDF1 As String = ""
        Dim pFilePDF2 As String = ""
        Dim pFilePDF3 As String = ""
        Dim pCSV As String = ""
        Dim pFile As String = ""

        'copy file from server to local

        'MdlConn.ReadConnection()
        Dim ls_SQL As String = ""
        Dim ds As New DataSet
        Dim dsDetail As New DataSet
        Dim sh_Affiliate As String = "", sh_ForwarderID As String, sh_Shippingno As String
        'Dim xlApp = New Excel.Application

        Try
            ls_SQL = " Select distinct * From Tally_Master where isnull(TallyCls2,'') = '1' "
            'ls_SQL = " Select distinct * From Tally_Master where ShippingInstructionNo = 'EATM60021B' "

            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pDate = Now
                    sh_Affiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    sh_ForwarderID = Trim(ds.Tables(0).Rows(xi)("ForwarderID"))
                    sh_Shippingno = Trim(ds.Tables(0).Rows(xi)("ShippingInstructionNo"))

                    ''PDF SHIPPING INSTRUCTION
                    'pFilePDF1 = CreateShippingInstructionToPDF(sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    'If pFilePDF1 = "" Then
                    '    GoTo keluar
                    'End If
                    ''PDF SHIPPING INSTRUCTION

                    'PDF TALLY
                    pFilePDF2 = CreateTallyToPDF(sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    If pFilePDF2 = "" Then
                        GoTo keluar
                    End If
                    'PDF TALLY

                    ''PDF COMMERCIAL INVOICE
                    'pFilePDF3 = CreateCommercialInvoiceToPDF(sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    'If pFilePDF3 = "" Then
                    '    GoTo keluar
                    'End If
                    ''PDF COMMERCIAL INVOICE

                    ''CSV
                    'pCSV = GetCSV(sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    'If pCSV = "" Then
                    '    GoTo keluar
                    'End If
                    ''CSV

                    'If sendEmailTallyInvoceShippingToForwarder(pFilePDF1, pFilePDF2, pFilePDF3, sh_Shippingno, sh_Affiliate, sh_ForwarderID, pCSV) = False Then GoTo keluar

                    Call UpdateTallyCls2(True, sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process PDF Shipping Instruction No: " & sh_Shippingno & " SUCCESSFULL" & vbCrLf &
                             rtbProcess.Text

                    Thread.Sleep(500)
keluar:
                Next
            Else
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Shipping Instruction STOPPED, because there is nothing PO to send " & vbCrLf & _
                    rtbProcess.Text
            End If
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Shipping Instruction STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally

        End Try
    End Sub

    Private Function CreateTallyToPDF(ByVal PshippingNo As String, ByVal pAffiliate As String, ByVal pForwarder As String) As String
        Dim CrReport As New TallyData()
        Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
        Dim pFile As String

        Try
            txtMsg.Text = ""
            Cursor.Current = Cursors.WaitCursor
            Dim dsPrint As New DataSet
            dsPrint = PrintTally(PshippingNo, pAffiliate, pForwarder)

            If dsPrint.Tables(0).Rows.Count = 0 Then Exit Try

            CrReport.SetDatabaseLogon(cfg.User, cfg.Password, cfg.Server, cfg.Database)
            CrReport.SetDataSource(dsPrint.Tables(0))

            CrDiskFileDestinationOptions.DiskFileName = Trim(txtSaveAsDOM.Text) & "\Tally-" & PshippingNo & "-" & pAffiliate.Trim & "-" & pForwarder.Trim & ".pdf"
            pFile = ""
            pFile = "Tally-" & PshippingNo & "-" & pAffiliate.Trim & "-" & pForwarder.Trim & ".pdf"
            CrExportOptions = CrReport.ExportOptions

            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With

            Try
                CrReport.Export()
            Catch err As Exception
                MessageBox.Show(err.ToString())
            End Try
            'PDF
            CreateTallyToPDF = pFile
        Catch ex As Exception
            CreateTallyToPDF = ""
        Finally
            If Not CrReport Is Nothing Then
                NAR(CrReport)
                GC.Collect()
            End If
        End Try
    End Function

    Private Function PrintTally(ByVal ls_value1 As String, ByVal ls_value2 As String, ByVal ls_value3 As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = ""
        ls_SQL = " select distinct  " & vbCrLf & _
                  " 	ContainerNo = TM.ContainerNo,  " & vbCrLf & _
                  " 	SealNo = TM.SealNo,  " & vbCrLf & _
                  " 	Tare = TM.Tare,  " & vbCrLf & _
                  " 	Gross = TM.Gross,  " & vbCrLf & _
                  " 	InvoiceNo = TM.ShippingInstructionNo,  " & vbCrLf & _
                  " 	PalletNo = TD.PalletNo,  " & vbCrLf & _
                  " 	OrderNo = TD.OrderNo,  " & vbCrLf & _
                  " 	PartNo = TD.PartNo,  " & vbCrLf & _
                  " 	CaseNo = Rtrim(TD.CaseNo) + CASE WHEN Rtrim(TD.CaseNo2) = '' then '' else + '-' + Rtrim(TD.CaseNo2) END," & vbCrLf & _
                  "     jmlCTN = TD1.totalBox, " & vbCrLf & _
                  " 	Length = (SUMTally.Length),  " & vbCrLf

        ls_SQL = ls_SQL + "     Width = (SUMTally.Width),  " & vbCrLf & _
                          "     Height = (SUMTally.Height),  " & vbCrLf & _
                          "     M3 = (SUMTally.M3),  " & vbCrLf & _
                          "     WGT = (SUMTally.WeightPallet)  " & vbCrLf & _
                          "     From Tally_master TM   " & vbCrLf & _
                          "     LEFT JOIN Tally_Detail TD  " & vbCrLf & _
                          "     ON TM.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          "     AND TM.ForwarderID = TD.ForwarderID  " & vbCrLf & _
                          "     AND TM.AffiliateID = TD.AffiliateID  " & vbCrLf & _
                          "     LEFT JOIN (select ShippingInstructionNo, ForwarderID, AffiliateID, SUM(TotalBox) as totalBox from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID)TD1 " & vbCrLf

        ls_SQL = ls_SQL + " 	ON TD1.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          " 		AND TD1.ForwarderID = TD.ForwarderID  " & vbCrLf & _
                          " 		AND TD1.AffiliateID = TD.AffiliateID  " & vbCrLf & _
                          " 		--AND TD1.PalletNo = TD.PalletNo  " & vbCrLf & _
                          " 		--AND TD1.OrderNo = TD.OrderNo  " & vbCrLf & _
                          " 		--AND TD1.PartNO = TD.PartNo  " & vbCrLf & _
                          " 	LEFT JOIN (select ShippingInstructionNo, ForwarderID, AffiliateID, PalletNo, OrderNo,PartNo,  " & vbCrLf & _
                          " 				Max(CaseNo) as CaseNo1 from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID, PalletNo, OrderNo,PartNo)TD2  " & vbCrLf & _
                          " 	ON TD2.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          " 		AND TD2.ForwarderID = TD.ForwarderID  " & vbCrLf & _
                          " 		AND TD2.AffiliateID = TD.AffiliateID  " & vbCrLf

        ls_SQL = ls_SQL + " 		AND TD2.PalletNo = TD.PalletNo  " & vbCrLf & _
                          " 		AND TD2.OrderNo = TD.OrderNo  " & vbCrLf & _
                          " 		AND TD2.PartNO = TD.PartNo  " & vbCrLf & _
                          "     LEFT JOIN (select JMLCTN = Sum(JMLCTN),ShippingInstructionNo, ForwarderID, AffiliateID From( " & vbCrLf & _
                          "                     select ShippingInstructionNo, ForwarderID, AffiliateID,  " & vbCrLf & _
                          "                         Count(CaseNo) as JMLCTN from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID) x " & vbCrLf & _
                          "                     group by ShippingInstructionNo, ForwarderID, AffiliateID)TD3  " & vbCrLf & _
                          "         ON TD3.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          "         AND TD3.ForwarderID = TD.ForwarderID " & vbCrLf & _
                          "         AND TD3.AffiliateID = TD.AffiliateID " & vbCrLf & _
                          " 	LEFT JOIN (select distinct ShippingInstructionNo, ForwarderID, AffiliateID, OrderNo, palletno, weightpallet,  " & vbCrLf & _
                          " 				width, height, length, M3 from Tally_Detail) SUMTally " & vbCrLf & _
                          " 	ON SUMTally.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          " 		AND SUMTally.ForwarderID = TD.ForwarderID  " & vbCrLf & _
                          " 		AND SUMTally.AffiliateID = TD.AffiliateID  " & vbCrLf & _
                          " 		AND SUMTally.PalletNo = TD.PalletNo  " & vbCrLf & _
                          " 		AND SUMTally.OrderNo = TD.OrderNo  " & vbCrLf

        ls_SQL = ls_SQL + "  " & vbCrLf & _
                          " WHERE TM.ShippingInstructionNo = '" & Trim(ls_value1) & "' " & vbCrLf & _
                          " AND TM.AffiliateID = '" & Trim(ls_value2) & "'" & vbCrLf & _
                          " AND TM.ForwarderID = '" & Trim(ls_value3) & "' "
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Sub UpdateTallyCls2(ByVal pIsNewData As Boolean, _
                         Optional ByVal pShippno As String = "", _
                         Optional ByVal pAFF As String = "", _
                         Optional ByVal pFWD As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.Tally_Master " & vbCrLf & _
                      " SET TallyCls2='2'" & vbCrLf & _
                      " WHERE ShippingInstructionNo='" & Trim(pShippno) & "'  " & vbCrLf & _
                      " AND AffiliateID='" & Trim(pAFF) & "' " & vbCrLf & _
                      " AND ForwarderID='" & Trim(pFWD) & "' " & vbCrLf
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
        End Try
    End Sub

    Private Sub Excel_MovingList()
        'On Error GoTo ErrHandler
        Dim strFileSize As String = ""
        Dim ls_sql As String = ""

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        Dim fileTocopy As String
        Dim NewFileCopy As String
        Dim NewFileCopyas As String

        Dim KTime1 As String = ""
        Dim KTime2 As String = ""
        Dim KTime3 As String = ""
        Dim KTime4 As String = ""
        Dim pNamaFile As String = ""

        Dim jkanbanno As String
        Dim jQty As Long
        Dim jQtyBox As Long
        Dim jQtyPallet As Long

        Dim ds As New DataSet
        Dim dsHeader As New DataSet
        Dim dsDetail As New DataSet
        Dim dsETAETD As New DataSet
        Dim dsDetailDelivery As New DataSet

        Dim ls_SJ As String = ""
        Dim ls_Supplier As String = ""
        Dim ls_supplierName As String = ""
        Dim ls_supplierAdd As String = ""
        Dim ls_delivery As String = ""
        Dim ls_DeliveryName As String = ""
        Dim ls_deliveryAdd As String = ""
        Dim ls_orderNo As String = ""
        Dim ls_orderNo1 As String = ""
        Dim ls_ETDV As Date
        Dim ls_ETDP As Date
        Dim ls_ETAP As Date
        Dim ls_ETAF As Date
        Dim ls_Aff As String = ""
        Dim ls_AFFName As String = ""
        Dim ls_AffADD As String
        Dim ls_Attn As String = ""
        Dim ls_telp As String = ""
        Dim ls_Consignee As String = ""
        Dim i_loop As Long
        Dim xlApp = New Excel.Application

        Try
            ls_sql = " select distinct Consignee = isnull(ConsigneeCode,''), attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), period = PME.Period, DOM.SuratJalanNo, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') + ' ' + isnull(MS.City,'') + ' ' + isnull(MS.Postalcode,'') SUPPAddress,  " & vbCrLf & _
                     " PMEF.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  + ' ' + isnull(MF.City,'') + ' ' + isnull(MF.PostalCode,'') as FWDAddress, PME.ETDVendor1 as ETDVendor, PME.ETDPort1 as ETDPort, PME.ETAPort1 as ETAPort, PME.ETAFactory1 as ETAFactory, " & vbCrLf & _
                     " PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'')  + ' ' + isnull(MA.City,'') + ' ' + isnull(MA.PostalCode,'') as AFFAddress ,CASE WHEN PME.ShipCls = 'A' then 'AIR' else 'BOAT' END ShipCls from  " & vbCrLf & _
                     " DOSUpplier_Master_Export DOM " & vbCrLf & _
                     " INNER JOIN PO_Master_Export PME ON DOM.PONo = PME.PONo and DOM.SupplierID = PME.SupplierID and DOM.AffiliateID = PME.AffiliateID AND DOM.PONo = PME.PONo AND DOM.OrderNo = PME.OrderNo1 " & vbCrLf & _
                     " INNER JOIN PO_Master_Export PMEF ON DOM.PONo = PMEF.PONo and DOM.SupplierID = PMEF.SupplierID and DOM.AffiliateID = PMEF.AffiliateID AND DOM.PONo = PMEF.PONo AND DOM.SplitReffPONo = PMEF.OrderNo1 " & vbCrLf & _
                     " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                     " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PMEF.ForwarderID " & vbCrLf & _
                     " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                     " where isnull(DOM.MovingList,0) = '1' and isnull(DOM.PONO,'') <> '' and DOM.SuratJalanno <> '' "

            'MdlConn.ReadConnection()
            ds = cls.uf_GetDataSet(ls_sql)

            Dim ls_file As String

            For i_loop = 0 To ds.Tables(0).Rows.Count - 1

                '=======================================Delivery Instruction To Forwarder===========================================

                Dim fi3 As New FileInfo(Trim(txtAttachmentDOM.Text) & "\PO MOVING LIST.xlsx")

                If Not fi3.Exists Then
                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Note STOPPED, because File Excel isn't Found " & vbCrLf & _
                                    rtbProcess.Text
                    Exit Sub
                End If

                ls_SJ = Trim(ds.Tables(0).Rows(i_loop)("SuratJalanNo"))
                ls_Supplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))
                ls_supplierName = Trim(ds.Tables(0).Rows(i_loop)("suppliername"))
                ls_supplierAdd = Trim(ds.Tables(0).Rows(i_loop)("SuppAddress"))
                ls_delivery = Trim(ds.Tables(0).Rows(i_loop)("ForwarderID"))
                ls_DeliveryName = Trim(ds.Tables(0).Rows(i_loop)("ForwarderName"))
                ls_deliveryAdd = Trim(ds.Tables(0).Rows(i_loop)("FWDAddress"))
                ls_orderNo = Trim(ds.Tables(0).Rows(i_loop)("PONo"))
                ls_orderNo1 = Trim(ds.Tables(0).Rows(i_loop)("OrderNo"))
                ls_ETDV = Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "yyyy-MM-dd")
                ls_ETDP = Format((ds.Tables(0).Rows(i_loop)("ETDPort")), "yyyy-MM-dd")
                ls_ETAP = Format((ds.Tables(0).Rows(i_loop)("ETAPort")), "yyyy-MM-dd")
                ls_ETAF = Format((ds.Tables(0).Rows(i_loop)("ETAFactory")), "yyyy-MM-dd")
                ls_Aff = Trim(ds.Tables(0).Rows(i_loop)("AFF"))
                ls_Consignee = Trim(ds.Tables(0).Rows(i_loop)("Consignee"))
                ls_AFFName = Trim(ds.Tables(0).Rows(i_loop)("AFFName"))
                ls_AffADD = Trim(ds.Tables(0).Rows(i_loop)("AFFAddress"))
                ls_Attn = Trim(ds.Tables(0).Rows(i_loop)("attn"))
                ls_telp = Trim(ds.Tables(0).Rows(i_loop)("telp"))

                dsDetailDelivery = BindDataDeliveryInstruction(ls_SJ, ls_orderNo, ls_orderNo1, ls_Aff, ls_Supplier)


                If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                    Dim dsEmail As New DataSet
                    dsEmail = EmailToEmailCCKanban_Export(ls_Aff, ls_Supplier)
                    '1 CC Affiliate
                    '2 CC PASI
                    '3 CC & TO Supplier
                    For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(i)("KanbanCC")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(i)("KanbanCC")
                        End If

                        If dsEmail.Tables(0).Rows(i)("KanbanTO") <> "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(i)("KanbanTO")
                        End If
                    Next

                    Dim k As Long
                    Dim dsAffiliate As New DataSet
                    dsAffiliate = Affiliate(Trim(ls_Aff))

                    Dim dsSupplier As New DataSet
                    dsSupplier = Supplier(Trim(ls_Supplier))

                    Dim status As Boolean
                    status = True

                    If dsDetailDelivery.Tables(0).Rows.Count = 0 Then
                        status = False
                    Else
                        status = True
                    End If

                    If status = True Then
                        NewFileCopy = Trim(txtAttachmentDOM.Text) & "\PO MOVING LIST.xlsx"
                        ls_file = NewFileCopy
                        ExcelBook = xlApp.Workbooks.Open(ls_file)
                        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                        ExcelSheet.Range("H1").Value = "POMV"
                        ExcelSheet.Range("H2").Value = receiptEmail
                        ExcelSheet.Range("H3").Value = ls_Aff
                        ExcelSheet.Range("H3").Value = ls_Consignee
                        ExcelSheet.Range("H4").Value = ls_delivery
                        ExcelSheet.Range("H5").Value = ls_Supplier

                        ExcelSheet.Range("AE11:AT11").Merge()
                        ExcelSheet.Range("AE11:AT11").Value = ls_supplierName
                        ExcelSheet.Range("AE12:AT15").Merge()
                        ExcelSheet.Range("AE12:AT15").Value = ls_supplierAdd
                        ExcelSheet.Range("AE19:AT19").Merge()
                        ExcelSheet.Range("AE19:AT19").Value = ls_DeliveryName
                        ExcelSheet.Range("AE20:AT22").Merge()
                        ExcelSheet.Range("AE20:AT22").Value = ls_deliveryAdd
                        ExcelSheet.Range("I28:P28").Merge()


                        'ExcelSheet.Range("I28:P28").Value = ls_SJ
                        'ExcelSheet.Range("I23:X23").Value = "ATTN : " & Trim(ls_Attn) & "     TELP : " & Trim(ls_telp)
                        'Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "yyyy-MM-dd")
                        ExcelSheet.Range("G11:K11").Value = Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy-MM")

                        ExcelSheet.Range("G13:K13").Value = ls_orderNo1
                        ExcelSheet.Range("G15:K15").Value = ls_orderNo

                        ExcelSheet.Range("G17:K17").Value = ds.Tables(0).Rows(i_loop)("ShipCls")

                        'ExcelSheet.Range("AE11:AE11").Value = Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy-MM-dd")

                        ExcelSheet.Range("R11:V11").Merge()
                        ExcelSheet.Range("R11:V11").Value = ls_ETDV
                        ExcelSheet.Range("R13:V13").Merge()
                        ExcelSheet.Range("R13:V13").Value = ls_ETDP
                        ExcelSheet.Range("R15:V15").Merge()
                        ExcelSheet.Range("R15:V15").Value = ls_ETAP
                        ExcelSheet.Range("R17:V17").Merge()
                        ExcelSheet.Range("R17:V17").Value = ls_ETAF

                        ExcelSheet.Range("G19:V19").Merge()
                        ExcelSheet.Range("G19:V19").Value = ls_AFFName

                        ExcelSheet.Range("G20:V22").Merge()
                        ExcelSheet.Range("G20:V22").Value = ls_AffADD
                        k = 0

                        For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                            'For i = 0 To 3
                            k = k
                            ExcelSheet.Range("B" & k + 27 & ": C" & k + 27).Merge()
                            ExcelSheet.Range("D" & k + 27 & ": H" & k + 27).Merge()
                            ExcelSheet.Range("I" & k + 27 & ": P" & k + 27).Merge()
                            ExcelSheet.Range("Q" & k + 27 & ": R" & k + 27).Merge()
                            ExcelSheet.Range("S" & k + 27 & ": T" & k + 27).Merge()
                            ExcelSheet.Range("U" & k + 27 & ": X" & k + 27).Merge()
                            ExcelSheet.Range("Y" & k + 27 & ": AB" & k + 27).Merge()
                            ExcelSheet.Range("AC" & k + 27 & ": AF" & k + 27).Merge()
                            ExcelSheet.Range("AG" & k + 27 & ": AJ" & k + 27).Merge()

                            ExcelSheet.Range("B" & k + 27 & ": C" & k + 27).Value = k + 1
                            ExcelSheet.Range("D" & k + 27 & ": H" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partno"))
                            ExcelSheet.Range("I" & k + 27 & ": P" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Partname"))

                            ExcelSheet.Range("Q" & k + 27 & ": R" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("uom")
                            ExcelSheet.Range("S" & k + 27 & ": T" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("qtybox")
                            ExcelSheet.Range("U" & k + 27 & ": X" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("qty")
                            ExcelSheet.Range("Y" & k + 27 & ": AB" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("BoxQty")

                            ExcelSheet.Range("AC" & k + 27 & ": AF" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo1"))
                            ExcelSheet.Range("AG" & k + 27 & ": AJ" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo2"))

                            k = k + 1
                        Next

                        DrawAllBorders(ExcelSheet.Range("B27" & ": AJ" & k + 26))

                        'Save ke Local
                        xlApp.DisplayAlerts = False

                        ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\PO Moving List-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & "-" & Trim(ls_SJ) & ".xlsx")
                        pNamaFile = "\PO Moving List-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & "-" & Trim(ls_SJ) & ".xlsx"


                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If

                    '---------------------------------------excel---------------------------------------'
                    If sendEmailtoSupllierDeliveryConfirmation("MOVING", pNamaFile, "", Trim(ls_orderNo), ls_delivery, Trim(ls_Aff), Trim(ls_orderNo1), Trim(ls_Supplier), Trim(ls_SJ)) = False Then GoTo keluar
                    Call UpdateStatusMovingExport(ls_Aff, ls_Supplier, ls_orderNo, ls_orderNo1, ls_SJ)
                End If
                '=======================================Delivery Instruction To Forwarder===========================================

keluar:
                xlApp.Workbooks.Close()
                xlApp.Quit()
            Next
            Exit Sub
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery to Forwarder STOPPED, because " & Err.Description & " " & vbCrLf & _
                                    rtbProcess.Text
            xlApp.Workbooks.Close()
            xlApp.Quit()
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try

        'ErrHandler:
        '            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery to Forwarder STOPPED, because " & Err.Description & " " & vbCrLf & _
        '                                rtbProcess.Text
        '            xlApp.Workbooks.Close()
        '            xlApp.Quit()
    End Sub

    Private Sub UpdateStatusMovingExport(ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pOrderNo As String, ByVal pSJ As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " update dbo.DOSupplier_Master_Export set MovingList = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND SuratJalanno = '" & pSJ & "'" & vbCrLf & _
                         " AND PONo = '" & pPoNo & "'" & vbCrLf & _
                         " AND OrderNo = '" & pOrderNo & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub

    Private Sub pGetPDFShipping()
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""
        Dim pFilePDF1 As String = ""
        Dim pFilePDF2 As String = ""
        Dim pFilePDF3 As String = ""
        Dim pCSV As String = ""
        Dim pFile As String = ""

        'copy file from server to local

        'MdlConn.ReadConnection()
        Dim ls_SQL As String = ""
        Dim ds As New DataSet
        Dim dsDetail As New DataSet
        Dim sh_Affiliate As String = "", sh_ForwarderID As String, sh_Shippingno As String
        'Dim xlApp = New Excel.Application

        Try
            ls_SQL = " Select distinct * From ShippingInstruction_Master where isnull(TallyCls,'') = '1' "
            'ls_SQL = " Select distinct * From ShippingInstruction_Master where ShippingInstructionNo = 'EAPY60009B' "

            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pDate = Now
                    sh_Affiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    sh_ForwarderID = Trim(ds.Tables(0).Rows(xi)("ForwarderID"))
                    sh_Shippingno = Trim(ds.Tables(0).Rows(xi)("ShippingInstructionNo"))

                    'PDF SHIPPING INSTRUCTION
                    pFilePDF1 = CreateShippingInstructionToPDF(sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    If pFilePDF1 = "" Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & pFilePDF1 & " NG. " & vbCrLf & _
                        rtbProcess.Text
                        GoTo keluar
                    Else
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & pFilePDF1 & " OK. " & vbCrLf & _
                        rtbProcess.Text
                    End If
                    'PDF SHIPPING INSTRUCTION

                    ''PDF TALLY
                    'pFilePDF2 = CreateTallyToPDF(sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    'If pFilePDF2 = "" Then
                    '    GoTo keluar
                    'End If
                    ''PDF TALLY

                    ''PDF COMMERCIAL INVOICE
                    'pFilePDF3 = CreateCommercialInvoiceToPDF(sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    'If pFilePDF3 = "" Then
                    '    GoTo keluar
                    'End If
                    ''PDF COMMERCIAL INVOICE

                    'CSV
                    pCSV = GetCSV(sh_Shippingno, sh_Affiliate, sh_ForwarderID)
                    If pCSV = "" Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & pCSV & " NG. " & vbCrLf & _
                        rtbProcess.Text
                        GoTo keluar
                    Else
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & pCSV & " OK. " & vbCrLf & _
                        rtbProcess.Text
                    End If
                    'CSV

                    If sendEmailTallyInvoceShippingToForwarder(pFilePDF1, pFilePDF2, pFilePDF3, sh_Shippingno, sh_Affiliate, sh_ForwarderID, pCSV) = False Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & pCSV & " NG. " & vbCrLf & _
                        rtbProcess.Text
                        GoTo keluar
                    Else
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & pCSV & " OK. " & vbCrLf & _
                        rtbProcess.Text
                    End If


                    Call UpdateShippingExcelCls(True, sh_Shippingno, sh_Affiliate, sh_ForwarderID)

                    Thread.Sleep(500)
keluar:
                Next
            Else
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Shipping Instruction STOPPED, because there is nothing PO to send " & vbCrLf & _
                    rtbProcess.Text
            End If
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Shipping Instruction STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally

        End Try
    End Sub

    Private Function CreateShippingInstructionToPDF(ByVal PshippingNo As String, ByVal pAffiliate As String, ByVal pForwarder As String) As String
        Dim CrReport As New rptShippingInstruction()
        Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
        Dim pFile As String

        Try
            txtMsg.Text = ""
            Cursor.Current = Cursors.WaitCursor
            Dim dsPrint As New DataSet
            dsPrint = PrintShippingInstruction(PshippingNo, pAffiliate, pForwarder)

            If dsPrint.Tables(0).Rows.Count = 0 Then Exit Try

            CrReport.SetDatabaseLogon(cfg.User, cfg.Password, cfg.Server, cfg.Database)
            CrReport.SetDataSource(dsPrint.Tables(0))

            CrDiskFileDestinationOptions.DiskFileName = Trim(txtSaveAsDOM.Text) & "\ShippingInstruction-" & PshippingNo & "-" & pAffiliate.Trim & "-" & pForwarder.Trim & ".pdf"
            pFile = ""
            pFile = "ShippingInstruction-" & PshippingNo & "-" & pAffiliate.Trim & "-" & pForwarder.Trim & ".pdf"
            CrExportOptions = CrReport.ExportOptions

            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With

            Try
                CrReport.Export()
            Catch err As Exception
                MessageBox.Show(err.ToString())
            End Try
            'PDF
            CreateShippingInstructionToPDF = pFile

            CrReport.Dispose()
            CrReport.Close()
        Catch ex As Exception
            CreateShippingInstructionToPDF = ""
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Create PDF Shipping Instruction STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally
            If Not CrReport Is Nothing Then
                NAR(CrReport)
                GC.Collect()
            End If
            If Not CrReport Is Nothing Then
                CrReport.Dispose()
                CrReport.Close()
            End If
        End Try
    End Function

    Private Function PrintShippingInstruction(ByVal ls_value1 As String, ByVal ls_value2 As String, ByVal ls_value3 As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = ""
        ls_SQL = "  select       " & vbCrLf & _
                  "  ShippingInstructionNo = SIM.ShippingInstructionNo,           " & vbCrLf & _
                  "  FWD = Rtrim(MF.ForwarderName) + ' ' + Rtrim(MF.Address) + ' ' + Rtrim(MF.City) + ' ' + Rtrim(MF.PostalCode),          " & vbCrLf & _
                  "  ATT = isnull(Rtrim(MF.Attn),''),           " & vbCrLf & _
                  "  FAx = isnull(Rtrim(MF.Fax),''),           " & vbCrLf & _
                  "  Tujuan = isnull(TM.DestinationPort,''),           " & vbCrLf & _
                  "  Shipment = Case when ShipCls = 'B' then 'SEA FREIGHT' ELSE 'AIR FREIGHT' END,           " & vbCrLf & _
                  "  Vessel = Vessel,           " & vbCrLf & _
                  "  ETD = Convert(Char(12), convert(Datetime, isnull(SIM.ETDPort,POM.ETDPort1)),106),           " & vbCrLf & _
                  "  ETA = Convert(Char(12), convert(Datetime, isnull(SIM.ETAPort,POM.ETAPort1)),106),           " & vbCrLf & _
                  "  tgltiba = Convert(Char(12), convert(Datetime, isnull(SIM.ETAPort,POM.ETAPort1)),106),          "

        ls_SQL = ls_SQL + "  part = 'Automotive Component',           " & vbCrLf & _
                          "  jumlah = 0,           " & vbCrLf & _
                          "  pallet = isnull(SUMTally.palletno,0),      " & vbCrLf & _
                          "  Box = isnull(Sumtally.box,0), " & vbCrLf & _
                          "  Qty = SUM(isnull(SD.ShippingQty,0)),          " & vbCrLf & _
                          "  beratBersih = SUM(((netweight/ISNULL(SD.POQtyBox,MPM.QtyBox))* SD.ShippingQty)/1000),           " & vbCrLf & _
                          "  beratKotor = SIM.GrossWeight, --SUM(((grossweight/MPM.QtyBox)* SD.ShippingQty)/1000),           " & vbCrLf & _
                          "  Buyer = Rtrim(BuyerName),           " & vbCrLf & _
                          "  BuyerAddress = Rtrim(BuyerAddress), " & vbCrLf & _
                          "  Consignee = Rtrim(MA.ConsigneeName), ConsigneeAddress = Rtrim(MA.ConsigneeAddress), Attn =isnull(MSA.AffiliatePOTo,''), " & vbCrLf & _
                          "  Freight = isnull(Freight,''),           " & vbCrLf & _
                          "  Stuffing = Convert(Char(12), convert(Datetime, isnull(TM.Stuffingdate,'')),106)          "

        ls_SQL = ls_SQL + "  From ShippingInstruction_master SIM           " & vbCrLf & _
                          "  LEFT JOIN ShippingInstruction_Detail SD           " & vbCrLf & _
                          "  ON SIM.ShippingInstructionNo = SD.ShippingInstructionNo           " & vbCrLf & _
                          "  	AND SIM.AffiliateID = SD.AffiliateID         " & vbCrLf & _
                          "  	AND SIM.ForwarderID = SD.ForwarderID           " & vbCrLf & _
                          "  LEFT JOIN PO_Detail_Export POD           " & vbCrLf & _
                          "  ON POD.Pono = SD.OrderNo           " & vbCrLf & _
                          "  	AND POD.AffiliateID = SD.AffiliateID           " & vbCrLf & _
                          "  	AND POD.PartNo = SD.PartNo           " & vbCrLf & _
                          "  	AND POD.SupplierID = SD.SupplierID           " & vbCrLf & _
                          "  LEFT JOIN PO_Master_Export POM          "

        ls_SQL = ls_SQL + "  ON POM.PONo = POD.PONo           " & vbCrLf & _
                          "  	AND POM.AffiliateID = POD.AffiliateID           " & vbCrLf & _
                          "  	AND POM.SupplierID = POD.SupplierID           " & vbCrLf & _
                          "  	AND POM.OrderNo1 = POD.OrderNo1         " & vbCrLf & _
                          "  LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SIM.AffiliateID           " & vbCrLf & _
                          "  LEFT JOIN ms_emailAffiliate_Export MSA ON MSA.AffiliateID = MA.AffiliateID           " & vbCrLf & _
                          "  LEFT JOIN MS_Parts MP ON MP.PartNo = SD.PartNo           " & vbCrLf & _
                          "  LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SD.PartNo           " & vbCrLf & _
                          "  	AND MPM.AffiliateID = SD.AffiliateID           " & vbCrLf & _
                          "  	AND MPM.SupplierID = SD.SupplierID           " & vbCrLf & _
                          "  LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SIM.ForwarderID      " & vbCrLf & _
                          "  LEFT JOIN Tally_Master TM ON TM.ShippingInstructionNo = SIM.ShippingInstructionNo      "

        ls_SQL = ls_SQL + "  	AND TM.AffiliateID = SIM.AffiliateID      " & vbCrLf & _
                          "  	AND TM.ForwarderID = SIM.ForwarderID      " & vbCrLf & _
                          "  LEFT JOIN (Select distinct ShippingInstructionNo, ForwarderID, AffiliateID, palletno = Count(palletno), weightpallet = sum(weightpallet),      " & vbCrLf & _
                          "  			width = sum(width), height = sum(height), length = sum(length), M3 = Sum(M3), box = SUM(box)   " & vbCrLf & _
                          "  		   from (     " & vbCrLf & _
                          "  					select ShippingInstructionNo, ForwarderID, AffiliateID, palletno, weightpallet = sum(weightpallet),      " & vbCrLf & _
                          "  					width = sum(width), height = sum(height), length = sum(length), M3 = SUM(M3), box =SUM(TOTALBOX) from Tally_Detail     					 " & vbCrLf & _
                          "  					group by ShippingInstructionNo, ForwarderID, AffiliateID, palletno   " & vbCrLf & _
                          "  				) x group by ShippingInstructionNo, ForwarderID, AffiliateID) SUMTally     " & vbCrLf & _
                          "  ON SUMTally.ShippingInstructionNo = SD.ShippingInstructionNo      " & vbCrLf & _
                          "  AND SUMTally.ForwarderID = SD.ForwarderID      "

        ls_SQL = ls_SQL + "  AND SUMTally.AffiliateID = SD.AffiliateID        " & vbCrLf & _
                          "  " & vbCrLf & _
                          " WHERE SIM.ShippingInstructionNo = '" & Trim(ls_value1) & "' " & vbCrLf & _
                          " AND SIM.AffiliateID = '" & Trim(ls_value2) & "'" & vbCrLf & _
                          " AND SIM.ForwarderID = '" & Trim(ls_value3) & "' " & vbCrLf & _
                          "  GROUP BY      " & vbCrLf & _
                          "  SIM.ShippingInstructionNo,          " & vbCrLf & _
                          "  Rtrim(MF.ForwarderName) ,Rtrim(MF.Address) , Rtrim(MF.City) , Rtrim(MF.PostalCode), " & vbCrLf & _
                          "  isnull(Rtrim(MF.Attn),''),  " & vbCrLf & _
                          "  isnull(Rtrim(MF.Fax),''),          " & vbCrLf & _
                          "  isnull(TM.DestinationPort,''), " & vbCrLf & _
                          "  POM.ETDPort1, SIM.ETDPort, SIM.ETAPort, TM.Stuffingdate, MSA.AffiliatePOTo, " & vbCrLf

        ls_SQL = ls_SQL + "  POM.ETAPort1, " & vbCrLf & _
                          "  Vessel = ISNULL(SIM.Vessels,''), " & vbCrLf & _
                          "  Rtrim(BuyerName),Rtrim(BuyerAddress),Rtrim(MA.ConsigneeName), Rtrim(MA.ConsigneeAddress) , " & vbCrLf & _
                          "  ShipCls, SUMTally.palletno,Sumtally.box, Freight,SIM.GrossWeight  "

        ls_SQL = " SELECT DISTINCT " & vbCrLf & _
                  "   ShippingInstructionNo = SIM.ShippingInstructionNo,            " & vbCrLf & _
                  "   FWD = Rtrim(MF.ForwarderName) + ' ' + Rtrim(MF.Address) + ' ' + Rtrim(MF.City) + ' ' + Rtrim(MF.PostalCode),           " & vbCrLf & _
                  "   ATT = isnull(Rtrim(MF.Attn),''),            " & vbCrLf & _
                  "   FAx = isnull(Rtrim(MF.Fax),''),            " & vbCrLf & _
                  "   Tujuan = isnull(MA.DestinationPort,''), " & vbCrLf & _
                  "   Shipment = Case when TypeOfService = 'FCL' then 'SEA FREIGHT' WHEN TypeOfService = 'LCL' then 'SEA FREIGHT' ELSE 'AIR FREIGHT' END, " & vbCrLf & _
                  "   Vessel = ISNULL(SIM.Vessels,''), " & vbCrLf & _
                  "   ETD = Convert(Char(12), convert(Datetime, SIM.ETDPort),106),            " & vbCrLf & _
                  "   ETA = Convert(Char(12), convert(Datetime, SIM.ETAPort),106),            " & vbCrLf & _
                  "   tgltiba = Convert(Char(12), convert(Datetime, SIM.ETAPort),106),             "

        ls_SQL = ls_SQL + "   part = 'Automotive Component',            " & vbCrLf & _
                          "   jumlah = 0, " & vbCrLf & _
                          "   pallet = isnull(SIM.TotalPallet,0), " & vbCrLf & _
                          "   Box = isnull(SUM(SD.ShippingQty/ISNULL(SD.POQtyBox,MPM.QtyBox)),0), " & vbCrLf & _
                          "   Qty = SUM(isnull(SD.ShippingQty,0)), " & vbCrLf & _
                          "   beratBersih = SUM(((netweight/ISNULL(SD.POQtyBox,MPM.QtyBox))* SD.ShippingQty)/1000),            " & vbCrLf & _
                          "   beratKotor = ISNULl(SIM.GrossWeight,0), " & vbCrLf & _
                          "   Buyer = Rtrim(BuyerName),            " & vbCrLf & _
                          "   BuyerAddress = Rtrim(BuyerAddress),  " & vbCrLf & _
                          "   Consignee = Rtrim(MA.ConsigneeName), ConsigneeAddress = Rtrim(MA.ConsigneeAddress), Attn =isnull(MA.Att,''),  " & vbCrLf & _
                          "   Freight = isnull(Freight,'') "

        ls_SQL = ls_SQL + " From ShippingInstruction_master SIM            " & vbCrLf & _
                          " LEFT JOIN ShippingInstruction_Detail SD            " & vbCrLf & _
                          " 	ON SIM.ShippingInstructionNo = SD.ShippingInstructionNo            " & vbCrLf & _
                          " 	AND SIM.AffiliateID = SD.AffiliateID          " & vbCrLf & _
                          " 	AND SIM.ForwarderID = SD.ForwarderID " & vbCrLf & _
                          " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SIM.ForwarderID " & vbCrLf & _
                          " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SIM.AffiliateID " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SD.PartNo  "

        ls_SQL = ls_SQL + " 	and MPM.SupplierID = SD.SupplierID  " & vbCrLf & _
                          " 	and MPM.AffiliateID = SD.AffiliateID " & vbCrLf & _
                          " WHERE SIM.ShippingInstructionNo = '" & Trim(ls_value1) & "' " & vbCrLf & _
                          " AND SIM.AffiliateID = '" & Trim(ls_value2) & "'" & vbCrLf & _
                          " AND SIM.ForwarderID = '" & Trim(ls_value3) & "' " & vbCrLf & _
                          " GROUP BY SIM.ShippingInstructionNo, " & vbCrLf & _
                          " 	Rtrim(MF.ForwarderName) ,Rtrim(MF.Address) , Rtrim(MF.City) , Rtrim(MF.PostalCode),  " & vbCrLf & _
                          " 	isnull(Rtrim(MF.Attn),''),   " & vbCrLf & _
                          " 	isnull(Rtrim(MF.Fax),''),           " & vbCrLf & _
                          " 	isnull(MA.DestinationPort,''), " & vbCrLf & _
                          " 	TypeOfService, SIM.VesselS, "

        ls_SQL = ls_SQL + " 	SIM.ETDPort, SIM.ETAPort, SIM.TotalPallet, SIM.GrossWeight, " & vbCrLf & _
                          " 	Rtrim(BuyerName),Rtrim(BuyerAddress),Rtrim(MA.ConsigneeName), Rtrim(MA.ConsigneeAddress), MA.Att, SIM.Freight " & vbCrLf & _
                          "  "

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function GetCSV(ByVal sh_Shippingno, ByVal sh_Affiliate, ByVal sh_ForwarderID) As String
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        'MdlConn.ReadConnection()
        Dim ls_SQL As String = ""
        Dim ds As New DataSet
        Dim dsDetail As New DataSet
        Dim xlApp = New Excel.Application

        Try
            dsDetail = qCsv(sh_Shippingno, sh_Affiliate, sh_ForwarderID)

            xlApp = CreateObject("Excel.Application")
            ExcelBook = xlApp.Workbooks.Add
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            i = 0
            ExcelSheet.Range("A" & i + 1).Value = "Invoice No."
            ExcelSheet.Range("B" & i + 1).Value = "Consignee Code"
            ExcelSheet.Range("C" & i + 1).Value = "Buyer Code"
            ExcelSheet.Range("D" & i + 1).Value = "Shipment"
            ExcelSheet.Range("E" & i + 1).Value = "Shipping Line"
            ExcelSheet.Range("F" & i + 1).Value = "Vessel"
            ExcelSheet.Range("G" & i + 1).Value = "Voyage"
            ExcelSheet.Range("H" & i + 1).Value = "From Port"
            ExcelSheet.Range("I" & i + 1).Value = "VIA"
            ExcelSheet.Range("J" & i + 1).Value = "To Port"
            ExcelSheet.Range("K" & i + 1).Value = "ETD"
            ExcelSheet.Range("L" & i + 1).Value = "ETA"
            ExcelSheet.Range("M" & i + 1).Value = "Order No"
            ExcelSheet.Range("N" & i + 1).Value = "Original O/No"
            ExcelSheet.Range("O" & i + 1).Value = "Part No"
            ExcelSheet.Range("P" & i + 1).Value = "Part Group Name"
            ExcelSheet.Range("Q" & i + 1).Value = "Box No. From"
            ExcelSheet.Range("R" & i + 1).Value = "Box No. To"
            ExcelSheet.Range("S" & i + 1).Value = "Carton Count"
            ExcelSheet.Range("T" & i + 1).Value = "Quantity"
            ExcelSheet.Range("U" & i + 1).Value = "Total Quantity"
            ExcelSheet.Range("V" & i + 1).Value = "Net(Weight(KGM))"

            If dsDetail.Tables(0).Rows.Count > 0 Then
                For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                    'Header
                    ExcelSheet.Range("A" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("InvoiceNo"))
                    ExcelSheet.Range("B" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Consignee"))
                    ExcelSheet.Range("C" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Buyer"))
                    ExcelSheet.Range("D" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Shipment"))
                    ExcelSheet.Range("E" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("ShippingLine"))
                    ExcelSheet.Range("F" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Vessel"))
                    ExcelSheet.Range("G" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Voyage"))
                    ExcelSheet.Range("H" & i + 2).Value = "JAKARTA" 'Trim(dsDetail.Tables(0).Rows(i)("FromPort"))
                    ExcelSheet.Range("I" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("VIA"))
                    ExcelSheet.Range("J" & i + 2).Value = Trim(Replace(dsDetail.Tables(0).Rows(i)("ToPort"), ",", ""))
                    ExcelSheet.Range("K" & i + 2).NumberFormat = "@"
                    ExcelSheet.Range("K" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("ETD"))
                    ExcelSheet.Range("L" & i + 2).NumberFormat = "@"
                    ExcelSheet.Range("L" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("ETA"))
                    ExcelSheet.Range("M" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("OrderNo"))
                    ExcelSheet.Range("N" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("OriginalNo"))
                    ExcelSheet.Range("O" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                    ExcelSheet.Range("P" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("PartGroupName"))
                    ExcelSheet.Range("Q" & i + 2).Value = Trim(Split(dsDetail.Tables(0).Rows(i)("BoxNo"), "-")(0))
                    ExcelSheet.Range("R" & i + 2).Value = Trim(Split(dsDetail.Tables(0).Rows(i)("BoxNo"), "-")(1))
                    ExcelSheet.Range("S" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("CartonCount"))
                    ExcelSheet.Range("T" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("QtyBox"))
                    ExcelSheet.Range("U" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Quantity"))
                    ExcelSheet.Range("V" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Net"))
                Next
            End If

            xlApp.DisplayAlerts = False

            Dim temp_Filename As String = "CSV " & Trim(sh_Affiliate) & "-" & Trim(sh_ForwarderID) & "-" & Trim(sh_Shippingno) & ".csv"
            GetCSV = temp_Filename

            ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\" & "CSV " & Trim(sh_Affiliate) & "-" & Trim(sh_ForwarderID) & "-" & Trim(sh_Shippingno) & ".csv", FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, CreateBackup:=False)
            ExcelBook.Close()

            xlApp.Quit()

        Catch ex As Exception
            GetCSV = ""
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Function

    Private Function qCsv(ByVal shShippingNo As String, ByVal shAffiliate As String, ByVal shForwarder As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        
        ls_SQL = " SELECT  DISTINCT " & vbCrLf & _
                  " 	InvoiceNo = SHM.ShippingInstructionNo,  " & vbCrLf & _
                  " 	Consignee = MA.ConsigneeCode, " & vbCrLf & _
                  " 	Buyer = MA.BuyerCode, " & vbCrLf & _
                  " 	Shipment = case when POM.ShipCls='A' then 'Air Freight' else 'Sea Freight' End, " & vbCrLf & _
                  " 	ShippingLine = ISNULL(SHM.ShippingLineS,''),   " & vbCrLf & _
                  " 	Vessel = isnull(SHM.NamaKapalS,''),   " & vbCrLf & _
                  " 	Voyage = isnull(SHM.VesselS,''), " & vbCrLf & _
                  " 	FromPort = MF.Port, " & vbCrLf & _
                  " 	VIA = ISNULL(SHM.Via,''), " & vbCrLf & _
                  " 	ToPort = MA.DestinationPort, "

        ls_SQL = ls_SQL + " 	ETD = Replace(CONVERT(CHAR(11), CONVERT(DATETIME, SHM.ETDPort), 104),'.',''),     " & vbCrLf & _
                          " 	ETA = Replace(CONVERT(CHAR(11), CONVERT(DATETIME, SHM.ETAPort), 104),'.',''), " & vbCrLf & _
                          " 	OrderNo = RD.OrderNo,   " & vbCrLf & _
                          " 	OriginalNo = RD.PONo, " & vbCrLf & _
                          " 	PartNo = SDM.PartNo,   " & vbCrLf & _
                          " 	PartGroupName = isnull(PartGroupName,''), " & vbCrLf & _
                          " 	SDM.BoxNo, " & vbCrLf & _
                          " 	CartonCount = SDM.BoxQty, " & vbCrLf & _
                          " 	QtyBox = ISNULL(SDM.POQtyBox,MPM.QtyBox),   " & vbCrLf & _
                          " 	Quantity = ISNULL(SDM.POQtyBox,MPM.QtyBox) * SDM.BoxQty,   " & vbCrLf & _
                          " 	Net = MPM.NetWeight /1000 "

        ls_SQL = ls_SQL + " FROM ShippingInstruction_Master SHM  " & vbCrLf & _
                          " INNER JOIN ShippingInstruction_Detail SDM  " & vbCrLf & _
                          " 	ON ltrim(SDM.ShippingInstructionNo) = ltrim(SHM.ShippingInstructionNo)    " & vbCrLf & _
                          "   	AND ltrim(SDM.ForwarderID) = rtrim(SHM.ForwarderID)    " & vbCrLf & _
                          "   	AND rtrim(SDM.AffiliateID) = rtrim(SHM.AffiliateID) " & vbCrLf & _
                          " LEFT JOIN PO_Master_Export POM ON POM.OrderNo1 = SDM.OrderNo    " & vbCrLf & _
                          "      AND POM.AffiliateID = SDM.AffiliateID   " & vbCrLf & _
                          "      AND POM.SupplierID = SDM.SupplierID " & vbCrLf & _
                          " LEFT JOIN ReceiveForwarder_Detail RD ON RD.SuratJalanNO = SDM.SuratJalanno    " & vbCrLf & _
                          "    	AND RD.AffiliateID = SDM.AffiliateID     	 " & vbCrLf & _
                          " 	AND RD.SupplierID = SDM.SupplierID     	 "

        ls_SQL = ls_SQL + "    	AND RD.OrderNO = SDM.OrderNo    " & vbCrLf & _
                          "   	AND RD.PartNo = SDM.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SDM.AffiliateID " & vbCrLf & _
                          " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SHM.ForwarderID " & vbCrLf & _
                          " LEFT JOIN MS_Parts MP ON MP.PartNo = SDM.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SDM.PartNo  " & vbCrLf & _
                          " 	AND MPM.AffiliateID = SDM.AffiliateID AND MPM.SupplierID = SDM.SupplierID " & vbCrLf & _
                          " Where SHM.ShippingInstructionNo = '" & Trim(shShippingNo) & "' " & vbCrLf & _
                          " AND SHM.AffiliateID = '" & Trim(shAffiliate) & "' " & vbCrLf & _
                          " AND SHM.ForwarderID = '" & Trim(shForwarder) & "' "

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Private Function sendEmailTallyInvoceShippingToForwarder(ByVal pFileName1 As String, ByVal pFileName2 As String, ByVal pFileName3 As String, ByVal shshipno As String, ByVal shaffiliate As String, ByVal shforwarder As String, ByVal pCsv As String) As Boolean
        Try
            Dim TempFilePath As String
            Dim TempFileName1 As String
            Dim TempFileName2 As String
            Dim TempFileName3 As String
            Dim TempFileName4 As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)
            TempFileName1 = "\" & pFileName1
            TempFileName2 = "\" & pFileName2
            TempFileName3 = "\" & pFileName3
            TempFileName4 = "\" & pCsv

            Dim dsEmail As New DataSet
            dsEmail = EmailSH(shaffiliate, "PASI", shforwarder)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "FWD" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("SupplierDeliveryTo")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("SupplierDeliveryTo")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "FWD" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("SupplierDeliverycc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("SupplierDeliverycc")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            'receiptCCEmail = "pasi-opa02@pemi.co.id;edi@tos.co.id"
            'receiptEmail = "pasi-opa02@pemi.co.id;edi@tos.co.id"

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailTallyInvoceShippingToForwarder = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailTallyInvoceShippingToForwarder = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            'mailMessage.Subject = "[TRIAL] Send Shipping Data To Forwarder, Invoice No: " & shshipno.Trim & ""
            If pFileName1 <> "" Then
                mailMessage.Subject = "SI-" & shaffiliate.Trim & "-" & shshipno.Trim & " Shipping Instruction "
            Else
                mailMessage.Subject = "INV-" & shaffiliate.Trim & "-" & shshipno.Trim & " TAX INVOICE "
            End If
            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If
            'Dim receiptBCCEmail As String = "pasi.purchase@gmail.com"
            'If receiptBCCEmail <> "" Then
            '    For Each recipientBCC In receiptBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If
            GetSettingEmail_Export("PO")
            ls_Body = clsNotification.GetNotification("22", "", shshipno.Trim)
            mailMessage.Body = ls_Body

            Dim filename1 As String = TempFilePath & TempFileName1
            If TempFileName1 <> "\" Then mailMessage.Attachments.Add(New Attachment(filename1))

            Dim filename2 As String = TempFilePath & TempFileName2
            If TempFileName2 <> "\" Then mailMessage.Attachments.Add(New Attachment(filename2))

            Dim filename3 As String = TempFilePath & TempFileName3
            If TempFileName3 <> "\" Then mailMessage.Attachments.Add(New Attachment(filename3))

            Dim filename4 As String = TempFilePath & TempFileName4
            If TempFileName4 <> "\" Then mailMessage.Attachments.Add(New Attachment(filename4))

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            'smtp.Host = "smtp.atisicloud.com"
            'smtp.Host = "mail.fast.net.id"

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
            sendEmailTallyInvoceShippingToForwarder = True
            'Delete the file
            'Kill(TempFilePath & TempFileName)
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Tally,Invoice,Shipping: " & pPONo & " to Supplier SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
            Exit Function
        Catch ex As Exception
            sendEmailTallyInvoceShippingToForwarder = False
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Tally,Invoice,Shipping: " & pPONo & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text

        End Try

    End Function

    Private Sub UpdateShippingExcelCls(ByVal pIsNewData As Boolean, _
                         Optional ByVal pShippno As String = "", _
                         Optional ByVal pAFF As String = "", _
                         Optional ByVal pFWD As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & vbCrLf & _
                      " SET TallyCls='2'" & vbCrLf & _
                      " WHERE ShippingInstructionNo='" & Trim(pShippno) & "'  " & vbCrLf & _
                      " AND AffiliateID='" & Trim(pAFF) & "' " & vbCrLf & _
                      " AND ForwarderID='" & Trim(pFWD) & "' " & vbCrLf
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
        End Try
    End Sub

    Private Sub pGetPDFINVEX()
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""
        Dim pFilePDF1 As String = ""
        Dim pFilePDF2 As String = ""
        Dim pFilePDF3 As String = ""
        Dim pFile As String = ""

        Dim pTerm As String = ""
        Dim pService As String = ""

        'copy file from server to local

        'MdlConn.ReadConnection()
        Dim ls_SQL As String = ""
        Dim ds As New DataSet
        Dim dsDetail As New DataSet
        Dim sh_Affiliate As String = "", sh_ForwarderID As String, sh_Shippingno As String
        'Dim xlApp = New Excel.Application

        Try
            ls_SQL = " Select distinct * From ShippingInstruction_Master where isnull(sendInvoice,'') = '1' "

            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pDate = Now
                    sh_Affiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    sh_ForwarderID = Trim(ds.Tables(0).Rows(xi)("ForwarderID"))
                    sh_Shippingno = Trim(ds.Tables(0).Rows(xi)("ShippingInstructionNo"))
                    pTerm = Trim(ds.Tables(0).Rows(xi)("TermDelivery"))
                    pService = Trim(ds.Tables(0).Rows(xi)("TypeOfService"))

                    'If pTerm = "1" Or pTerm = "2" Then
                    '    pTerm = "FCA"
                    'ElseIf pTerm = "3" Or pTerm = "4" Then
                    '    pTerm = "CIF"
                    'ElseIf pTerm = "5" Then
                    '    pTerm = "DDU PASI"
                    'ElseIf pTerm = "6" Then
                    '    pTerm = "DDU Affiliate"
                    'ElseIf pTerm = "7" Then
                    '    pTerm = "EX-Work"
                    'ElseIf pTerm = "8" Then
                    '    pTerm = "FOB"
                    'End If

                    pTerm = uf_DesPrice(pTerm)

                    'PDF COMMERCIAL INVOICE
                    pFilePDF3 = CreateInvoiceEXToPDF(sh_Shippingno, sh_Affiliate, sh_ForwarderID, pTerm, pService)
                    If pFilePDF3 = "" Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & pFilePDF3 & " NG. " & vbCrLf & _
                        rtbProcess.Text
                        GoTo keluar
                    Else
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Create File " & pFilePDF3 & " OK. " & vbCrLf & _
                        rtbProcess.Text
                    End If
                    'PDF COMMERCIAL INVOICE

                    If sendEmailTallyInvoceShippingToForwarder("", "", pFilePDF3, sh_Shippingno, sh_Affiliate, sh_ForwarderID, "") = False Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & pFilePDF3 & " NG. " & vbCrLf & _
                        rtbProcess.Text
                        GoTo keluar
                    Else
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " [Export] Send E-Mail " & pFilePDF3 & " NG. " & vbCrLf & _
                        rtbProcess.Text
                    End If


                    Call UpdateShippingSendInvoiceCls(True, sh_Shippingno, sh_Affiliate, sh_ForwarderID)

                    Thread.Sleep(500)
keluar:
                Next
            Else
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Shipping Instruction STOPPED, because there is nothing PO to send " & vbCrLf & _
                    rtbProcess.Text
            End If
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Shipping Instruction STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally

        End Try
    End Sub

    Private Function CreateInvoiceEXToPDF(ByVal PshippingNo As String, ByVal pAffiliate As String, ByVal pForwarder As String, ByVal pTerm As String, ByVal pService As String) As String
        Dim CrReport As New Invoice()
        Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
        Dim pFile As String

        Try
            txtMsg.Text = ""
            Cursor.Current = Cursors.WaitCursor
            Dim dsPrint As New DataSet
            dsPrint = PrintInvoiceEX(PshippingNo, pAffiliate, pForwarder, pTerm, pService)

            If dsPrint.Tables(0).Rows.Count = 0 Then Exit Try

            CrReport.SetDatabaseLogon(cfg.User, cfg.Password, cfg.Server, cfg.Database)
            CrReport.SetDataSource(dsPrint.Tables(0))

            CrDiskFileDestinationOptions.DiskFileName = Trim(txtSaveAsDOM.Text) & "\Invoice-" & PshippingNo & "-" & pAffiliate.Trim & "-" & pForwarder.Trim & ".pdf"
            pFile = ""
            pFile = "Invoice-" & PshippingNo & "-" & pAffiliate.Trim & "-" & pForwarder.Trim & ".pdf"
            CrExportOptions = CrReport.ExportOptions

            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With

            Try
                CrReport.Export()
            Catch err As Exception
                MessageBox.Show(err.ToString())
            End Try
            'PDF
            CreateInvoiceEXToPDF = pFile
        Catch ex As Exception
            CreateInvoiceEXToPDF = ""
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Create PDF Invoice STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally
            If Not CrReport Is Nothing Then
                NAR(CrReport)
                GC.Collect()
            End If
            If Not CrReport Is Nothing Then
                CrReport.Dispose()
                CrReport.Close()
            End If
        End Try
    End Function

    Private Function PrintInvoiceEX(ByVal ls_value1 As String, ByVal ls_value2 As String, ByVal ls_value3 As String, ByVal pTerm As String, ByVal pService As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        Dim tentukanBoat As String = pService.ToString.Trim
        Dim tentukanTerm As String = pTerm.ToString.Trim

        Dim PriceCls As String = 0

        ''If tentukanBoat = "FCL" Or tentukanBoat = "LCL" Then
        'If tentukanTerm = "FCA" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
        '    PriceCls = "2"
        'ElseIf tentukanTerm = "FCA" Then
        '    PriceCls = "1"
        'End If

        'If tentukanTerm = "CIF" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
        '    PriceCls = "4"
        'ElseIf tentukanTerm = "CIF" Then
        '    PriceCls = "3"
        'End If

        'If tentukanTerm = "DDU PASI" Then
        '    PriceCls = "5"
        'ElseIf tentukanTerm = "DDU Affiliate" Then
        '    PriceCls = "6"
        'ElseIf tentukanTerm = "EX-Work" Then
        '    PriceCls = "7"
        'ElseIf tentukanTerm = "FOB" Then
        '    PriceCls = "8"
        'End If

        PriceCls = uf_PriceCls(tentukanTerm)

        ls_SQL = ""
        ls_SQL = "  select distinct  " & vbCrLf & _
              "  buyer = Rtrim(MA.BuyerName) + CHAR(13)+CHAR(10) + Rtrim(MA.BuyerAddress),  " & vbCrLf & _
              "  Consignee = Rtrim(Coalesce(MA.ConsigneeName, MA.AffiliateName)) + CHAR(13)+CHAR(10) + Rtrim(coalesce(MA.ConsigneeAddress, Rtrim(MA.Address) + Rtrim(MA.City) )),  " & vbCrLf & _
              "  Attn = ISNULL(ma.Att,''),  " & vbCrLf & _
              "  Vessel = ISNULL(SHM.Vessels,''), " & vbCrLf & _
              "  Fromto = Isnull(MF.Port,''),  " & vbCrLf & _
              "  Toto = CASE WHEN SHM.TypeOfService = 'AIR FREIGHT' THEN isnull(MA.DestinationPortAir,'') ELSE isnull(MA.DestinationPort,'') END,  " & vbCrLf & _
              "  About = Convert(Char(12), convert(Datetime, isnull(SHM.ETAPort,POM.ETAPort1)),106),  " & vbCrLf & _
              "  ONAbout = Convert(Char(12), convert(Datetime, isnull(SHM.ETDPort,POM.ETDPort1)),106),  " & vbCrLf & _
              "  Via = SHM.Via,  " & vbCrLf & _
              "  InvoiceNo = SHM.ShippingInstructionNo,   "

        ls_SQL = ls_SQL + "  OrderNo = (SELECT (STUFF((SELECT distinct ', ' + RTrim(ShippingInstruction_Detail.orderNo) FROM ShippingInstruction_Detail WHERE ShippingInstructionNo = '" & ls_value1 & "' AND AffiliateID = '" & ls_value2 & "' AND ForwarderID = '" & ls_value3 & "' FOR XML PATH('')), 1, 2, ''))),  " & vbCrLf & _
                          "  InvDate = Convert(Char(12), convert(Datetime, isnull(SHM.ShippingInstructionDate,'')),106),  " & vbCrLf & _
                          "  Place = 'JAKARTA',  " & vbCrLf & _
                          "  Privilege = '',  " & vbCrLf & _
                          "  AWB = '',  " & vbCrLf & _
                          "  ContainerNo = '', --TM.ContainerNo,  " & vbCrLf & _
                          "  Insurance = '',  " & vbCrLf & _
                          "  Remarks = '',  " & vbCrLf & _
                          "  paymentTerm = CASE WHEN POM.CommercialCls = '1' Then Isnull(MA.PaymentTerm,'') ELSE 'NO COMMERCIAL VALUE' END,  " & vbCrLf & _
                          "  Marks = '',--Description = '',  " & vbCrLf & _
                          "  QtyBox = SHD.QtyBox,   " & vbCrLf

        ls_SQL = ls_SQL + "  Qty = RB.Box,  " & vbCrLf & _
                          "  Price = isnull(SHD.Price,ISNULL(MPR.Price,0)),  " & vbCrLf & _
                          "  Amount = 0,  " & vbCrLf & _
                          "  Net =  (isnull(NetWeight,0)/1000),  " & vbCrLf & _
                          "  Gross =(isnull(SHM.GrossWeight,0)/1000),  " & vbCrLf & _
                          "  DocNo = '',  " & vbCrLf & _
                          "  RevNo = '',  " & vbCrLf & _
                          "  partCust = isnull(PartGroupName,''),  " & vbCrLf & _
                          "  PartYazaki = SHD.PartNo,  " & vbCrLf & _
                          "  CaseNo = Rtrim(RB.Label1) + '-' + Rtrim(RB.Label2),  " & vbCrLf & _
                          "  totalCarton = 0 , " & vbCrLf & _
                          "  Term = RTRIM(MPC.Description) /*CASE WHEN RTRIM(MPC.Description) = 'FCA - BOAT' THEN 'FCA' " & vbCrLf & _
                          "  WHEN RTRIM(MPC.Description) = 'FCA - AIR' THEN 'FCA'	 " & vbCrLf & _
                          "  WHEN RTRIM(MPC.Description) = 'CIF - BOAT' THEN 'CIF' " & vbCrLf & _
                          "  WHEN RTRIM(MPC.Description) = 'CIF - AIR' THEN 'CIF' " & vbCrLf & _
                          "  ELSE RTRIM(MPC.Description) END*/,	" & vbCrLf & _
                          "  SHM.TotalPallet, SHM.Measurement, SHM.GrossWeight, SHM.Freight, CASE WHEN ISNULL(HSCodeCls,'0')  = '0'  THEN '' else HSCode END HSCode, SHD.OrderNo NewOrderNo, SHM.TypeOfService,  SHM.NamaKapalS  From  " & vbCrLf

        ls_SQL = ls_SQL + "  ShippingInstruction_Detail SHD   " & vbCrLf & _
                          "  LEFT JOIN ShippingInstruction_Master SHM ON SHM.ShippingInstructionNo = SHD.ShippingInstructionNo  " & vbCrLf & _
                          "  AND SHM.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  AND SHM.ForwarderID = SHD.ForwarderID  " & vbCrLf & _
                          "  LEFT JOIN Tally_Master TM ON TM.ShippingInstructionNo = SHM.ShippingInstructionNo and TM.AffiliateID = SHM.AffiliateID and TM.ForwarderID = SHM.ForwarderID " & vbCrLf & _
                          "  LEFT JOIN MS_Parts MP ON MP.PartNo = SHD.PartNo  " & vbCrLf & _
                          "  LEFT JOIN MS_PartMapping MPM ON MPM.Partno = SHD.PartNo and MPM.AffiliateID = SHD.AffiliateID and MPM.SupplierID = SHD.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SHD.ForwarderID  LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_master RM ON RM.SuratJalanNo = SHD.SuratJalanno  AND RM.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  AND RM.OrderNo = SHD.OrderNo  " & vbCrLf & _
                          "  AND SHD.SupplierID = RM.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_Detail RD ON RD.SuratJalanNO = RM.SuratJalanno  " & vbCrLf

        ls_SQL = ls_SQL + "  AND RD.AffiliateID = RM.AffiliateID  " & vbCrLf & _
                          "  AND RD.SupplierID = RM.SupplierID  " & vbCrLf & _
                          "  AND RD.PONo = RM.PONO  " & vbCrLf & _
                          "  AND RD.OrderNO = Rm.OrderNo  " & vbCrLf & _
                          "  AND RD.PartNo = SHD.PartNo " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = SHD.SuratJalanNo  " & vbCrLf & _
                          "  AND RB.SupplierID = SHD.SupplierID   " & vbCrLf & _
                          "  AND RB.AffiliateID = SHD.AffiliateID   " & vbCrLf & _
                          "  --AND RB.PONo = RD.PONo   " & vbCrLf & _
                          "  AND RB.OrderNo = SHD.OrderNo   " & vbCrLf & _
                          "  AND RB.PartNo = SHD.PartNo   " & vbCrLf

        ls_SQL = ls_SQL + "  AND RB.StatusDefect = '0'   " & vbCrLf & _
                          "  LEFT JOIN PO_Detail_Export POD ON POD.PONo = RD.PONO  " & vbCrLf & _
                          "  AND POD.OrderNo1 = RD.OrderNo  " & vbCrLf & _
                          "  AND POD.AffiliateID = RD.AffiliateID  AND POD.SupplierID = RD.SupplierID  " & vbCrLf & _
                          "  AND POD.PartNO = RD.PartNo  " & vbCrLf & _
                          "  LEFT JOIN PO_Master_export POM ON POM.PONo = POD.PONO  " & vbCrLf & _
                          "  AND POM.OrderNo1 = POD.OrderNo1  " & vbCrLf & _
                          "  AND POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "  AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN MS_PriceCls MPC ON MPC.PriceCls = SHM.TermDelivery " & vbCrLf & _
                          "  LEFT JOIN MS_Price MPR ON MPR.PartNO = SHD.PartNo  " & vbCrLf & _
                          "  AND MPR.AffiliateID = SHD.AffiliateID  " & vbCrLf

        ls_SQL = ls_SQL + "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SHM.ETDPort,'')), 112) >= CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.EffectiveDate,'')), 112)  " & vbCrLf & _
                          "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SHM.ETDPort,'')), 112) between CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Startdate,'')), 112) and CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Enddate,'')), 112)  " & vbCrLf & _
                          "  AND MPR.CurrCls = '02' AND MPR.PriceCls = '" & PriceCls & "'" & vbCrLf & _
                          "  WHERE SHM.ShippingInstructionNo = '" & Trim(ls_value1) & "'  " & vbCrLf & _
                          "  AND SHM.AffiliateID = '" & Trim(ls_value2) & "' " & vbCrLf & _
                          "  AND SHM.ForwarderID = '" & Trim(ls_value3) & "'  order by SHD.partno, Rtrim(RB.Label1) + '-' + Rtrim(Rb.Label2)  "

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Sub UpdateShippingSendInvoiceCls(ByVal pIsNewData As Boolean, _
                        Optional ByVal pShippno As String = "", _
                        Optional ByVal pAFF As String = "", _
                        Optional ByVal pFWD As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & vbCrLf & _
                      " SET sendInvoice='2'" & vbCrLf & _
                      " WHERE ShippingInstructionNo='" & Trim(pShippno) & "'  " & vbCrLf & _
                      " AND AffiliateID='" & Trim(pAFF) & "' " & vbCrLf & _
                      " AND ForwarderID='" & Trim(pFWD) & "' " & vbCrLf
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO to Supplier STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
        End Try

    End Sub

    Private Sub pGetExcelSTOCKOPNAME()
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""
        Dim irow As Integer = 0
        Dim ls_Affiliate As String = ""

        'copy file from server to local
        Dim NewFileCopy As String

        'MdlConn.ReadConnection()
        Dim ls_SQL As String = ""
        Dim ds As New DataSet
        Dim dsDetail As New DataSet
        Dim ls_Fwd As String = "", ls_ReqDate As Date, ls_PONO As String = "", ls_PartNo As String = "", ls_email As String = ""
        Dim xlApp = New Excel.Application

        Try
            ls_SQL = " Select * from ForwarderStockOpname_Request where SendExcel = '1' "

            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pDate = Now
                    ls_Fwd = Trim(ds.Tables(0).Rows(xi)("ForwarderID"))
                    ls_ReqDate = Format((ds.Tables(0).Rows(xi)("ReqDate")), "dd MMM yyyy")
                    ls_PONO = Trim(ds.Tables(0).Rows(xi)("OrderNo"))
                    ls_PartNo = Trim(ds.Tables(0).Rows(xi)("PartNo"))
                    ls_email = Trim(ds.Tables(0).Rows(xi)("EmailFrom"))

                    dsDetail = getSTOCKOPNAME(ls_Fwd, ls_PONO, ls_PartNo)

                    'Create Excel File
                    Dim fi As New FileInfo(Trim(txtAttachmentDOM.Text) & "\Template Stock Opname.xlsx") 'File dari Local
                    If Not fi.Exists Then
                        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Stock Opname STOPPED, because file Excel isn't Found" & vbCrLf & _
                                        rtbProcess.Text
                        Exit Sub
                    End If

                    NewFileCopy = Trim(txtAttachmentDOM.Text) & "\Template Stock Opname.xlsx"
                    Dim ls_file As String = NewFileCopy
                    ExcelBook = xlApp.Workbooks.Open(ls_file)
                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                    receiptEmail = ls_email
                    receiptCCEmail = ""

                    ExcelSheet.Range("D3").Value = Trim(ls_Fwd)
                    ExcelSheet.Range("D4").Value = Format(ls_ReqDate, "dd MMM yyyy")
                    ExcelSheet.Range("D4").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    irow = 7

                    If dsDetail.Tables(0).Rows.Count > 0 Then
                        For i = 0 To dsDetail.Tables(0).Rows.Count - 1

                            If i = 0 Then
                                irow = 7
                                ls_Affiliate = Trim(dsDetail.Tables(0).Rows(i)("AffiliateID"))
                            Else
                                If ls_Affiliate <> Trim(dsDetail.Tables(0).Rows(i)("AffiliateID")) Then
                                    irow = irow + 2
                                End If
                            End If

                            If ls_Affiliate <> Trim(dsDetail.Tables(0).Rows(i)("AffiliateID")) Or i = 0 Then
                                ExcelSheet.Range("B" & i + irow - 1 & ": C" & i + irow - 1).MergeCells = True
                                ExcelSheet.Range("B" & i + irow - 1).Value = Trim(dsDetail.Tables(0).Rows(i)("AffiliateID"))
                                ExcelSheet.Range("B" & i + irow).Value = "NO."
                                ExcelSheet.Range("C" & i + irow).Value = "ORDER NO"
                                ExcelSheet.Range("D" & i + irow).Value = "PART NO."
                                ExcelSheet.Range("E" & i + irow).Value = "PART NAME."
                                ExcelSheet.Range("F" & i + irow).Value = "BOX NO."
                                ExcelSheet.Range("G" & i + irow).Value = "SURAT JALAN NO."
                                ExcelSheet.Range("H" & i + irow).Value = "UOM"
                                ExcelSheet.Range("I" & i + irow).Value = "QTY/BOX"
                                ExcelSheet.Range("J" & i + irow).Value = "GOOD REC QTY"
                                ExcelSheet.Range("K" & i + irow).Value = "BOX QTY"
                                ExcelSheet.Range("L" & i + irow).Value = "SUPPLIER ID"
                                ExcelSheet.Range("M" & i + irow).Value = "SUPPLIER NAME"
                                ls_Affiliate = Trim(dsDetail.Tables(0).Rows(i)("AffiliateID"))
                                DrawAllBorders(ExcelSheet.Range("B" & i + irow - 1 & ": C" & i + irow - 1))
                                DrawAllBorders(ExcelSheet.Range("B" & i + irow & ": M" & i + irow))
                                ExcelSheet.Range("B" & i + irow & ": M" & i + irow).Interior.Color = Color.Gray
                                ExcelSheet.Range("B" & i + irow - 1 & ": B" & i + irow - 1).Interior.Color = Color.Gray

                                irow = irow + 1
                            End If

                            'Detail
                            ExcelSheet.Range("B" & i + irow).Value = i + 1
                            ExcelSheet.Range("B" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                            ExcelSheet.Range("C" & i + irow).Value = Trim(dsDetail.Tables(0).Rows(i)("Orderno"))
                            ExcelSheet.Range("C" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("D" & i + irow).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                            ExcelSheet.Range("D" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("E" & i + irow).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName"))
                            ExcelSheet.Range("E" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("F" & i + irow).Value = Trim(dsDetail.Tables(0).Rows(i)("BoxNo"))
                            ExcelSheet.Range("F" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("G" & i + irow).Value = Trim(dsDetail.Tables(0).Rows(i)("SJNo"))
                            ExcelSheet.Range("G" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("H" & i + irow).Value = Trim(dsDetail.Tables(0).Rows(i)("UOM"))
                            ExcelSheet.Range("H" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("I" & i + irow).Value = Trim(dsDetail.Tables(0).Rows(i)("Qtybox"))
                            ExcelSheet.Range("I" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("J" & i + irow).Value = (dsDetail.Tables(0).Rows(i)("BoxQty") * dsDetail.Tables(0).Rows(i)("QtyBox"))
                            ExcelSheet.Range("J" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("K" & i + irow).Value = Trim(dsDetail.Tables(0).Rows(i)("BoxQty"))
                            ExcelSheet.Range("K" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("L" & i + irow).Value = dsDetail.Tables(0).Rows(i)("SupplierID")
                            ExcelSheet.Range("L" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            ExcelSheet.Range("M" & i + irow).Value = dsDetail.Tables(0).Rows(i)("SupplierName")
                            ExcelSheet.Range("M" & i + irow).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                            DrawAllBorders(ExcelSheet.Range("B" & i + irow & ": M" & i + irow))
                        Next
                    End If

                    xlApp.DisplayAlerts = False

                    Dim temp_Filename As String = "Stock Opname " & Trim(ls_Fwd) & "-" & Format(ls_ReqDate, "dd MMM yyyy") & ".xlsx"
                    ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\" & temp_Filename)
                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    If sendEmailStockOpnameToForwarder(temp_Filename, ls_Fwd, ls_ReqDate, ls_email) = False Then GoTo keluar

                    Call UpdateForwarderStockOpname_Request(True, ls_ReqDate, ls_Fwd)

                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    Thread.Sleep(500)
keluar:
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                Next
            Else
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send STOCK OPNAME STOPPED, because there is nothing PO to send " & vbCrLf & _
                    rtbProcess.Text
            End If
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send STOCK OPNAME STOPPED, because " & ex.Message & " " & vbCrLf & _
                        rtbProcess.Text
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Function getSTOCKOPNAME(ByVal pFwd As String, ByVal pPONo As String, ByVal pPartNo As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = " select distinct " & vbCrLf & _
                  " OrderNo = RM.OrderNo, " & vbCrLf & _
                  " AffiliateID = RM.AffiliateID, " & vbCrLf & _
                  " AffiliateName = MA.AffiliateName, " & vbCrLf & _
                  " SupplierID = RM.SupplierID, " & vbCrLf & _
                  " SupplierName = MS.SupplierName, " & vbCrLf & _
                  " PartNo = RD.PartNo, " & vbCrLf & _
                  " PartName = MP.PartName, " & vbCrLf & _
                  " UOM = isnull(UC.Description,''), " & vbCrLf & _
                  " QtyBox = ISNULL(PDE.POQtyBox,PMP.QtyBox), " & vbCrLf & _
                  " BoxNo = isnull(Rtrim(RB.label1) + '-' + Rtrim(RB.label2),''), " & vbCrLf & _
                  " SJNo = RM.SuratJalanNo, "

        ls_SQL = ls_SQL + " BoxQty = isnull(RB.box,0)  " & vbCrLf & _
                          " FROM dbo.ReceiveForwarder_Master RM   " & vbCrLf & _
                          " LEFT JOIN dbo.ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratJalanNo   " & vbCrLf & _
                          " 	AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                          " 	AND RM.AffiliateID = RD.AffiliateID   " & vbCrLf & _
                          " 	AND RM.PONo = RD.PONo   " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Parts MP ON RD.PartNo = MP.PartNo   " & vbCrLf & _
                          " LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo  " & vbCrLf & _
                          " 	AND RB.SupplierID = RD.SupplierID   " & vbCrLf & _
                          " 	AND RB.AffiliateID = RD.AffiliateID   " & vbCrLf & _
                          " 	AND RB.PONo = RD.PONo   "

        ls_SQL = ls_SQL + " 	AND RB.OrderNo = RD.OrderNo   " & vbCrLf & _
                          " 	AND RB.PartNo = RD.PartNo   " & vbCrLf & _
                          " 	AND RB.StatusDefect = '0'   " & vbCrLf & _
                          " LEFT JOIN dbo.MS_PartMapping PMP ON RD.PartNo = PMP.PartNo and RM.AffiliateID = PMP.AffiliateID and RM.SupplierID = PMP.SupplierID   " & vbCrLf & _
                          " LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Supplier MSS ON RM.SupplierID = MSS.SupplierID   " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = RM.AffiliateID " & vbCrLf & _
                          " LEFT JOIN dbo.PO_Master_Export PME ON PME.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                          " 	AND (RM.OrderNo =  PME.OrderNo1 or RM.OrderNo =  PME.OrderNo2  " & vbCrLf & _
                          " 	or RM.OrderNo =  PME.OrderNo3 or RM.OrderNo =  PME.OrderNo4 or RM.OrderNo =  PME.OrderNo5)  " & vbCrLf & _
                          " LEFT JOIN dbo.PO_Detail_Export PDE on RB.PONo = PDE.PONo and RB.AffiliateID = PDE.AffiliateID and RB.SupplierID = PDE.SupplierID and RB.PartNo = PDE.PartNo  " & vbCrLf & _
                          " LEFT JOIN MS_Supplier MS ON MS.SupplierID = RM.SupplierID " & vbCrLf & _
                          " WHERE ISNULL(RD.OrderNo, '') <> ''    "

        ls_SQL = ls_SQL + " and RTrim(RD.SuratJalanNo) + Rtrim(RD.AffiliateID) + Rtrim(RD.OrderNo)+RTRIM(RD.SupplierID)+RTRIM(RD.PartNo)  " & vbCrLf & _
                          " 	NOT IN (SELECT DISTINCT RTrim(SuratJalanNo) + Rtrim(AffiliateID) + Rtrim(OrderNo)+RTRIM(SupplierID)+RTRIM(PartNo)  " & vbCrLf & _
                          " 			From ShippingInstruction_Detail where  " & vbCrLf & _
                          " 			suratjalanno = RD.SuratJalanNo and AffiliateID = RD.AffiliateID AND SupplierID = RD.SupplierID and partno = RD.Partno and orderno = RD.OrderNo)  " & vbCrLf & _
                          " AND RM.ForwarderID = '" & Trim(pFwd) & "' " & vbCrLf

        If pPONo <> "" Then
            ls_SQL = ls_SQL + " AND RD.OrderNo = '" & Trim(pPONo) & "'" & vbCrLf
        End If

        If pPartNo <> "" Then
            ls_SQL = ls_SQL + " AND RD.PartNo = '" & Trim(pPartNo) & "'" & vbCrLf
        End If

        ls_SQL = ls_SQL + " Order By RM.AffiliateID, RM.orderNo,RD.partNo"

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Private Function sendEmailStockOpnameToForwarder(ByVal pFileName As String, ByVal sForwarder As String, ByVal sReqDate As String, ByVal sEmail As String) As Boolean
        Try
            Dim TempFilePath As String
            Dim TempFileName As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)
            TempFileName = "\" & pFileName

            Dim dsEmail As New DataSet
            dsEmail = EmailSH("", "PASI", sForwarder)
            'To Supplier, CC Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                    receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("SupplierDeliveryCC")
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Trim(sEmail)
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Stock Opname STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailStockOpnameToForwarder = False
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Stock Opname STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailStockOpnameToForwarder = False
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            '            mailMessage.Subject = "[TRIAL] TALLY DATA: " & shaffiliate.Trim & "-" & shshipno.Trim & ""
            mailMessage.Subject = "STOCK OPNAME-" & sForwarder.Trim & "-" & sReqDate.Trim & " "
            'TA-AFF-Invoice No. Tally Data [TRIAL]
            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If
            GetSettingEmail_Export("PO")
            ls_Body = clsNotification.GetNotification("25", "", sForwarder.Trim)
            mailMessage.Body = ls_Body

            Dim filename As String = TempFilePath & TempFileName
            mailMessage.Attachments.Add(New Attachment(filename))
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient

            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)

            sendEmailStockOpnameToForwarder = True
            'Delete the file
            'Kill(TempFilePath & TempFileName)
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Stock Opname to Forwarder SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
            Exit Function
        Catch ex As Exception
            sendEmailStockOpnameToForwarder = False
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Stock Opname to Forwarder STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text

        End Try

    End Function

    Private Sub UpdateForwarderStockOpname_Request(ByVal pIsNewData As Boolean, _
                         ByVal pReqDate As Date, _
                         Optional ByVal pFWD As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.ForwarderStockOpname_Request " & vbCrLf & _
                      " SET SendExcel='2'" & vbCrLf & _
                      " WHERE ForwarderID='" & pFWD & "'  " & vbCrLf & _
                      " AND ReqDate='" & Format(pReqDate, "yyyy-MM-dd") & "' " & vbCrLf
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Stock Opname STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
        End Try

    End Sub

    Private Sub Excel_CancelationList()
        Dim strFileSize As String = ""
        Dim ls_sql As String = ""

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim sheetNumber As Integer = 1
        Dim i As Integer, xi As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        Dim fileTocopy As String
        Dim NewFileCopy As String
        Dim NewFileCopyas As String

        Dim KTime1 As String = ""
        Dim KTime2 As String = ""
        Dim KTime3 As String = ""
        Dim KTime4 As String = ""
        Dim pNamaFile As String = ""

        Dim jkanbanno As String
        Dim jQty As Long
        Dim jQtyBox As Long
        Dim jQtyPallet As Long

        Dim ds As New DataSet
        Dim dsHeader As New DataSet
        Dim dsDetail As New DataSet
        Dim dsETAETD As New DataSet
        Dim dsDetailDelivery As New DataSet

        Dim ls_SJ As String = ""
        Dim ls_Supplier As String = ""
        Dim ls_supplierName As String = ""
        Dim ls_supplierAdd As String = ""
        Dim ls_delivery As String = ""
        Dim ls_DeliveryName As String = ""
        Dim ls_deliveryAdd As String = ""
        Dim ls_orderNo As String = ""
        Dim ls_orderNo1 As String = ""
        Dim ls_ETDV As Date
        Dim ls_ETDP As Date
        Dim ls_ETAP As Date
        Dim ls_ETAF As Date
        Dim ls_Aff As String = ""
        Dim ls_AFFName As String = ""
        Dim ls_AffADD As String
        Dim ls_Attn As String = ""
        Dim ls_telp As String = ""
        Dim ls_StatusCancel As String = ""

        Dim i_loop As Long
        Dim xlApp = New Excel.Application

        Try
            ls_sql = "SELECT DISTINCT attn = ISNULL(MF.Attn,''), telp = ISNULL(MF.MobilePhone,''), period = PME.Period, ISNULL(DOM.SuratJalanNo, '') SuratJalanNo, PME.PONo, PME.OrderNo1 AS orderNo, PME.AffiliateID, " & vbCrLf & _
                "PME.SupplierID, MS.SupplierName, ISNULL(MS.Address,'') + ' ' + ISNULL(MS.City,'') + ' ' + ISNULL(MS.Postalcode,'') SUPPAddress, " & vbCrLf & _
                "PME.ForwarderID, MF.ForwarderName, ISNULL(MF.Address,'')  + ' ' + ISNULL(MF.City,'') + ' ' + ISNULL(MF.PostalCode,'') AS FWDAddress, " & vbCrLf & _
                "ETDVendor1 AS ETDVendor, ETDPort1 AS ETDPort, ETAPort1 AS ETAPort, ETAFactory1 AS ETAFactory, " & vbCrLf & _
                "PME.AffiliateID AS AFF, MA.AffiliateName AS AFFName, ISNULL(MA.Address,'')  + ' ' + ISNULL(MA.City,'') + ' ' + ISNULL(MA.PostalCode,'') AS AFFAddress, " & vbCrLf & _
                "CASE WHEN PME.ShipCls = 'A' THEN 'AIR' ELSE 'BOAT' END ShipCls, SplitStatus " & vbCrLf & _
                "FROM PO_Master_ExportCancel PME " & vbCrLf & _
                "LEFT JOIN DOSUpplier_Master_Export DOM ON DOM.PONo = PME.PONo AND DOM.SupplierID = PME.SupplierID AND DOM.AffiliateID = PME.AffiliateID AND DOM.SplitReffPONo = PME.OrderNo1 " & vbCrLf & _
                "LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                "LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID " & vbCrLf & _
                "LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                "WHERE ISNULL(PME.ExcelCls, '0') = '1' "

            'MdlConn.ReadConnection()
            ds = cls.uf_GetDataSet(ls_sql)

            Dim ls_file As String

            For i_loop = 0 To ds.Tables(0).Rows.Count - 1

                Dim fi3 As New FileInfo(Trim(txtAttachmentDOM.Text) & "\PO CANCELATION LIST.xlsx")

                If Not fi3.Exists Then
                    rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery Note STOPPED, because File Excel isn't Found " & vbCrLf & _
                                    rtbProcess.Text
                    Exit Sub
                End If

                pAffCode = Trim(ds.Tables(0).Rows(i_loop)("AFF"))
                pSupplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))

                ls_SJ = Trim(ds.Tables(0).Rows(i_loop)("SuratJalanNo"))
                ls_Supplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))
                ls_supplierName = Trim(ds.Tables(0).Rows(i_loop)("suppliername"))
                ls_supplierAdd = Trim(ds.Tables(0).Rows(i_loop)("SuppAddress"))
                ls_delivery = Trim(ds.Tables(0).Rows(i_loop)("ForwarderID"))
                ls_DeliveryName = Trim(ds.Tables(0).Rows(i_loop)("ForwarderName"))
                ls_deliveryAdd = Trim(ds.Tables(0).Rows(i_loop)("FWDAddress"))
                ls_orderNo = Trim(ds.Tables(0).Rows(i_loop)("PONo"))
                ls_orderNo1 = Trim(ds.Tables(0).Rows(i_loop)("OrderNo"))
                ls_ETDV = Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "yyyy-MM-dd")
                ls_ETDP = Format((ds.Tables(0).Rows(i_loop)("ETDPort")), "yyyy-MM-dd")
                ls_ETAP = Format((ds.Tables(0).Rows(i_loop)("ETAPort")), "yyyy-MM-dd")
                ls_ETAF = Format((ds.Tables(0).Rows(i_loop)("ETAFactory")), "yyyy-MM-dd")
                ls_Aff = Trim(ds.Tables(0).Rows(i_loop)("AFF"))
                ls_AFFName = Trim(ds.Tables(0).Rows(i_loop)("AFFName"))
                ls_AffADD = Trim(ds.Tables(0).Rows(i_loop)("AFFAddress"))
                ls_Attn = Trim(ds.Tables(0).Rows(i_loop)("attn"))
                ls_telp = Trim(ds.Tables(0).Rows(i_loop)("telp"))
                ls_StatusCancel = Trim(ds.Tables(0).Rows(i_loop)("SplitStatus"))

                dsDetailDelivery = BindDataPOCancelationList(ls_SJ, ls_orderNo, ls_orderNo1, ls_Aff, ls_Supplier)

                If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                    Dim k As Long
                    Dim dsAffiliate As New DataSet
                    dsAffiliate = Affiliate(Trim(ls_Aff))

                    Dim dsSupplier As New DataSet
                    dsSupplier = Supplier(Trim(ls_Supplier))

                    Dim status As Boolean
                    status = True

                    If dsDetailDelivery.Tables(0).Rows.Count = 0 Then
                        status = False
                    Else
                        status = True
                    End If

                    If status = True Then
                        NewFileCopy = Trim(txtAttachmentDOM.Text) & "\PO CANCELATION LIST.xlsx"
                        ls_file = NewFileCopy
                        ExcelBook = xlApp.Workbooks.Open(ls_file)
                        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                        ExcelSheet.Range("H1").Value = "POCL"
                        ExcelSheet.Range("H2").Value = receiptEmail
                        ExcelSheet.Range("H3").Value = ls_Aff
                        ExcelSheet.Range("H4").Value = ls_delivery
                        ExcelSheet.Range("H5").Value = ls_Supplier

                        ExcelSheet.Range("AE11:AT11").Merge()
                        ExcelSheet.Range("AE11:AT11").Value = ls_supplierName
                        ExcelSheet.Range("AE12:AT15").Merge()
                        ExcelSheet.Range("AE12:AT15").Value = ls_supplierAdd
                        ExcelSheet.Range("AE19:AT19").Merge()
                        ExcelSheet.Range("AE19:AT19").Value = ls_DeliveryName
                        ExcelSheet.Range("AE20:AT22").Merge()
                        ExcelSheet.Range("AE20:AT22").Value = ls_deliveryAdd
                        ExcelSheet.Range("I28:P28").Merge()

                        ExcelSheet.Range("G11:K11").Value = Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy-MM")

                        ExcelSheet.Range("G13:K13").Value = ls_orderNo
                        ExcelSheet.Range("G15:K15").Value = ls_orderNo1

                        ExcelSheet.Range("G17:K17").Value = ds.Tables(0).Rows(i_loop)("ShipCls")

                        ExcelSheet.Range("R11:V11").Merge()
                        ExcelSheet.Range("R11:V11").Value = ls_ETDV
                        ExcelSheet.Range("R13:V13").Merge()
                        ExcelSheet.Range("R13:V13").Value = ls_ETDP
                        ExcelSheet.Range("R15:V15").Merge()
                        ExcelSheet.Range("R15:V15").Value = ls_ETAP
                        ExcelSheet.Range("R17:V17").Merge()
                        ExcelSheet.Range("R17:V17").Value = ls_ETAF

                        ExcelSheet.Range("G19:V19").Merge()
                        ExcelSheet.Range("G19:V19").Value = ls_AFFName

                        ExcelSheet.Range("G20:V22").Merge()
                        ExcelSheet.Range("G20:V22").Value = ls_AffADD
                        k = 0

                        For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                            'For i = 0 To 3
                            k = k
                            ExcelSheet.Range("B" & k + 27 & ": C" & k + 27).Merge()
                            ExcelSheet.Range("D" & k + 27 & ": H" & k + 27).Merge()
                            ExcelSheet.Range("I" & k + 27 & ": P" & k + 27).Merge()
                            ExcelSheet.Range("Q" & k + 27 & ": R" & k + 27).Merge()
                            ExcelSheet.Range("S" & k + 27 & ": T" & k + 27).Merge()
                            ExcelSheet.Range("U" & k + 27 & ": X" & k + 27).Merge()
                            ExcelSheet.Range("Y" & k + 27 & ": AB" & k + 27).Merge()
                            ExcelSheet.Range("AC" & k + 27 & ": AF" & k + 27).Merge()
                            ExcelSheet.Range("AG" & k + 27 & ": AJ" & k + 27).Merge()
                            ExcelSheet.Range("AK" & k + 27 & ": AN" & k + 27).Merge()

                            ExcelSheet.Range("B" & k + 27 & ": C" & k + 27).Value = k + 1
                            ExcelSheet.Range("D" & k + 27 & ": H" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partno"))
                            ExcelSheet.Range("I" & k + 27 & ": P" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Partname"))

                            ExcelSheet.Range("Q" & k + 27 & ": R" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("uom")
                            ExcelSheet.Range("S" & k + 27 & ": T" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("qtybox")

                            ExcelSheet.Range("U" & k + 27 & ": X" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("poqty")
                            ExcelSheet.Range("Y" & k + 27 & ": AB" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("qty")
                            ExcelSheet.Range("AC" & k + 27 & ": AF" & k + 27).Value = dsDetailDelivery.Tables(0).Rows(j)("BoxQty")

                            ExcelSheet.Range("AG" & k + 27 & ": AJ" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo1"))
                            ExcelSheet.Range("AK" & k + 27 & ": AN" & k + 27).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo2"))

                            k = k + 1
                        Next

                        DrawAllBorders(ExcelSheet.Range("B27" & ": AN" & k + 26))

                        'Save ke Local
                        xlApp.DisplayAlerts = False

                        If Trim(ls_SJ) = "" Then
                            ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\PO Cancelation List-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo1) & ".xlsx")
                            pNamaFile = "\PO Cancelation List-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo1) & ".xlsx"
                        Else
                            ExcelBook.SaveAs(Trim(txtSaveAsDOM.Text) & "\PO Cancelation List-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo1) & "-" & Trim(ls_SJ) & ".xlsx")
                            pNamaFile = "\PO Cancelation List-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo1) & "-" & Trim(ls_SJ) & ".xlsx"
                        End If

                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If

                    Select Case ls_StatusCancel
                        Case "2", "3", "4"
                            If sendEmailPOCancel(pNamaFile, pAffCode, ls_delivery, ls_orderNo, ls_orderNo1, "Supplier") = False Then GoTo keluar

                            If ls_StatusCancel = "4" Then
                                If sendEmailPOCancel(pNamaFile, pAffCode, ls_delivery, ls_orderNo, ls_orderNo1, "Forwarder") = False Then GoTo keluar
                                If sendEmailPOCancel(pNamaFile, pAffCode, ls_delivery, ls_orderNo, ls_orderNo1, "Affiliate") = False Then GoTo keluar
                            End If
                        Case "5", "6"
                            If sendEmailPOCancel(pNamaFile, pAffCode, ls_delivery, ls_orderNo, ls_orderNo1, "Forwarder") = False Then GoTo keluar
                            If sendEmailPOCancel(pNamaFile, pAffCode, ls_delivery, ls_orderNo, ls_orderNo1, "Affiliate") = False Then GoTo keluar
                    End Select

                    Call UpdateStatusPOCancelation(ls_Aff, ls_Supplier, ls_orderNo, ls_orderNo1, ls_SJ)
                End If

keluar:
                xlApp.Workbooks.Close()
                xlApp.Quit()
            Next
            Exit Sub
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Delivery to Forwarder STOPPED, because " & Err.Description & " " & vbCrLf & _
                                    rtbProcess.Text
            xlApp.Workbooks.Close()
            xlApp.Quit()
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Function BindDataPOCancelationList(ByVal pSJ As String, ByVal PONO As String, ByVal ls_orderNo As String, ByVal Aff As String, ByVal Supp As String)
        Dim ls_sql As String
        'MdlConn.ReadConnection()

        ls_sql = "SELECT DISTINCT " & vbCrLf & _
            "orderno = POD.PONo, " & vbCrLf & _
            "Partno = POD.PartNo, " & vbCrLf & _
            "Partname = MP.PartName, " & vbCrLf & _
            "labelno1 = RTRIM(ISNULL(PL1.LabelNo, '')), " & vbCrLf & _
            "labelno2 = RTRIM(ISNULL(PL2.LabelNo, '')), " & vbCrLf & _
            "uom = MU.Description, " & vbCrLf & _
            "qtybox = ISNULL(DOD.POQtyBox,MPM.QtyBox), " & vbCrLf & _
            "poqty = CONVERT(CHAR, ISNULL(POD.SplitReffQty, 0)), " & vbCrLf & _
            "qty = CONVERT(CHAR, ISNULL(DOD.DOQty, POD.Week1)), " & vbCrLf & _
            "remaining = CONVERT(CHAR, POD.Week1), " & vbCrLf & _
            "DeliveryQty = CONVERT(CHAR, ISNULL(DOD.DOQty, POD.Week1)), " & vbCrLf & _
            "boxqty = CEILING(ISNULL(DOD.DOQty, POD.Week1) / ISNULL(DOD.POQtyBox,MPM.QtyBox)), " & vbCrLf & _
            "weight = NetWeight, " & vbCrLf & _
            "barcode = CONVERT(CHAR(25), '') + CONVERT(CHAR(20), POD.AFfiliateID) + CONVERT(CHAR(20), POD.Pono) + CONVERT(CHAR(25), POD.PartNo) + CONVERT(CHAR, ISNULL(DOD.DOQty, POD.Week1)) " & vbCrLf & _
            "From PO_Master_ExportCancel POM " & vbCrLf & _
            "LEFT JOIN PO_Detail_ExportCancel POD ON POM.Pono = POD.PONo AND POM.OrderNo1 = POD.OrderNo1 AND POM.AffiliateID = POD.AffiliateID And POM.SupplierID = POD.SupplierID " & vbCrLf & _
            "LEFT JOIN DOSupplier_Master_Export DOM ON DOM.SupplierID = POM.SupplierID AND DOM.AffiliateID = POM.AffiliateID AND DOM.PONo = POM.PONo AND DOM.ORderNo = POM.OrderNo1 " & vbCrLf & _
            "LEFT JOIN DOSupplier_Detail_Export DOD ON DOD.SuratJalanNo = DOM.SuratJalanNo AND DOD.AffiliateID = DOM.AffiliateID AND DOD.SupplierID = DOM.SupplierID AND DOD.PONo = DOM.POno AND DOD.OrderNo = DOM.OrderNo AND DOD.PartNo = POD.PartNo " & vbCrLf & _
            "LEFT JOIN ( " & vbCrLf & _
            "   SELECT OrderNo, SuratJalanno, POno, AffiliateID, SupplierID, PartNo, MIN(BoxNo) AS labelno " & vbCrLf & _
            "   FROM DOSupplier_DetailBox_Export " & vbCrLf & _
            "   GROUP BY OrderNo, SuratJalanno, POno, AffiliateID, SupplierID, PartNo " & vbCrLf & _
            ")PL1 ON PL1.PONo = POD.PONo AND PL1.AffiliateID = POD.AffiliateID AND PL1.SupplierID = POD.SupplierID AND PL1.PartNO = POD.PartNo AND PL1.SuratJalanno = DOM.SuratJalanno AND PL1.OrderNo = DOD.OrderNo " & vbCrLf & _
            "LEFT JOIN ( " & vbCrLf & _
            "   SELECT OrderNo, SuratJalanno, POno, AffiliateID, SupplierID, PartNo, Max(BoxNo) AS labelno " & vbCrLf & _
            "   FROM DOSupplier_DetailBox_Export " & vbCrLf & _
            "   GROUP BY OrderNo, SuratJalanno, POno, AffiliateID, SupplierID, PartNo " & vbCrLf & _
            ")PL2 ON PL2.PONo = POD.PONo AND PL2.AffiliateID = POD.AffiliateID AND PL2.SupplierID = POD.SupplierID AND PL2.PartNO = POD.PartNo AND PL2.SuratJalanno = DOM.SuratJalanno AND PL2.OrderNo = DOD.OrderNo " & vbCrLf & _
            "LEFT JOIN MS_PartMapping MPM On MPM.PartNo = POD.PartNo AND MPM.AffiliateID = POD.AffiliateID AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
            "INNER JOIN MS_Parts MP ON MP.PartNo = POD.partNo " & vbCrLf & _
            "INNER JOIN ms_unitcls MU ON MU.UnitCls = MP.Unitcls " & vbCrLf & _
            "WHERE POM.PONo = '" & Trim(PONO) & "' " & vbCrLf & _
            "AND POD.OrderNo1 = '" & Trim(ls_orderNo) & "' " & vbCrLf & _
            "AND POD.AffiliateID = '" & Trim(Aff) & "' " & vbCrLf & _
            "AND POD.SupplierID = '" & Trim(Supp) & "' " & vbCrLf & _
            "AND ISNULL(DOM.SuratJalanno, '') = '" & Trim(pSJ) & "' "

        'End If

        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Private Function sendEmailPOCancel(ByVal pFileName As String, ByVal pAFF As String, ByVal pFWD As String, ByVal pPONo As String, ByVal pOrderNo1 As String, ByVal pTo As String) As Boolean
        Try
            Dim TempFilePath As String
            Dim TempFileName As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFilePath = Trim(txtSaveAsDOM.Text)
            TempFileName = "\" & pFileName

            If pTo = "Supplier" Then
                Dim dsEmail As New DataSet
                dsEmail = EmailToEmailCCPOMonthly(pAffCode, "PASI", pSupplier)
                'To Supplier, CC Supplier
                For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                    If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                        fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                    End If
                    If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                        If receiptEmail = "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                        Else
                            receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                        End If
                    End If
                    If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                        End If
                    End If
                Next
            ElseIf pTo = "Forwarder" Then
                Dim dsEmail As New DataSet
                dsEmail = EmailSendForwarder_Export(pFWD, pAFF)
                'To Supplier, CC Supplier
                For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                    If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                        If fromEmail = "" Then
                            fromEmail = dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                        Else
                            fromEmail = fromEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                        End If
                    End If
                    If dsEmail.Tables(0).Rows(iRow)("flag") <> "PASI" Then
                        If receiptEmail = "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                        Else
                            receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                        End If
                    End If
                    If dsEmail.Tables(0).Rows(iRow)("flag") <> "PASI" Then
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                        End If
                    End If
                Next
            Else
                Dim dsEmail As New DataSet
                dsEmail = EmailToEmailCCPOMonthly(pAffCode, "PASI", pSupplier)
                For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                    If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                        fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                    End If
                    If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                        If receiptEmail = "" Then
                            receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                        Else
                            receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                        End If
                    End If
                    If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                        If receiptCCEmail = "" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                        Else
                            receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                        End If
                    End If
                Next
            End If

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Cancelation List STOPPED, because Recipient's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPOCancel = True
                Exit Function
            End If
            If fromEmail = "" Then
                rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Cancelation List STOPPED, because Mailer's e-mail address is not found" & vbCrLf & _
                                rtbProcess.Text
                sendEmailPOCancel = True
                Exit Function
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "CA-" & pSupplier & "-" & pOrderNo1.Trim & " PO Cancelation "

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            GetSettingEmail_Export("PO")

            ls_Body = clsNotification.GetNotification("27", "", pOrderNo1.Trim)

            mailMessage.Body = ls_Body
            Dim filename As String = TempFilePath & TempFileName
            mailMessage.Attachments.Add(New Attachment(filename))
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            
            smtp.Host = smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
            sendEmailPOCancel = True

            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Cancelation List, PO No: " & pOrderNo1.Trim & " SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
            Exit Function
        Catch ex As Exception
            sendEmailPOCancel = False
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send PO Cancelation List, PO No: " & pOrderNo1.Trim & " STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
        End Try
    End Function

    Private Sub UpdateStatusPOCancelation(ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pOrderNo As String, ByVal pSJ As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE PO_Master_ExportCancel SET ExcelCls = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND PONo = '" & pPoNo & "'" & vbCrLf & _
                         " AND OrderNo1 = '" & pOrderNo & "'"

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub

    Shared Sub NAR(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub

    Private Sub up_InvoiceSupplierExport()
        Dim SJNo = "", AffID = "", SuppID = "", FwdID = "", PoNo = "", OrderNo As String = ""
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim replyEmailExcel As String = ""
        Dim fileNameTemplate As String = Trim(txtAttachmentDOM.Text) & "\Template Invoice Supplier Export.xlsm"
        Dim fileName As String = ""

        Dim receiptCCEmailGR As String = ""
        Dim receiptEmailGR As String = ""
        Dim replyEmailExcelGR As String = ""
        Dim fileNameTemplateGR As String = Trim(txtAttachmentDOM.Text) & "\Template GoodReceivingExport.xlsx"
        Dim fileNameGR As String = ""
        Dim ls_sql As String = ""

        Dim ds As New DataSet
        Dim dsEmail As New DataSet
        Dim dsAffiliate As New DataSet
        Dim dsSupplier As New DataSet

        Dim dsGR As New DataSet

        ls_sql = "select TRIM(SuratJalanNo) SuratJalan, TRIM(AffiliateID) AffID, TRIM(SupplierID) SuppID, TRIM(ForwarderID) FwdID, TRIM(PONo) PoNo, TRIM(OrderNo) OrderNo from ReceiveForwarder_Master where ExcelCls = '1' "
        ds = cls.uf_GetDataSet(ls_sql)

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim fi As New FileInfo(fileNameTemplate)

        If Not fi.Exists Then
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Invoice Export and Good Receiving to Supplier STOPPED, because File Excel Invoice isn't Found " & vbCrLf &
                            rtbProcess.Text
            Exit Sub
        End If

        fi = New FileInfo(fileNameTemplateGR)

        If Not fi.Exists Then
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Invoice Export and Good Receiving to Supplier STOPPED, because File Excel Good Receivng isn't Found " & vbCrLf &
                            rtbProcess.Text
            Exit Sub
        End If

        fileName = txtSaveAsDOM.Text & "\Template Invoice Export " & Format(Now, "yyyyMMddHHmmss") & ".xlsm"
        System.IO.File.Copy(fileNameTemplate, fileName)

        fileNameGR = txtSaveAsDOM.Text & "\Template Good Receiving Export " & Format(Now, "yyyyMMddHHmmss") & ".xlsx"
        System.IO.File.Copy(fileNameTemplateGR, fileNameGR)

        SJNo = ds.Tables(0).Rows(0)("SuratJalan") : FwdID = ds.Tables(0).Rows(0)("FwdID")
        AffID = ds.Tables(0).Rows(0)("AffID") : SuppID = ds.Tables(0).Rows(0)("SuppID")
        PoNo = ds.Tables(0).Rows(0)("PoNo") : OrderNo = ds.Tables(0).Rows(0)("OrderNo")

        dsEmail = EmailToEmailCCInvoice_Export(SuppID)
        dsAffiliate = Affiliate(Trim(AffID))
        dsSupplier = Supplier(Trim(SuppID))

        For i = 0 To dsEmail.Tables(0).Rows.Count - 1
            If dsEmail.Tables(0).Rows(i)("Flag") = "PASI" Then
                replyEmailExcel = dsEmail.Tables(0).Rows(i)("InvoiceTO")
                replyEmailExcelGR = dsEmail.Tables(0).Rows(i)("InvoiceTO")
            Else
                receiptEmail = dsEmail.Tables(0).Rows(i)("InvoiceTO")
                receiptEmailGR = dsEmail.Tables(0).Rows(i)("InvoiceTO")

                receiptCCEmail = dsEmail.Tables(0).Rows(i)("InvoiceCC")
                receiptCCEmailGR = dsEmail.Tables(0).Rows(i)("GRCC")
            End If
        Next

        If dsAffiliate.Tables(0).Rows.Count = 0 Then ' STOP Karena ga ada di Master Affiliate
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Invoice Export and Good Receiving to Supplier STOPPED, because Data for Affiliate : '" & AffID & "' isn't Found " & vbCrLf & rtbProcess.Text
            Exit Sub
        End If

        If dsSupplier.Tables(0).Rows.Count = 0 Then ' STOP Karena ga ada di Master Supplier
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Invoice Export and Good Receiving to Supplier STOPPED, because Data for Supplier : '" & SuppID & "' isn't Found " & vbCrLf & rtbProcess.Text
            Exit Sub
        End If

        ls_sql = "Exec GetDataInvoiceSupplier_Exp '" & SJNo & "', '" & AffID & "', '" & SuppID & "', '" & PoNo & "', '" & OrderNo & "' "
        ds = cls.uf_GetDataSet(ls_sql)

        If ds.Tables(0).Rows.Count = 0 Then ' STOP Karena ga ada Data nya
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Invoice Export and Good Receiving to Supplier STOPPED, because Data for SJ : '" & SJNo & "' and Supp '" & SuppID & "' and Order '" & OrderNo & "' isn't Found " & vbCrLf & rtbProcess.Text
            Exit Sub
        End If

        ls_sql = "Exec GetDataGR_Exp '" & SJNo & "', '" & AffID & "', '" & SuppID & "', '" & FwdID & "', '" & PoNo & "', '" & OrderNo & "' "
        dsGR = cls.uf_GetDataSet(ls_sql)

        up_GoodReceivingExport(dsGR, SJNo, AffID, SuppID, FwdID, PoNo, OrderNo, fileNameGR)

        If fileNameGR = "Error" Then
            Exit Sub
        End If

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application

        Try
            ExcelBook = xlApp.Workbooks.Open(fileName)
            ExcelSheet = CType(ExcelBook.Worksheets(1), Excel.Worksheet)
            'xlApp.Visible = True

            ExcelSheet.Range("H2").Value = replyEmailExcel
            ExcelSheet.Range("H3").Value = AffID
            ExcelSheet.Range("H4").Value = FwdID
            ExcelSheet.Range("H5").Value = SuppID

            'Invoice Date
            ExcelSheet.Range("W11").Value = Format(Now, "dd-MMM-yy")

            'Customer
            ExcelSheet.Range("I13").Value = dsAffiliate.Tables(0).Rows(0)("AffiliateName").ToString.Trim
            ExcelSheet.Range("I14").Value = dsAffiliate.Tables(0).Rows(0)("Address").ToString.Trim

            'Supplier
            ExcelSheet.Range("AM13").Value = dsSupplier.Tables(0).Rows(0)("SupplierName").ToString.Trim
            ExcelSheet.Range("AM14").Value = dsSupplier.Tables(0).Rows(0)("Address").ToString.Trim

            Dim k = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                ExcelSheet.Range("B" & k + 36 & ": AZ" & k + 36).RowHeight = 13

                ExcelSheet.Range("B" & k + 36 & ": C" & k + 36).Merge() ' No
                ExcelSheet.Range("B" & k + 36 & ": C" & k + 36).Value = k + 1
                ExcelSheet.Range("B" & k + 36 & ": C" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("B" & k + 36 & ": C" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("D" & k + 36 & ": I" & k + 36).Merge() ' SJ No
                ExcelSheet.Range("D" & k + 36 & ": I" & k + 36).Value = ds.Tables(0).Rows(i)("SJNo")
                ExcelSheet.Range("D" & k + 36 & ": I" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("D" & k + 36 & ": I" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("J" & k + 36 & ": L" & k + 36).Merge() ' PO No
                ExcelSheet.Range("J" & k + 36 & ": L" & k + 36).Value = ds.Tables(0).Rows(i)("PONo")
                ExcelSheet.Range("J" & k + 36 & ": L" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("J" & k + 36 & ": L" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("M" & k + 36 & ": P" & k + 36).Merge() ' Order No
                ExcelSheet.Range("M" & k + 36 & ": P" & k + 36).Value = ds.Tables(0).Rows(i)("OrderNo")
                ExcelSheet.Range("M" & k + 36 & ": P" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("M" & k + 36 & ": P" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("Q" & k + 36 & ": U" & k + 36).Merge() ' Part No
                ExcelSheet.Range("Q" & k + 36 & ": U" & k + 36).Value = ds.Tables(0).Rows(i)("PartNo")
                ExcelSheet.Range("Q" & k + 36 & ": U" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("Q" & k + 36 & ": U" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("V" & k + 36 & ": AD" & k + 36).Merge() ' Part Name
                ExcelSheet.Range("V" & k + 36 & ": AD" & k + 36).Value = ds.Tables(0).Rows(i)("PartDesc")
                ExcelSheet.Range("V" & k + 36 & ": AD" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("V" & k + 36 & ": AD" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AE" & k + 36 & ": AF" & k + 36).Merge() ' UOM
                ExcelSheet.Range("AE" & k + 36 & ": AF" & k + 36).Value = ds.Tables(0).Rows(i)("UOM")
                ExcelSheet.Range("AE" & k + 36 & ": AF" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("AE" & k + 36 & ": AF" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AG" & k + 36 & ": AH" & k + 36).Merge() ' Qty/Box
                ExcelSheet.Range("AG" & k + 36 & ": AH" & k + 36).Value = CDbl(ds.Tables(0).Rows(i)("MOQ"))
                ExcelSheet.Range("AG" & k + 36 & ": AH" & k + 36).NumberFormat = "#,##0"
                ExcelSheet.Range("AG" & k + 36 & ": AH" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                ExcelSheet.Range("AG" & k + 36 & ": AH" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AI" & k + 36 & ": AL" & k + 36).Merge() ' Supp Deliv Qty
                ExcelSheet.Range("AI" & k + 36 & ": AL" & k + 36).Value = CDbl(ds.Tables(0).Rows(i)("SuppQty"))
                ExcelSheet.Range("AI" & k + 36 & ": AL" & k + 36).NumberFormat = "#,##0"
                ExcelSheet.Range("AI" & k + 36 & ": AL" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                ExcelSheet.Range("AI" & k + 36 & ": AL" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).Merge() ' Invoice Qty
                ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).Value = CDbl(ds.Tables(0).Rows(i)("InvQty"))
                ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).NumberFormat = "#,##0"
                ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).Font.Bold = False
                ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AP" & k + 36 & ": AQ" & k + 36).Merge() ' Curr
                ExcelSheet.Range("AP" & k + 36 & ": AQ" & k + 36).Value = ds.Tables(0).Rows(i)("Currency")
                ExcelSheet.Range("AP" & k + 36 & ": AQ" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("AP" & k + 36 & ": AQ" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AR" & k + 36 & ": AZ" & k + 36).MergeCells = False
                ExcelSheet.Range("AR" & k + 36 & ": AU" & k + 36).Merge() ' Price
                ExcelSheet.Range("AR" & k + 36 & ": AU" & k + 36).Value = CDbl(ds.Tables(0).Rows(i)("Price"))
                ExcelSheet.Range("AR" & k + 36 & ": AU" & k + 36).NumberFormat = "#,##0"
                ExcelSheet.Range("AR" & k + 36 & ": AU" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                ExcelSheet.Range("AR" & k + 36 & ": AU" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AV" & k + 36 & ": AZ" & k + 36).Merge() ' Amount
                ExcelSheet.Range("AV" & k + 36 & ": AZ" & k + 36).Value = CDbl(ds.Tables(0).Rows(i)("Amount"))
                ExcelSheet.Range("AV" & k + 36 & ": AZ" & k + 36).NumberFormat = "#,##0"
                ExcelSheet.Range("AV" & k + 36 & ": AZ" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                ExcelSheet.Range("AV" & k + 36 & ": AZ" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                clsGeneral.DrawAllBorders(ExcelSheet.Range("B" & k + 36 & ": AZ" & k + 36))
                ExcelSheet.Range("B" & k + 36 & ": AZ" & k + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("B" & k + 36 & ": AZ" & k + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                k += 1
            Next

            ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).Merge() ' words TOTAL
            ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).Value = "TOTAL"
            ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).Font.Bold = True
            ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            ExcelSheet.Range("AM" & k + 36 & ": AO" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

            ExcelSheet.Range("AP" & k + 36 & ": AQ" & k + 36).Merge() ' Tot Curr
            ExcelSheet.Range("AP" & k + 36 & ": AQ" & k + 36).Value = ds.Tables(0).Rows(ds.Tables(0).Rows.Count - 1)("Currency")
            ExcelSheet.Range("AP" & k + 36 & ": AQ" & k + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            ExcelSheet.Range("AP" & k + 36 & ": AQ" & k + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

            ExcelSheet.Range("AR" & k + 36 & ": AZ" & k + 36).Merge() ' Total Amount
            ExcelSheet.Range("AR" & k + 36 & ": AZ" & k + 36).Value = CDbl(ds.Tables(1).Rows(0)("TotAmount"))
            ExcelSheet.Range("AR" & k + 36 & ": AZ" & k + 36).NumberFormat = "#,##0"
            ExcelSheet.Range("AR" & k + 36 & ": AZ" & k + 36).Font.Bold = True

            clsGeneral.DrawAllBorders(ExcelSheet.Range("AM" & k + 36 & ": AZ" & k + 36))
            ExcelSheet.Range("AM" & k + 36 & ": AZ" & k + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            ExcelSheet.Range("AM" & k + 36 & ": AZ" & k + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            ExcelSheet.Range("B37").Interior.Color = Color.White
            ExcelSheet.Range("B37").Font.Color = Color.Black
            ExcelSheet.Range("B" & k + 36).Value = "E"
            ExcelSheet.Range("B" & k + 36).Interior.Color = Color.Black
            ExcelSheet.Range("B" & k + 36).Font.Color = Color.White

            xlApp.DisplayAlerts = False

            Dim pFilename As String = ""
            pFilename = txtSaveAsDOM.Text & "\Invoice Export " & SuppID & "-" & OrderNo & ".xlsm"
            ExcelBook.SaveAs(pFilename)
            ExcelBook.Close()
            xlApp.Workbooks.Close()
            xlApp.Quit()

            My.Computer.FileSystem.DeleteFile(fileName)

            If sendEmailInvoice_EXPORT(pFilename, replyEmailExcel, receiptEmail, receiptCCEmail, SJNo, SuppID) = False Then Exit Sub
            If sendEmailGR_EXPORT(fileNameGR, replyEmailExcelGR, receiptEmailGR, receiptCCEmailGR, SJNo, SuppID) = False Then Exit Sub

            Call UpdateStatusReceiveFWDExport(SJNo, AffID, SuppID, PoNo, OrderNo)
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Invoice Export and Good Receiving to Supplier STOPPED, because " & Err.Description & " " & vbCrLf &
                                    rtbProcess.Text
            xlApp.Workbooks.Close()
            xlApp.Quit()
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Sub up_GoodReceivingExport(ds As DataSet, SjNo As String, AffID As String, SuppID As String, FwdID As String, PoNo As String, OrderNo As String, ByRef fileNameGR As String)
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application
        Dim tempFileName = fileNameGR

        Try
            ExcelBook = xlApp.Workbooks.Open(fileNameGR)
            ExcelSheet = CType(ExcelBook.Worksheets(1), Excel.Worksheet)
            'xlApp.Visible = True

            ExcelSheet.Range("M8").Value = ds.Tables(0).Rows(0)("SJNo").ToString()
            ExcelSheet.Range("M9").Value = ds.Tables(0).Rows(0)("DeliveryDate").ToString()
            ExcelSheet.Range("M10").Value = ds.Tables(0).Rows(0)("ReceiveDate").ToString()
            ExcelSheet.Range("M11").Value = ds.Tables(0).Rows(0)("AffiliateID").ToString()
            ExcelSheet.Range("M12").Value = ds.Tables(0).Rows(0)("DeliveryTo").ToString()

            ExcelSheet.Range("AJ8").Value = ds.Tables(0).Rows(0)("ReceiveBy").ToString()
            ExcelSheet.Range("AJ9").Value = ds.Tables(0).Rows(0)("JenisArmada").ToString()
            ExcelSheet.Range("AJ10").Value = ds.Tables(0).Rows(0)("NoPol").ToString()
            ExcelSheet.Range("AJ11").Value = ds.Tables(0).Rows(0)("DeliveryName").ToString()
            ExcelSheet.Range("AJ12").Value = ds.Tables(0).Rows(0)("TotalBox").ToString()

            Dim k = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                ExcelSheet.Range("B" & k + 18 & ": AJ" & k + 18).RowHeight = 13

                ExcelSheet.Range("B" & k + 18 & ": C" & k + 18).Merge() ' No
                ExcelSheet.Range("B" & k + 18 & ": C" & k + 18).Value = k + 1
                ExcelSheet.Range("B" & k + 18 & ": C" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("B" & k + 18 & ": C" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("D" & k + 18 & ": I" & k + 18).Merge() ' PO No
                ExcelSheet.Range("D" & k + 18 & ": I" & k + 18).Value = ds.Tables(0).Rows(i)("PONo")
                ExcelSheet.Range("D" & k + 18 & ": I" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("D" & k + 18 & ": I" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("J" & k + 18 & ": N" & k + 18).Merge() ' Part No
                ExcelSheet.Range("J" & k + 18 & ": N" & k + 18).Value = ds.Tables(0).Rows(i)("PartNo")
                ExcelSheet.Range("J" & k + 18 & ": N" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("J" & k + 18 & ": N" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("O" & k + 18 & ": W" & k + 18).Merge() ' Part Name
                ExcelSheet.Range("O" & k + 18 & ": W" & k + 18).Value = ds.Tables(0).Rows(i)("PartDesc")
                ExcelSheet.Range("O" & k + 18 & ": W" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("O" & k + 18 & ": W" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("X" & k + 18 & ": Y" & k + 18).Merge() ' UOM
                ExcelSheet.Range("X" & k + 18 & ": Y" & k + 18).Value = ds.Tables(0).Rows(i)("UOM")
                ExcelSheet.Range("X" & k + 18 & ": Y" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                ExcelSheet.Range("X" & k + 18 & ": Y" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("Z" & k + 18 & ": AA" & k + 18).Merge() ' MOQ
                ExcelSheet.Range("Z" & k + 18 & ": AA" & k + 18).Value = ds.Tables(0).Rows(i)("MOQ")
                ExcelSheet.Range("Z" & k + 18 & ": AA" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("Z" & k + 18 & ": AA" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AB" & k + 18 & ": AE" & k + 18).Merge() ' Supplier Delivery Qty
                ExcelSheet.Range("AB" & k + 18 & ": AE" & k + 18).Value = ds.Tables(0).Rows(i)("SuppQty")
                ExcelSheet.Range("AB" & k + 18 & ": AE" & k + 18).NumberFormat = "#,##0"
                ExcelSheet.Range("AB" & k + 18 & ": AE" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("AB" & k + 18 & ": AE" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AF" & k + 18 & ": AI" & k + 18).Merge() ' Receiving Qty
                ExcelSheet.Range("AF" & k + 18 & ": AI" & k + 18).Value = CDbl(ds.Tables(0).Rows(i)("RecQty"))
                ExcelSheet.Range("AF" & k + 18 & ": AI" & k + 18).NumberFormat = "#,##0"
                ExcelSheet.Range("AF" & k + 18 & ": AI" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("AF" & k + 18 & ": AI" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AJ" & k + 18 & ": AM" & k + 18).Merge() ' Receiving Qty Box
                ExcelSheet.Range("AJ" & k + 18 & ": AM" & k + 18).Value = CDbl(ds.Tables(0).Rows(i)("RecBox"))
                ExcelSheet.Range("AJ" & k + 18 & ": AM" & k + 18).NumberFormat = "#,##0"
                ExcelSheet.Range("AJ" & k + 18 & ": AM" & k + 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("AJ" & k + 18 & ": AM" & k + 18).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                clsGeneral.DrawAllBorders(ExcelSheet.Range("B" & k + 18 & ": AM" & k + 18))
                ExcelSheet.Range("B" & k + 18 & ": AM" & k + 18).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("B" & k + 18 & ": AM" & k + 18).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                k += 1
            Next

            xlApp.DisplayAlerts = False

            Dim pFilename As String = ""
            pFilename = txtSaveAsDOM.Text & "\Good Receiving Export " & SuppID & "-" & OrderNo & ".xlsx"
            ExcelBook.SaveAs(pFilename)
            ExcelBook.Close()
            xlApp.Workbooks.Close()
            xlApp.Quit()

            My.Computer.FileSystem.DeleteFile(tempFileName)
            fileNameGR = pFilename
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Invoice Export and Good Receiving to Supplier STOPPED, because " & Err.Description & " " & vbCrLf &
                                    rtbProcess.Text
            xlApp.Workbooks.Close()
            xlApp.Quit()

            fileNameGR = "Error"
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Sub up_MovingGoodExport()
        Dim SJNo = "", SJNoOld = "", AffID = "", SuppID = "", FwdID = "", FwdID_Old = "", PoNo = "", OrderNo = ""
        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim replyEmailExcel As String = ""

        Dim receiptCCEmail_Split As String = ""
        Dim receiptEmail_Split As String = ""

        Dim fileNameTemplate As String = Trim(txtAttachmentDOM.Text) & "\Template Moving Delivery Split GR.xlsx"
        Dim fileName As String = ""
        Dim ls_sql As String = ""

        Dim ds As New DataSet
        Dim dsEmail As New DataSet
        Dim dsAffiliate As New DataSet
        Dim dsSupplier As New DataSet

        ls_sql = "Select Top 1 TRIM(SuratJalanNo) SuratJalan, TRIM(AffiliateID) AffID, TRIM(SupplierID) SuppID, TRIM(PONo) PoNo, TRIM(OrderNo) OrderNo, SplitDelivery SJDelivery, SJDeliveryOri From DOSupplier_Master_Export where SplitCls = '1' order by UpdateDate "
        ds = cls.uf_GetDataSet(ls_sql)

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim fi As New FileInfo(fileNameTemplate)

        If Not fi.Exists Then
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder STOPPED, because File Excel isn't Found " & vbCrLf & _
                            rtbProcess.Text
            Exit Sub
        End If

        fileName = txtSaveAsDOM.Text & "\Template Moving Delivery Split GR " & Format(Now, "yyyyMMddHHmmss") & ".xlsx"
        System.IO.File.Copy(fileNameTemplate, fileName)

        SJNo = ds.Tables(0).Rows(0)("SuratJalan") : SJNoOld = ds.Tables(0).Rows(0)("SJDelivery")
        AffID = ds.Tables(0).Rows(0)("AffID") : SuppID = ds.Tables(0).Rows(0)("SuppID")
        PoNo = ds.Tables(0).Rows(0)("PoNo") : OrderNo = ds.Tables(0).Rows(0)("OrderNo")

        'Get Email Forwarder yang Baru
        ls_sql = "select Top 1 ForwarderID from ReceiveForwarder_Master where AffiliateID = '" & AffID & "' and SupplierID = '" & SuppID & "' and PONo = '" & PoNo & "' and OrderNo = '" & OrderNo & "'"
        ds = cls.uf_GetDataSet(ls_sql)

        If ds.Tables(0).Rows.Count = 0 Then
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder STOPPED, because Data Email Forwarder for SJ : '" & SJNo & "' and Supp '" & SuppID & "' and Order '" & OrderNo & "' isn't Found " & vbCrLf & rtbProcess.Text
            Exit Sub
        End If

        FwdID = ds.Tables(0).Rows(0)("ForwarderID").ToString ' Pasti ke 0

        'Get Data Detail
        ls_sql = "Exec sp_DeliverySplitBatch_Select_Detail_Moving '" & AffID & "', '" & SuppID & "', '" & SJNo & "', '" & SJNoOld & "', '" & PoNo & "', '" & OrderNo & "' "
        ds = cls.uf_GetDataSet(ls_sql)

        If ds.Tables(0).Rows.Count = 0 Or ds.Tables(1).Rows.Count = 0 Then ' STOP Karena ga ada Data nya
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder STOPPED, because Data for SJ : '" & SJNo & "' and Supp '" & SuppID & "' and Order '" & OrderNo & "' isn't Found " & vbCrLf & rtbProcess.Text
            Exit Sub
        End If

        FwdID_Old = ds.Tables(0).Rows(0)("OldForwarder").ToString ' Pasti ke 0

        dsEmail = EmailToEmailCCMovingGR_Export(FwdID, FwdID_Old)
        dsAffiliate = Affiliate(Trim(AffID))
        dsSupplier = Supplier(Trim(SuppID))

        For i = 0 To dsEmail.Tables(0).Rows.Count - 1
            If dsEmail.Tables(0).Rows(i)("Flag") = "PASI" Then
                replyEmailExcel = dsEmail.Tables(0).Rows(i)("ForwarderTO")
            ElseIf dsEmail.Tables(0).Rows(i)("Flag") = "FWD" Then
                receiptEmail_Split = dsEmail.Tables(0).Rows(i)("ForwarderTO")
                receiptCCEmail_Split = dsEmail.Tables(0).Rows(i)("ForwarderCC")
            ElseIf dsEmail.Tables(0).Rows(i)("Flag") = "FWDOld" Then
                receiptEmail = dsEmail.Tables(0).Rows(i)("ForwarderTO")
                receiptCCEmail = dsEmail.Tables(0).Rows(i)("ForwarderCC")
            End If
        Next

        If dsAffiliate.Tables(0).Rows.Count = 0 Then ' STOP Karena ga ada di Master Affiliate
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder STOPPED, because Data for Affiliate : '" & AffID & "' isn't Found " & vbCrLf & rtbProcess.Text
            Exit Sub
        End If

        If dsSupplier.Tables(0).Rows.Count = 0 Then ' STOP Karena ga ada di Master Supplier
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder STOPPED, because Data for Supplier : '" & SuppID & "' isn't Found " & vbCrLf & rtbProcess.Text
            Exit Sub
        End If

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application

        Try
            ExcelBook = xlApp.Workbooks.Open(fileName)
            ExcelSheet = CType(ExcelBook.Worksheets(1), Excel.Worksheet)
            'xlApp.Visible = True

            Dim k = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                ExcelSheet.Range("A" & k + 5 & ": P" & k + 5).RowHeight = 20

                ExcelSheet.Range("A" & k + 5).Value = ds.Tables(0).Rows(i)("No").ToString 'No
                ExcelSheet.Range("B" & k + 5).Value = ds.Tables(0).Rows(i)("Affiliate").ToString 'Affiliate ID
                ExcelSheet.Range("C" & k + 5).Value = ds.Tables(0).Rows(i)("Supplier").ToString 'Supplier ID
                ExcelSheet.Range("D" & k + 5).Value = ds.Tables(0).Rows(i)("Forwarder").ToString 'Forwarder ID
                ExcelSheet.Range("E" & k + 5).Value = ds.Tables(0).Rows(i)("PoNo").ToString 'Po No
                ExcelSheet.Range("F" & k + 5).Value = ds.Tables(0).Rows(i)("OrderNo").ToString 'Order No
                ExcelSheet.Range("G" & k + 5).Value = ds.Tables(0).Rows(i)("SuratJalan").ToString 'DN No
                ExcelSheet.Range("H" & k + 5).Value = ds.Tables(0).Rows(i)("PartNo").ToString 'Part No
                ExcelSheet.Range("I" & k + 5).Value = ds.Tables(0).Rows(i)("BoxNo").ToString 'Box No
                ExcelSheet.Range("J" & k + 5).Value = ds.Tables(0).Rows(i)("Consignee").ToString 'Consignee Code
                ExcelSheet.Range("K" & k + 5).Value = ds.Tables(0).Rows(i)("ReceiveDate").ToString 'Receive Date
                ExcelSheet.Range("L" & k + 5).Value = ds.Tables(0).Rows(i)("UOM").ToString 'UOM
                ExcelSheet.Range("M" & k + 5).Value = CDbl(ds.Tables(0).Rows(i)("MOQ").ToString) 'Qty/Box
                ExcelSheet.Range("M" & k + 5).NumberFormat = "#,##0"

                ExcelSheet.Range("N" & k + 5).Value = ds.Tables(0).Rows(i)("StatusGR").ToString 'Status GR
                ExcelSheet.Range("O" & k + 5).Value = ds.Tables(0).Rows(i)("PlanUpdate").ToString 'Plan Update Date
                ExcelSheet.Range("P" & k + 5).Value = ds.Tables(0).Rows(i)("Remarks").ToString 'Remarks

                ExcelSheet.Range("A" & k + 5 & ": P" & k + 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("A" & k + 5 & ": P" & k + 5).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter


                clsGeneral.DrawAllBorders(ExcelSheet.Range("A" & k + 5 & ": P" & k + 5))
                ExcelSheet.Range("A" & k + 5 & ": P" & k + 5).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("A" & k + 5 & ": P" & k + 5).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

                If ds.Tables(0).Rows(i)("Status").ToString() = "1" Then
                    ExcelSheet.Range("A" & k + 5 & ": P" & k + 5).Interior.Color = Color.Yellow
                End If

                k += 1
            Next

            xlApp.DisplayAlerts = False

            Dim pFilename As String = ""
            pFilename = txtSaveAsDOM.Text & "\Moving Delivery Split GR " & PoNo & " (" & OrderNo & ").xlsx"
            ExcelBook.SaveAs(pFilename)
            ExcelBook.Close()
            xlApp.Workbooks.Close()
            xlApp.Quit()

            My.Computer.FileSystem.DeleteFile(fileName)

            '<!-- Process KIRIM EXCEL KE FORWARDER BARU YANG DI SPLIT AJAH --!>
            fileNameTemplate = Trim(txtAttachmentDOM.Text) & "\Template Moving Delivery Split GR Forwarder Baru.xlsx"
            fileName = txtSaveAsDOM.Text & "\Template Moving Delivery Split GR Forwarder Baru " & Format(Now, "yyyyMMddHHmmss") & ".xlsx"
            System.IO.File.Copy(fileNameTemplate, fileName)

            ExcelBook = xlApp.Workbooks.Open(fileName)
            ExcelSheet = CType(ExcelBook.Worksheets(1), Excel.Worksheet)

            k = 0
            For i = 0 To ds.Tables(1).Rows.Count - 1
                ExcelSheet.Range("A" & k + 5 & ": L" & k + 5).RowHeight = 20

                ExcelSheet.Range("A" & k + 5).Value = ds.Tables(1).Rows(i)("No").ToString 'No
                ExcelSheet.Range("B" & k + 5).Value = ds.Tables(1).Rows(i)("Affiliate").ToString 'Affiliate ID
                ExcelSheet.Range("C" & k + 5).Value = ds.Tables(1).Rows(i)("Supplier").ToString 'Supplier ID
                ExcelSheet.Range("D" & k + 5).Value = ds.Tables(1).Rows(i)("Forwarder").ToString 'Forwarder ID
                ExcelSheet.Range("E" & k + 5).Value = ds.Tables(1).Rows(i)("PoNo").ToString 'Po No
                ExcelSheet.Range("F" & k + 5).Value = ds.Tables(1).Rows(i)("OrderNo").ToString 'Order No
                ExcelSheet.Range("G" & k + 5).Value = ds.Tables(1).Rows(i)("SuratJalan").ToString 'DN No
                ExcelSheet.Range("H" & k + 5).Value = ds.Tables(1).Rows(i)("PartNo").ToString 'Part No
                ExcelSheet.Range("I" & k + 5).Value = ds.Tables(1).Rows(i)("BoxNo").ToString 'Box No
                ExcelSheet.Range("J" & k + 5).Value = ds.Tables(1).Rows(i)("Consignee").ToString 'Consignee Code
                ExcelSheet.Range("K" & k + 5).Value = ds.Tables(1).Rows(i)("UOM").ToString 'UOM
                ExcelSheet.Range("L" & k + 5).Value = CDbl(ds.Tables(1).Rows(i)("MOQ").ToString) 'Qty/Box
                ExcelSheet.Range("L" & k + 5).NumberFormat = "#,##0"

                ExcelSheet.Range("A" & k + 5 & ": L" & k + 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("A" & k + 5 & ": L" & k + 5).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                clsGeneral.DrawAllBorders(ExcelSheet.Range("A" & k + 5 & ": L" & k + 5))
                ExcelSheet.Range("A" & k + 5 & ": L" & k + 5).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("A" & k + 5 & ": L" & k + 5).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

                k += 1
            Next

            xlApp.DisplayAlerts = False

            Dim pFilename2 As String = ""
            pFilename2 = txtSaveAsDOM.Text & "\Delivery Split GR " & PoNo & " (" & OrderNo & ").xlsx"
            ExcelBook.SaveAs(pFilename2)
            ExcelBook.Close()
            xlApp.Workbooks.Close()
            xlApp.Quit()

            My.Computer.FileSystem.DeleteFile(fileName)

            'Kirim Excel ke Forwarder Baru
            If sendEmailMovingGRFWDSplit_EXPORT(pFilename2, replyEmailExcel, receiptEmail_Split, receiptCCEmail_Split, SJNo) = False Then Exit Sub

            'Kirim Excel ke Forwarder Lama
            If sendEmailMovingGR_EXPORT(pFilename, replyEmailExcel, receiptEmail, receiptCCEmail, SJNoOld, SJNo) = False Then Exit Sub

            Call UpdateStatusDOSupplierMovingGRExport(SJNo, AffID, SuppID, PoNo, OrderNo)
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder STOPPED, because " & Err.Description & " " & vbCrLf & _
                                    rtbProcess.Text
            xlApp.Workbooks.Close()
            xlApp.Quit()
        Finally
            If Not xlApp Is Nothing Then
                NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                NAR(ExcelBook)
                NAR(ExcelSheet)
                xlApp.Quit()
                NAR(xlApp)
                GC.Collect()
            End If
        End Try
    End Sub

    Private Function EmailToEmailCCInvoice_Export(ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = " select 'PASI' as FLAG , InvoiceCC = TRIM(InvoiceCC), InvoiceTO = TRIM(InvoiceTO), GRCC = '', GRTO = '' from MS_EmailPasi_Export " & vbCrLf &
                 "UNION ALL " & vbCrLf &
                 " select 'SUPP' as FLAG , InvoiceCC = TRIM(InvoiceCC), InvoiceTO = TRIM(InvoiceTO), GRCC = TRIM(KanbanCC), GRTO = '' from MS_EmailSupplier_Export where supplierID = '" & pSupplierID & "' "
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function EmailToEmailCCMovingGR_Export(ByVal pForwarderID As String, ByVal pForwarderIDOld As String) As DataSet
        Dim ls_SQL As String = ""
        'MdlConn.ReadConnection()
        ls_SQL = " select 'PASI' as FLAG , ForwarderCC = TRIM(GoodReceiveCC), ForwarderTO = TRIM(GoodReceiveTO) from MS_EmailPasi_Export " & vbCrLf & _
                 "UNION ALL " & vbCrLf & _
                 " select 'FWDOld' as FLAG , ForwarderCC = TRIM(ForwarderReceivingTO), ForwarderTO = TRIM(ForwarderReceivingTO) from MS_EmailForwarder where ForwarderID = '" & pForwarderIDOld & "' " & vbCrLf & _
                 "UNION ALL " & vbCrLf & _
                 " select 'FWD' as FLAG , ForwarderCC = TRIM(ForwarderReceivingTO), ForwarderTO = TRIM(ForwarderReceivingTO) from MS_EmailForwarder where ForwarderID = '" & pForwarderID & "' "
        Dim ds As New DataSet
        ds = cls.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Function sendEmailInvoice_EXPORT(ByVal pFilename As String, ByVal pEmailFrom As String, ByVal pEmailTo As String, ByVal pEmailCC As String, ByVal pSjNo As String, ByVal pSuppID As String) As Boolean 'Link Affiliate Order Entry
        Try
            pEmailFrom = Replace(pEmailFrom, " ", "")
            pEmailTo = Replace(pEmailTo, " ", "")
            pEmailCC = Replace(pEmailCC, " ", "")

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(pEmailFrom)
            mailMessage.Subject = "Invoice Export SJ No. " & pSjNo

            If pEmailTo <> "" Then
                For Each recipient In pEmailTo.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If pEmailCC <> "" Then
                For Each recipientCC In pEmailCC.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            GetSettingEmail_Export("Kanban")

            ls_Body = clsNotification.GetNotification("53", "", "", "", pSjNo, "", "")
            mailMessage.Body = ls_Body

            mailMessage.Attachments.Add(New Attachment(pFilename))

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            smtp.Host = smtpClient
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            smtp.Port = portClient
            smtp.Send(mailMessage)
            sendEmailInvoice_EXPORT = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Invoice SJ No: " & pSjNo & " to Supplier SUCCESSFULL" & vbCrLf &
                             rtbProcess.Text
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Invoice SJ No: " & pSjNo & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf &
                            rtbProcess.Text
            sendEmailInvoice_EXPORT = False
        End Try
    End Function

    Private Function sendEmailGR_EXPORT(ByVal pFilename As String, ByVal pEmailFrom As String, ByVal pEmailTo As String, ByVal pEmailCC As String, ByVal pSjNo As String, ByVal pSuppID As String) As Boolean 'Link Affiliate Order Entry
        Try
            pEmailFrom = Replace(pEmailFrom, " ", "")
            pEmailTo = Replace(pEmailTo, " ", "")
            pEmailCC = Replace(pEmailCC, " ", "")

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(pEmailFrom)
            mailMessage.Subject = "Good Receiving Export SJ No. " & pSjNo

            If pEmailTo <> "" Then
                For Each recipient In pEmailTo.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If pEmailCC <> "" Then
                For Each recipientCC In pEmailCC.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            GetSettingEmail_Export("Kanban")

            ls_Body = clsNotification.GetNotification("57", "", "", "", pSjNo, "", "")
            mailMessage.Body = ls_Body

            mailMessage.Attachments.Add(New Attachment(pFilename))

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            smtp.Host = smtpClient
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            smtp.Port = portClient
            smtp.Send(mailMessage)
            sendEmailGR_EXPORT = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send GR SJ No: " & pSjNo & " to Supplier SUCCESSFULL" & vbCrLf &
                             rtbProcess.Text
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send GR SJ No: " & pSjNo & " to Supplier STOPPED, because " & ex.Message & " " & vbCrLf &
                            rtbProcess.Text
            sendEmailGR_EXPORT = False
        End Try
    End Function

    Private Sub UpdateStatusReceiveFWDExport(ByVal pSJNo As String, ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pOrderNo As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = "administrator"

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " update ReceiveForwarder_Master set ExcelCls = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND SuratJalanNo = '" & pSJNo & "'" & vbCrLf & _
                         " AND OrderNo = '" & pOrderNo & "'" & vbCrLf & _
                         " AND PoNo = '" & pPoNo & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub

    Private Function sendEmailMovingGR_EXPORT(ByVal pFilename As String, ByVal pEmailFrom As String, ByVal pEmailTo As String, ByVal pEmailCC As String, ByVal pSjNo As String, ByVal pSjNoSplit As String) As Boolean 'Link Affiliate Order Entry
        Try
            pEmailFrom = Replace(pEmailFrom, " ", "")
            pEmailTo = Replace(pEmailTo, " ", "")
            pEmailCC = Replace(pEmailCC, " ", "")

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(pEmailFrom)
            mailMessage.Subject = "Split Delivery Export SJ No. " & pSjNo

            If pEmailTo <> "" Then
                For Each recipient In pEmailTo.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If pEmailCC <> "" Then
                For Each recipientCC In pEmailCC.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            GetSettingEmail_Export("Moving Good Receiving Export")

            ls_Body = clsNotification.GetNotification("55", "", "", "", pSjNo, "", "")
            mailMessage.Body = ls_Body

            mailMessage.Attachments.Add(New Attachment(pFilename))

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            smtp.Host = smtpClient
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            smtp.Port = portClient
            smtp.Send(mailMessage)
            sendEmailMovingGR_EXPORT = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder SJ No: " & pSjNoSplit & " to Forwarder SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder SJ No: " & pSjNoSplit & " to Forwarder STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
            sendEmailMovingGR_EXPORT = False
        End Try
    End Function

    Private Function sendEmailMovingGRFWDSplit_EXPORT(ByVal pFilename As String, ByVal pEmailFrom As String, ByVal pEmailTo As String, ByVal pEmailCC As String, ByVal pSjNo As String) As Boolean
        Try
            pEmailFrom = Replace(pEmailFrom, " ", "")
            pEmailTo = Replace(pEmailTo, " ", "")
            pEmailCC = Replace(pEmailCC, " ", "")

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(pEmailFrom)
            mailMessage.Subject = "Delivery Export SJ No. " & pSjNo

            If pEmailTo <> "" Then
                For Each recipient In pEmailTo.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If pEmailCC <> "" Then
                For Each recipientCC In pEmailCC.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            GetSettingEmail_Export("Moving Good Receiving Export")

            ls_Body = clsNotification.GetNotification("56", "", "", "", pSjNo, "", "")
            mailMessage.Body = ls_Body

            mailMessage.Attachments.Add(New Attachment(pFilename))

            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            smtp.Host = smtpClient
            smtp.UseDefaultCredentials = DefaultCredentials
            smtp.EnableSsl = SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential

            smtp.Port = portClient
            smtp.Send(mailMessage)
            sendEmailMovingGRFWDSplit_EXPORT = True
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder New SJ No: " & pSjNo & " to Forwarder SUCCESSFULL" & vbCrLf & _
                             rtbProcess.Text
        Catch ex As Exception
            rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process Send Moving Good to Forwarder New SJ No: " & pSjNo & " to Forwarder STOPPED, because " & ex.Message & " " & vbCrLf & _
                            rtbProcess.Text
            sendEmailMovingGRFWDSplit_EXPORT = False
        End Try
    End Function

    Private Sub UpdateStatusDOSupplierMovingGRExport(ByVal pSJNo As String, ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pOrderNo As String)
        Dim ls_SQL As String = ""

        Try
            'MdlConn.ReadConnection()
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " update DOSupplier_Master_Export set SplitCls = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND SuratJalanNo = '" & pSJNo & "'" & vbCrLf & _
                         " AND OrderNo = '" & pOrderNo & "'" & vbCrLf & _
                         " AND PoNo = '" & pPoNo & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try
    End Sub
#End Region
End Class
