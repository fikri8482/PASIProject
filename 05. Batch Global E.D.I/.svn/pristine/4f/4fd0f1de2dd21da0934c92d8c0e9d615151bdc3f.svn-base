Imports GlobalSetting
Imports System.Threading

Public Class frmSendMail

#Region "Declaration"
    Dim intervalpro As TimeSpan
    Dim processTime As Boolean

    Dim cls As clsGlobal
    Dim Log As clsLog
    Dim cfg As New clsConfig

    Dim UserLogin = "admin"

#End Region

#Region "Event"
    Private Sub frmSendMail_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Label12.Text = "Version " & Me.ProductVersion

        Try
            cls = New clsGlobal(cfg.ConnectionString, UserLogin)
            Log = New clsLog(cfg.ConnectionString, UserLogin)

            rtbProcess.Text = ""
            txtMsg.Text = ""

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
        ProcessSendEDI()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub timerProcess_Tick(sender As Object, e As System.EventArgs) Handles timerProcess.Tick
        If Format(Now, "yyyy-MM-dd HH:mm:ss") > txtNext.Text And processTime = False Then
            Me.Cursor = Cursors.WaitCursor
            ProcessSendEDI()
            Me.Cursor = Cursors.Default
        End If
    End Sub
#End Region

#Region "Procedure"
    Private Sub LoadSetting()
        Try
            txtAttachmentDom.Text = "D:\EDI\Domestic"
            txtArchiveDom.Text = "D:\EDI\Domestic\Archive"
            txtAttachmentExp.Text = "D:\EDI\Export"
            txtArchiveExp.Text = "D:\EDI\Export\Archive"
            txtSechedule.Text = 300
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ProcessSendEDI()
        Dim ErrMsg As String
        Dim errSummary As String

        timerProcess.Enabled = False
        processTime = True

        clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Batch Process", rtbProcess)

        ErrMsg = ""
        errSummary = ""
        SendInvoiceEDI_Dom()

        Thread.Sleep(5000)

        ErrMsg = ""
        errSummary = ""
        SendInvoiceEDI_Exp()

        clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Batch Process", rtbProcess)

        timerProcess.Enabled = True
        txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
        intervalpro = TimeSpan.FromSeconds(CDbl(txtSechedule.Text))
        Dim Last As Date = FormatDateTime(txtLast.Text)
        txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")
        processTime = False
    End Sub

    Private Sub SendInvoiceEDI_Dom()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Global E.D.I - Domestic", rtbProcess)

            Log.WriteToProcessLog(startTime, "SendDataEDI", "Start Send Global E.D.I - Domestic")

            clsPO.up_SendInvoiceEDI_Domestic(cfg, Log, cls, rtbProcess, txtAttachmentDom.Text.Trim & "\", txtArchiveDom.Text.Trim & "\", ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is no Invoice data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is no Invoice data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, "SendDataEDI", ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog("SendDataEDI", ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, "SendDataEDI", ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, "SendDataEDI", ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Global E.D.I - Domestic", rtbProcess)
            Log.WriteToProcessLog(startTime, "SendDataEDI", "End Send Global E.D.I - Domestic")
            Thread.Sleep(500)
        End Try
    End Sub

    Private Sub SendInvoiceEDI_Exp()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Send Global E.D.I - Export", rtbProcess)

            Log.WriteToProcessLog(startTime, "SendDataEDI", "Start Send Global E.D.I - Export")

            clsPO.up_SendInvoiceEDI_Export(cfg, Log, cls, rtbProcess, txtAttachmentExp.Text.Trim & "\", txtArchiveExp.Text.Trim & "\", ErrMsg, errSummary)
            Thread.Sleep(500)

            If ErrMsg = "-" Then
                ErrMsg = "There is no Invoice data to process."
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is no Invoice data to process." Then
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
                    Log.WriteToProcessLog(startTime, "SendDataEDI", ErrMsg)
                Else
                    clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
                    Log.WriteToErrorLog("SendDataEDI", ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                    Log.WriteToProcessLog(startTime, "SendDataEDI", ErrMsg, , , clsLog.ErrSeverity.ERR)
                End If
            End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, "SendDataEDI", ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Send Global E.D.I - Export", rtbProcess)
            Log.WriteToProcessLog(startTime, "SendDataEDI", "End Send Global E.D.I - Export")
            Thread.Sleep(500)
        End Try
    End Sub
#End Region
End Class
