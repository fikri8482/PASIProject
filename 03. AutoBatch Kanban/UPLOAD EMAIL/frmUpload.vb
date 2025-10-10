Imports GlobalSetting
Imports System.Threading

Public Class frmUpload

#Region "Decralation"
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

    Dim UserLogin As String = "admin"
    Dim screenName As String = ""

#End Region

#Region "Event"
    Private Sub frmUpload_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
            txtpath.Enabled = False
            txtPathBackup.Enabled = False
            txtPathBackupE.Enabled = False
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
        UploadEmail()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub frmGetMail_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        DisableCloseButton(Me)
    End Sub

    Private Sub timerProcess_Tick(sender As Object, e As System.EventArgs) Handles timerProcess.Tick
        If Format(Now, "yyyy-MM-dd HH:mm:ss") > txtNext.Text And processTime = False Then
            Me.Cursor = Cursors.WaitCursor
            UploadEmail()
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
            Dim ds As New DataSet

            ls_SQL = "SELECT * FROM dbo.MS_EmailSetting"
            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                txtpath.Text = Trim(ds.Tables(0).Rows(0)("AttachmentFolder"))
                txtPathBackup.Text = Trim(ds.Tables(0).Rows(0)("AttachmentBackupFolder"))
                txtSechedule.Text = ds.Tables(0).Rows(0)("Interval")
            Else
                txtpath.Text = ""
                txtPathBackup.Text = ""
                txtSechedule.Text = ""
            End If

            ls_SQL = "SELECT * FROM dbo.MS_EmailSetting_Export "
            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then                
                txtPathBackupE.Text = Trim(ds.Tables(0).Rows(0)("AttachmentBackupFolder"))
            Else                
                txtPathBackupE.Text = ""
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub UploadEmail()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""

        Try
            timerProcess.Enabled = False
            processTime = True

            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Batch Process", rtbProcess)

            ' ''01. Get Mail DOM
            uploadAttacment()

            ' ''02. Get Mail Export
            'GetExport()

            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Batch Process", rtbProcess)

        Catch ex As Exception
            cls.up_ShowMsg(ex.Message, txtMsg, GlobalSetting.clsGlobal.MsgTypeEnum.ErrorMsg)
            Log.WriteToErrorLog(Me.Tag, txtMsg.Text, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
        Finally
            timerProcess.Enabled = True
            txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            intervalpro = TimeSpan.FromSeconds(CDbl(txtSechedule.Text))
            Dim Last As Date = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")
            processTime = False
        End Try
    End Sub

    Private Sub uploadAttacment()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Dim jmlFile As Integer
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim startTime As DateTime = Now
        Try
            screenName = "UploadEmail"
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start upload attachment", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start upload attachment")

            Dim fi = From f In New IO.DirectoryInfo(txtpath.Text).GetFiles().Cast(Of IO.FileInfo)() _
                  Where f.Extension = ".xlsm" OrElse f.Extension = ".xlsx"
                  Order By f.Name
                  Select f

            jmlFile = CDbl(fi.Count)

            If jmlFile = "0" Then
                ErrMsg = "There is no attachment to process."
            Else
                clsFilterEmail.up_FilterEmail(cfg, Log, cls, rtbProcess, txtpath.Text.Trim, txtPathBackup.Text.Trim, txtPathBackupE.Text.Trim, screenName, ErrMsg, errSummary)
                Thread.Sleep(500)
            End If

            If ErrMsg <> "" Then
                If ErrMsg = "There is no attachment to process." Then
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
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End upload attachment", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End upload attachment")
            Thread.Sleep(500)
        End Try
    End Sub
#End Region

End Class
