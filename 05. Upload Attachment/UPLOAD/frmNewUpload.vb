Imports System.IO
Imports System.Data
Imports System.Data.SqlClient '
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices '
Imports System.Windows.Forms '
Imports System.Reflection
Imports System.Configuration
Imports System.Transactions

Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.Net
Imports System.Text
Imports System.Diagnostics
Imports System
Imports System.IO.Ports
Imports System.Xml
'Imports System.Web.Services.Protocols
Imports System.Management
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates




Imports GlobalSetting
Imports System.Threading

Public Class frmNewUpload

#Region "Declaration"
    Dim stream As FileStream
    Dim excelReader As IDataReader
    Dim i As Integer
    Dim ii As Integer
    Dim Conn As String


    Dim kanbanNo As String = "", kanbanNo2 As String = "", kanbanNo3 As String = "", kanbanNo4 As String = ""
    Dim kanbanCycle As Integer = 1, kanbanCycle2 As Integer = 2, kanbanCycle3 As Integer = 3, kanbanCycle4 As Integer = 4
    Dim affiliateID As String = "", supplierID As String = ""
    Dim partNo As String = "", pONo As String = "", locationCode As String = ""

    Dim IsProcessing As Boolean = False
    Dim timeLast As DateTime
    Dim LastProcess As Date
    Dim NextProcess As Date = "9999-12-31"
    Dim IntervalProcess As TimeSpan '= TimeSpan.FromSeconds(30)
    Dim ProcessStarted As Boolean = False
    Dim gs_StepProcess As String
    Dim ActiveCycle As Integer
    Dim MaxResponse As Integer
    Dim JamServer As Date
    Dim ProcessCompleted As Boolean = False

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String
    Dim iSplit As Integer
    Dim OrderSplit As String
    Dim statusSPlit As Boolean = False
#End Region

#Region "EDI NEW"
    Dim cls As clsGlobal
    Dim Log As clsLog
    Dim cfg As New clsConfig

    Dim UserLogin = "admin"
    Dim intervalpro As TimeSpan
    Dim processTime As Boolean
    Public SubjectEmail As String = "[TRIAL] "
    Dim tmpPathError As String = "\BACKUP ERROR FILE"

    Dim screenName As String = ""
    Dim templateCode As String = ""
#End Region

#Region "DECRALATION ON OFF"
    Dim IsNewRecEx As Boolean = True
#End Region

#Region "Initialization"
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            cls = New clsGlobal(cfg.ConnectionString, UserLogin)
            Log = New clsLog(cfg.ConnectionString, UserLogin)

            rtbProcess.Text = ""

            txtMsg.Text = ""
            lblDB.Text = "SERVER: [" & Trim(cfg.Server) & "], DATABASE: [" & cfg.Database & "]"

            loadSetting()

            timerProcess.Enabled = True

            If (CDbl(txtTime.Text)) = "0" Then
                timerProcess.Interval = 100
            Else
                timerProcess.Interval = CDbl(txtTime.Text) * 1000 '1 menit
            End If

            txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            intervalpro = TimeSpan.FromSeconds(CDbl(txtTime.Text))
            Dim Last As Date = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")

            btnAuto.Enabled = True
            txtpath.Enabled = False
            txtPathBackup.Enabled = False
            txtTime.Enabled = False
            btnExit.Enabled = True

            processTime = False
        Catch ex As Exception
            cls.up_ShowMsg(ex.Message, txtMsg, GlobalSetting.clsGlobal.MsgTypeEnum.ErrorMsg)
            Log.WriteToErrorLog(Me.Tag, txtMsg.Text, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
        End Try

    End Sub

    Private Sub timerProcess_Tick(sender As Object, e As System.EventArgs) Handles timerProcess.Tick
        If Format(Now, "yyyy-MM-dd HH:mm:ss") = txtNext.Text And processTime = False Then
            Me.Cursor = Cursors.WaitCursor
            UploadProcess()
            Me.Cursor = Cursors.Default
        End If
    End Sub
#End Region

#Region "Procedures"
    Private Sub loadSetting()
        Try
            Dim ls_SQL As String = ""

            ls_SQL = "SELECT ISNULL(AttachmentFolder,'')AttachmentFolder,ISNULL(AttachmentBackupFolder,'')AttachmentBackupFolder, " & vbCrLf & _
                     "ISNULL(Interval,0)Interval FROM dbo.MS_EmailSetting "

            Dim ds As New DataSet
            ds = cls.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                txtpath.Text = ds.Tables(0).Rows(0)("AttachmentFolder")
                txtPathBackup.Text = ds.Tables(0).Rows(0)("AttachmentBackupFolder")
                txtTime.Text = ds.Tables(0).Rows(0)("Interval")
                txtBackupErrorFile.Text = ds.Tables(0).Rows(0)("AttachmentFolder") & tmpPathError
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub clrScreen()
        txtMsg.Text = ""
        txtpath.Text = ""
        txtTime.Text = ""
        txtNext.Text = ""
        txtLast.Text = ""
    End Sub

    ' FILE FORMAT .XLSM
    Private Sub UploadData()
        Dim msgInfo As String = ""
        Dim strFileSize As String = "", excelName As String = ""
        Dim tmpAllNo As String = "", autoApprove As Integer, Status As Integer, Boxdetail As String
        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application
        Dim AllfileNameValid As String = "", AllfileNameInvalid As String = ""
        Dim fileNameValid As String(), fileNameInvalid As String()
        Dim poExportStatus As Integer
        Dim AdaQty As Boolean = False
        Dim postatusExportExist As Boolean = False
        Dim jmlSplit As Integer = 0

        Application.DoEvents()
        msgInfo = "search file upload process (file format .xlsm)"
        gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

        Dim di As New IO.DirectoryInfo(txtpath.Text)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xlsm")
        Dim fi As IO.FileInfo

        Dim jmlFile As Integer = aryFi.Length

        Application.DoEvents()
        gridProcess(rtbProcess, 1, 0, jmlFile & " file upload", True)
        gridProcess(rtbProcess, 1, 0, "End search file upload process (file format .xlsm)" & vbCrLf, True)

        For Each fi In aryFi
            Try

                Dim tmp As New clsTmp
                Dim sheetNumber As Integer = 1
                Dim i As Integer
                Dim z As Integer
                Dim startRow As Long = 0
                Dim inputMaster As Boolean = False, inputMaster2 As Boolean = False, _
                    inputMaster3 As Boolean = False, inputMaster4 As Boolean = False, inputMaster5 As Boolean = False
                Dim statusApprove As String = ""

                Me.Cursor = Cursors.WaitCursor
                txtMsg.Text = ""
                txtMsg.ForeColor = Color.Red


                'For Each fi In aryFi

                Dim ls_file As String = txtpath.Text & "\" & fi.Name
                excelName = fi.Name

                Application.DoEvents()
                msgInfo = "open file upload process..."
                gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

                ExcelBook = xlApp.Workbooks.Open(ls_file)
                ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                Application.DoEvents()
                gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                Dim ds As New DataSet
                Dim ds1 As New DataSet

                Application.DoEvents()
                msgInfo = "read file upload process..."
                gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

                templateCode = ExcelSheet.Range("H1").Value.ToString & ""
                tmp.AffiliateID = ExcelSheet.Range("H3").Value.ToString & ""

                If templateCode <> "TALLY" Then
                    tmp.SupplierID = ExcelSheet.Range("H5").Value.ToString & ""
                End If

                '=========== KHUSUS RECEIVING EXPORT ===============
                Dim recStatus As Boolean = False
                If IsNewRecEx = True Then
                    If templateCode = "REC-EX" Then
                        If ExcelSheet.Range("I28").Value Is Nothing Then
                            tmp.SuratJalanNo = ""
                        Else
                            tmp.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I28").Value.ToString & "", 20)
                        End If

                        If ExcelSheet.Range("AE13").Value Is Nothing Then
                            tmp.PONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                        Else
                            tmp.PONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                        End If

                        If ExcelSheet.Range("AE13").Value Is Nothing Then
                            tmp.OrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                        Else
                            tmp.OrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                        End If

                        If ExcelSheet.Range("AE15").Value Is Nothing Then
                            tmp.OrderNo1 = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                        Else
                            tmp.OrderNo1 = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE15").Value.ToString & "", 20)
                        End If

                        If ExcelSheet.Range("I19").Value Is Nothing Then
                            tmp.ForwarderID = ""
                        Else
                            tmp.ForwarderID = Microsoft.VisualBasic.Left(ExcelSheet.Range("H4").Value.ToString & "", 20)
                        End If

                        If clsTmpDB.CekRecEX(tmp) = True Then
                            recStatus = True
                        Else
                            recStatus = False
                        End If
                    End If
                End If
                '=========== KHUSUS RECEIVING EXPORT ===============

                tmp.DeliveryDate = Format(Now, "yyyy-MM-dd")
                tmp.AffiliateApproveDate = Format(Now, "yyyy-MM-dd")
                tmp.AffiliateApproveUser = "Administrator"
                tmp.SupplierApproveDate = Format(Now, "yyyy-MM-dd")
                tmp.SupplierApproveUser = "Administrator"
                tmp.EntryDate = Format(Now, "yyyy-MM-dd")
                tmp.EntryUser = "Administrator"
                tmp.UpdateUser = "Administrator"
                autoApprove = 0

                If templateCode = "KB" Then
                    startRow = 39
                    tmp.DeliveryLocation = ExcelSheet.Range("H4").Value.ToString & ""
                    If ExcelSheet.Range("I25").Value Is Nothing Then
                        tmp.PIC = ""
                    Else
                        tmp.PIC = Microsoft.VisualBasic.Left(ExcelSheet.Range("I25").Value.ToString & "", 15)
                    End If
                    If ExcelSheet.Range("I27").Value Is Nothing Then
                        tmp.Remarks = ""
                    Else
                        tmp.Remarks = ExcelSheet.Range("I27").Value.ToString()
                    End If
                    tmp.KanbanDate = Format(ExcelSheet.Range("I18").Value, "dd MMM yyyy")
                    kanbanNo = ExcelSheet.Range("AE35").Value.ToString & ""
                    kanbanNo2 = ExcelSheet.Range("AI35").Value.ToString & ""
                    kanbanNo3 = ExcelSheet.Range("AM35").Value.ToString & ""
                    kanbanNo4 = ExcelSheet.Range("AQ35").Value.ToString & ""

                    tmp.KanbanNo = Trim(kanbanNo) & ", " & Trim(kanbanNo2) & ", " & Trim(kanbanNo3) & ", " & Trim(kanbanNo4)

                    Application.DoEvents()
                    tmpAllNo = "KANBAN NO: " & tmp.KanbanNo
                    msgInfo = "processing upload " & tmpAllNo
                    gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

                ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then
                    startRow = 39
                    'tmp.SuratJalanNo = ExcelSheet.Range("J27").Value.ToString & ""
                    If ExcelSheet.Range("J27").Value Is Nothing Then
                        tmp.SuratJalanNo = ""
                    Else
                        tmp.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("J27").Value.ToString & "", 20)
                    End If

                    If ExcelSheet.Range("J25").Value Is Nothing Then
                        tmp.PIC = ""
                    Else
                        tmp.PIC = Microsoft.VisualBasic.Left(ExcelSheet.Range("J25").Value.ToString & "", 15)
                    End If
                    If ExcelSheet.Range("X27").Value Is Nothing Then
                        tmp.JenisArmada = ""
                    Else
                        tmp.JenisArmada = Microsoft.VisualBasic.Left(ExcelSheet.Range("X27").Value.ToString & "", 15)
                    End If
                    If ExcelSheet.Range("J29").Value Is Nothing Then
                        tmp.DriverName = ""
                    Else
                        tmp.DriverName = Microsoft.VisualBasic.Left(ExcelSheet.Range("J29").Value.ToString & "", 15)
                    End If
                    If ExcelSheet.Range("J31").Value Is Nothing Then
                        tmp.DriverCont = ""
                    Else
                        tmp.DriverCont = Microsoft.VisualBasic.Left(ExcelSheet.Range("J31").Value.ToString & "", 15)
                    End If
                    If ExcelSheet.Range("J33").Value Is Nothing Then
                        tmp.NoPol = ""
                    Else
                        tmp.NoPol = Microsoft.VisualBasic.Left(ExcelSheet.Range("J33").Value.ToString & "", 10)
                    End If
                    tmp.TotalBox = ExcelSheet.Range("X29").Value.ToString & ""

                    Application.DoEvents()
                    If templateCode = "DO" Then
                        tmpAllNo = "DO NO: " & tmp.SuratJalanNo
                    Else
                        tmpAllNo = "DO (PO non Kanban) NO: " & tmp.SuratJalanNo
                    End If

                    msgInfo = "processing upload " & tmpAllNo
                    gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

                ElseIf templateCode = "PO" Or templateCode = "POR" Then
                    startRow = 36
                    If ExcelSheet.Range("I16").Value Is Nothing Then
                        tmp.AffiliateName = ""
                    Else
                        tmp.AffiliateName = ExcelSheet.Range("I16").Value.ToString & ""
                    End If
                    If ExcelSheet.Range("AE14").Value Is Nothing Then
                        tmp.ShipCls = ""
                    Else
                        tmp.ShipCls = ExcelSheet.Range("AE14").Value.ToString & ""
                    End If
                    If ExcelSheet.Range("AE12").Value Is Nothing Then
                        tmp.CommercialCls = ""
                    Else
                        If ExcelSheet.Range("AE12").Value.ToString = "YES" Then
                            tmp.CommercialCls = "1"
                        Else
                            tmp.CommercialCls = "0"
                        End If
                    End If
                    Dim tempDate As Date = CDate(ExcelSheet.Range("AE9").Value)
                    tmp.Period = tempDate 'Format(ExcelSheet.Range("AE9").Value, "yyyy-MM-dd")
                    tmp.PONo = ExcelSheet.Range("I9").Value.ToString & ""
                    'tmp.PORevNo = ExcelSheet.Range("Y8").Value.ToString & ""
                    If ExcelSheet.Range("Y8").Value Is Nothing Then
                        tmp.PORevNo = ""
                    Else
                        tmp.PORevNo = ExcelSheet.Range("Y8").Value.ToString & ""
                    End If
                    If ExcelSheet.Range("G1").Value Is Nothing Then
                        statusApprove = "0"
                    Else
                        statusApprove = ExcelSheet.Range("G1").Value.ToString & ""
                    End If

                    Application.DoEvents()
                    If templateCode = "PO" Then
                        tmpAllNo = "PO NO: " & tmp.PONo
                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                    Else
                        tmpAllNo = "POREV NO: " & tmp.PORevNo
                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                    End If

                ElseIf templateCode = "POEM" Then 'PO Monthly
                    startRow = 37
                    'Order No /PO No
                    tmp.PONo = ExcelSheet.Range("S1").Value.ToString & ""
                    If ExcelSheet.Range("P9").Value IsNot Nothing Then
                        If ExcelSheet.Range("P9").Value.ToString & "" = "" Then
                            tmp.OrderNo1 = ExcelSheet.Range("I9").Value.ToString & ""
                            tmp.OrderNo = ExcelSheet.Range("I9").Value.ToString & ""
                        Else
                            tmp.OrderNo1 = ExcelSheet.Range("P9").Value.ToString & ""
                            tmp.OrderNo = ExcelSheet.Range("P9").Value.ToString & ""
                        End If
                    Else
                        tmp.OrderNo1 = ExcelSheet.Range("I9").Value.ToString & ""
                        tmp.OrderNo = ExcelSheet.Range("I9").Value.ToString & ""
                    End If

                    tmp.Period = Format(ExcelSheet.Range("AE9").Value, "yyyy-MM-dd")
                    'Affiliate ID
                    If ExcelSheet.Range("I16").Value Is Nothing Then
                        tmp.AffiliateName = ""
                    Else
                        tmp.AffiliateName = ExcelSheet.Range("I16").Value.ToString & ""
                    End If

                    'Supplier ID
                    If ExcelSheet.Range("I11").Value Is Nothing Then
                        tmp.SupplierID = ""
                    Else
                        tmp.SupplierID = ExcelSheet.Range("I11").Value.ToString & ""
                    End If

                    'Forwarder ID
                    If ExcelSheet.Range("AE13").Value Is Nothing Then
                        tmp.ForwarderID = ""
                    Else
                        tmp.ForwarderID = ExcelSheet.Range("AE13").Value.ToString & ""
                    End If

                    'ETD Vendor 1
                    If ExcelSheet.Range("Z37").Value Is Nothing Then
                        tmp.ETDVendor1 = ""
                    Else
                        tmp.ETDVendor1 = CDate(ExcelSheet.Range("Z37").Value)
                    End If

                    'Remarks
                    If ExcelSheet.Range("J27").Value Is Nothing Then
                        tmp.Remarks = ""
                    Else
                        tmp.Remarks = ExcelSheet.Range("J27").Value.ToString & ""
                    End If

                    'StatusApprove
                    If ExcelSheet.Range("G1").Value Is Nothing Then
                        statusApprove = "0"
                    Else
                        statusApprove = ExcelSheet.Range("G1").Value.ToString & ""
                    End If

                    Application.DoEvents()
                    If templateCode = "POEM" Then
                        tmpAllNo = "ORDER NO: " & tmp.PONo
                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                    End If

                ElseIf templateCode = "POEE" Then 'PO Emergency
                    startRow = 37
                    'Order No /PO No
                    tmp.PONo = ExcelSheet.Range("S1").Value.ToString & ""
                    tmp.OrderNo1 = ExcelSheet.Range("I9").Value.ToString & ""
                    tmp.OrderNo = ExcelSheet.Range("I9").Value.ToString & ""

                    'Affiliate ID
                    If ExcelSheet.Range("H3").Value Is Nothing Then
                        tmp.AffiliateName = ""
                    Else
                        tmp.AffiliateName = ExcelSheet.Range("H3").Value.ToString & ""
                    End If

                    'Supplier ID
                    If ExcelSheet.Range("H5").Value Is Nothing Then
                        tmp.SupplierID = ""
                    Else
                        tmp.SupplierID = ExcelSheet.Range("H5").Value.ToString & ""
                    End If

                    'Forwarder ID
                    If ExcelSheet.Range("H4").Value Is Nothing Then
                        tmp.ForwarderID = ""
                    Else
                        tmp.ForwarderID = ExcelSheet.Range("H4").Value.ToString & ""
                    End If

                    'ETD Vendor 1
                    If ExcelSheet.Range("Z37").Value Is Nothing Then
                        tmp.ETDVendor1 = ""
                    Else
                        tmp.ETDVendor1 = CDate(ExcelSheet.Range("Z37").Value)
                    End If

                    'Remarks
                    If ExcelSheet.Range("J27").Value Is Nothing Then
                        tmp.Remarks = ""
                    Else
                        tmp.Remarks = ExcelSheet.Range("J27").Value.ToString & ""
                    End If

                    'StatusApprove
                    If ExcelSheet.Range("G1").Value Is Nothing Then
                        statusApprove = "0"
                    Else
                        statusApprove = ExcelSheet.Range("G1").Value.ToString & ""
                    End If

                    Application.DoEvents()
                    If templateCode = "POEE" Then
                        tmpAllNo = "ORDER NO: " & tmp.PONo
                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                    End If
                ElseIf templateCode = "INV" Then
                    startRow = 36
                    tmp.InvoiceNo = ExcelSheet.Range("I11").Value.ToString & ""
                    tmp.InvoiceDate = Format(ExcelSheet.Range("W11").Value, "yyyy-MM-dd")
                    If ExcelSheet.Range("AM11").Value Is Nothing Then
                        tmp.PaymentItem = ""
                    Else
                        tmp.PaymentItem = ExcelSheet.Range("AM11").Value.ToString & ""
                    End If
                    tmp.DueDate = Format(ExcelSheet.Range("BA11").Value, "yyyy-MM-dd")
                    If ExcelSheet.Range("I25").Value Is Nothing Then
                        tmp.PIC = ""
                    Else
                        tmp.PIC = Microsoft.VisualBasic.Left(ExcelSheet.Range("I25").Value.ToString & "", 15)
                    End If
                    If ExcelSheet.Range("I27").Value Is Nothing Then
                        tmp.Remarks = ""
                    Else
                        tmp.Remarks = ExcelSheet.Range("I27").Value.ToString & ""
                    End If

                    Application.DoEvents()
                    tmpAllNo = "INVOICE NO: " & tmp.InvoiceNo
                    msgInfo = "processing upload " & tmpAllNo
                    gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                ElseIf templateCode = "DO-EX" Then 'DIAN EXPORT
                    startRow = 34

                    If ExcelSheet.Range("I28").Value Is Nothing Then
                        tmp.SuratJalanNo = ""
                    Else
                        tmp.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I28").Value.ToString & "", 20)
                    End If

                    If ExcelSheet.Range("AE13").Value Is Nothing Then
                        tmp.OrderNo = ""
                    Else
                        tmp.OrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                    End If

                    Application.DoEvents()
                    If templateCode = "DO-EX" Then
                        tmpAllNo = "DO NO: " & tmp.SuratJalanNo
                    End If

                    msgInfo = "processing upload " & tmpAllNo
                    gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                ElseIf templateCode = "REC-EX" Then 'DIAN EXPORT
                    startRow = 34

                    If ExcelSheet.Range("I28").Value Is Nothing Then
                        tmp.SuratJalanNo = ""
                    Else
                        tmp.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I28").Value.ToString & "", 20)
                    End If

                    If ExcelSheet.Range("AE13").Value Is Nothing Then
                        tmp.PONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                    Else
                        tmp.PONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                    End If

                    If ExcelSheet.Range("AE13").Value Is Nothing Then
                        tmp.OrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                    Else
                        tmp.OrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                    End If

                    If ExcelSheet.Range("AE15").Value Is Nothing Then
                        tmp.OrderNo1 = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                    Else
                        tmp.OrderNo1 = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE15").Value.ToString & "", 20)
                    End If

                    If ExcelSheet.Range("I19").Value Is Nothing Then
                        tmp.ForwarderID = ""
                    Else
                        tmp.ForwarderID = Microsoft.VisualBasic.Left(ExcelSheet.Range("H4").Value.ToString & "", 20)
                    End If

                    Application.DoEvents()
                    If templateCode = "REC-EX" Then
                        tmpAllNo = "SJ NO: " & tmp.SuratJalanNo
                    End If

                    msgInfo = "processing upload " & tmpAllNo
                    gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                ElseIf templateCode = "INV-EX" Then 'DIAN Export
                    startRow = 36
                    tmp.InvoiceNo = ExcelSheet.Range("I11").Value.ToString & ""
                    tmp.InvoiceDate = Format(ExcelSheet.Range("W11").Value, "yyyy-MM-dd")
                    If ExcelSheet.Range("AM11").Value Is Nothing Then
                        tmp.PaymentItem = ""
                    Else
                        tmp.PaymentItem = ExcelSheet.Range("AM11").Value.ToString
                    End If
                    tmp.DueDate = Format(ExcelSheet.Range("BA11").Value, "yyyy-MM-dd")
                    If ExcelSheet.Range("I25").Value Is Nothing Then
                        tmp.PIC = ""
                    Else
                        tmp.PIC = Microsoft.VisualBasic.Left(ExcelSheet.Range("I25").Value.ToString & "", 15)
                    End If
                    If ExcelSheet.Range("I27").Value Is Nothing Then
                        tmp.Remarks = ""
                    Else
                        tmp.Remarks = ExcelSheet.Range("I27").Value.ToString & ""
                    End If

                    Application.DoEvents()
                    tmpAllNo = "INVOICE NO: " & tmp.InvoiceNo
                    msgInfo = "processing upload " & tmpAllNo
                    gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                ElseIf templateCode = "TALLY" Then 'DIAN Export
                    startRow = 22
                    tmp.ForwarderID = ExcelSheet.Range("H3").Value.ToString & ""
                    tmp.AffiliateID = ExcelSheet.Range("S1").Value.ToString & ""

                    tmp.InvoiceNo = IIf(IsNothing(ExcelSheet.Range("I7").Value), "", ExcelSheet.Range("I7").Value)
                    tmp.Vassel = IIf(IsNothing(ExcelSheet.Range("AA7").Value), "", ExcelSheet.Range("AA7").Value)
                    tmp.NamaKapal = IIf(IsNothing(ExcelSheet.Range("AQ7").Value), "", ExcelSheet.Range("AQ7").Value)
                    tmp.StuffingDate = IIf(IsNothing(ExcelSheet.Range("AQ9").Value), "", ExcelSheet.Range("AQ9").Value)

                    tmp.ContainerNo = IIf(IsNothing(ExcelSheet.Range("I9").Value), "", ExcelSheet.Range("I9").Value)
                    tmp.DONo = IIf(IsNothing(ExcelSheet.Range("AA9").Value), "", ExcelSheet.Range("AA9").Value)

                    tmp.SealNo = IIf(IsNothing(ExcelSheet.Range("I11").Value), "", ExcelSheet.Range("I11").Value)
                    tmp.SizeContainer = Replace(IIf(IsDBNull(ExcelSheet.Range("AA11").Value), "", ExcelSheet.Range("AA11").Value), "'", " ")

                    tmp.Tare = IIf(IsNothing(ExcelSheet.Range("I13").Value), 0, ExcelSheet.Range("I13").Value)
                    tmp.ETDJakarta = IIf(IsNothing(ExcelSheet.Range("AA13").Value), "", ExcelSheet.Range("AA13").Value)

                    tmp.Gross = IIf(IsNothing(ExcelSheet.Range("I15").Value), 0, ExcelSheet.Range("I15").Value)
                    tmp.ShippingLine = IIf(IsNothing(ExcelSheet.Range("AA15").Value), "", ExcelSheet.Range("AA15").Value)

                    tmp.TotalCarton = IIf(IsNothing(ExcelSheet.Range("I17").Value), 0, ExcelSheet.Range("I17").Value)
                    tmp.DestinationPort = IIf(IsNothing(ExcelSheet.Range("AA17").Value), "", ExcelSheet.Range("AA17").Value)

                    Application.DoEvents()
                    tmpAllNo = "Tally: " & tmp.InvoiceNo
                    msgInfo = "processing upload " & tmpAllNo
                    gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)
                End If

                Try
                    Using SQLCon As New SqlConnection(uf_GetConString)
                        SQLCon.Open()

                        Dim SQLCom As SqlCommand = SQLCon.CreateCommand
                        Dim SQLTrans As SqlTransaction
                        Dim StartCol As Integer

                        SQLTrans = SQLCon.BeginTransaction
                        SQLCom.Connection = SQLCon
                        SQLCom.Transaction = SQLTrans
                        z = 0
                        If templateCode <> "POEM" Then StartCol = 4
                        For z = StartCol To 4
                            AdaQty = False
                            For i = startRow To 10000
                                If ExcelSheet.Range("B" & i).Value.ToString = "E" Then
                                    Application.DoEvents()
                                    gridProcess(rtbProcess, 1, 0, "End " & msgInfo, True)

                                    Application.DoEvents()
                                    msgInfo = "read file upload process..."
                                    gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                                    Me.Cursor = Cursors.Default
                                    Exit For
                                ElseIf ExcelSheet.Range("D" & i).Value Is Nothing Then
                                    txtMsg.Text = ""
                                    Me.Cursor = Cursors.Default
                                    Exit For
                                End If

                                If templateCode = "KB" Then

                                    autoApprove = clsTmpDB.CekKanbanAutoApprover(tmp, SQLCom)
                                    If autoApprove = 0 Then
                                        clsTmpDB.UpdateKanban(tmp, SQLCom)
                                    Else
                                        Application.DoEvents()
                                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo, True)

                                        Application.DoEvents()
                                        msgInfo = "read file upload process..."
                                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                                        Me.Cursor = Cursors.Default
                                        Exit For
                                    End If

                                ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then

                                    tmp.PartNo = ExcelSheet.Range("P" & i).Value.ToString & ""
                                    tmp.PONo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                    tmp.UnitCls = clsTmpDB.UnitCls(ExcelSheet.Range("AD" & i).Value.ToString & "")
                                    tmp.KanbanNo = ExcelSheet.Range("L" & i).Value.ToString & ""
                                    tmp.DOQty = IIf(IsNumeric(ExcelSheet.Range("AO" & i).Value) = False, 0, ExcelSheet.Range("AO" & i).Value)
                                    If tmp.DOQty > 0 Then
                                        tmp.POKanbanCls = clsTmpDB.POKanbanCls(tmp.PONo, tmp.PartNo, tmp.AffiliateID, tmp.SupplierID)
                                        If inputMaster2 = False Then
                                            clsTmpDB.insertMasterDO(tmp, SQLCom)
                                            inputMaster2 = True
                                        End If
                                        clsTmpDB.insertDetailDO(tmp)
                                    End If

                                ElseIf templateCode = "PO" Then 'PO

                                    tmp.PartNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                    If ExcelSheet.Range("AC" & i).Value = ExcelSheet.Range("AC" & i + 1).Value Then
                                        tmp.DifferentCls = "0"
                                    Else
                                        tmp.DifferentCls = "1"
                                    End If
                                    tmp.POKanbanCls = IIf(ExcelSheet.Range("Q" & i).Value.ToString = "YES", "1", "0")
                                    tmp.POQty = Val(ExcelSheet.Range("AX" & i + 1).Value) + Val(ExcelSheet.Range("AZ" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BB" & i + 1).Value) + Val(ExcelSheet.Range("BD" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BF" & i + 1).Value) + Val(ExcelSheet.Range("BH" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BJ" & i + 1).Value) + Val(ExcelSheet.Range("BL" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BN" & i + 1).Value) + Val(ExcelSheet.Range("BP" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BR" & i + 1).Value) + Val(ExcelSheet.Range("BT" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BV" & i + 1).Value) + Val(ExcelSheet.Range("BX" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BZ" & i + 1).Value) + Val(ExcelSheet.Range("CB" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CD" & i + 1).Value) + Val(ExcelSheet.Range("CF" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CH" & i + 1).Value) + Val(ExcelSheet.Range("CJ" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CL" & i + 1).Value) + Val(ExcelSheet.Range("CN" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CP" & i + 1).Value) + Val(ExcelSheet.Range("CR" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CT" & i + 1).Value) + Val(ExcelSheet.Range("CV" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CX" & i + 1).Value) + Val(ExcelSheet.Range("CZ" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("DB" & i + 1).Value) + Val(ExcelSheet.Range("DD" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("DF" & i + 1).Value)

                                    tmp.POQtyOld = Val(ExcelSheet.Range("AX" & i).Value) + Val(ExcelSheet.Range("AZ" & i).Value) _
                                                + Val(ExcelSheet.Range("BB" & i).Value) + Val(ExcelSheet.Range("BD" & i).Value) _
                                                + Val(ExcelSheet.Range("BF" & i).Value) + Val(ExcelSheet.Range("BH" & i).Value) _
                                                + Val(ExcelSheet.Range("BJ" & i).Value) + Val(ExcelSheet.Range("BL" & i).Value) _
                                                + Val(ExcelSheet.Range("BN" & i).Value) + Val(ExcelSheet.Range("BP" & i).Value) _
                                                + Val(ExcelSheet.Range("BR" & i).Value) + Val(ExcelSheet.Range("BT" & i).Value) _
                                                + Val(ExcelSheet.Range("BV" & i).Value) + Val(ExcelSheet.Range("BX" & i).Value) _
                                                + Val(ExcelSheet.Range("BZ" & i).Value) + Val(ExcelSheet.Range("CB" & i).Value) _
                                                + Val(ExcelSheet.Range("CD" & i).Value) + Val(ExcelSheet.Range("CF" & i).Value) _
                                                + Val(ExcelSheet.Range("CH" & i).Value) + Val(ExcelSheet.Range("CJ" & i).Value) _
                                                + Val(ExcelSheet.Range("CL" & i).Value) + Val(ExcelSheet.Range("CN" & i).Value) _
                                                + Val(ExcelSheet.Range("CP" & i).Value) + Val(ExcelSheet.Range("CR" & i).Value) _
                                                + Val(ExcelSheet.Range("CT" & i).Value) + Val(ExcelSheet.Range("CV" & i).Value) _
                                                + Val(ExcelSheet.Range("CX" & i).Value) + Val(ExcelSheet.Range("CZ" & i).Value) _
                                                + Val(ExcelSheet.Range("DB" & i).Value) + Val(ExcelSheet.Range("DD" & i).Value) _
                                                + Val(ExcelSheet.Range("DF" & i).Value)

                                    tmp.CurrCls = ""
                                    tmp.Price = 0
                                    tmp.Amount = 0
                                    If ExcelSheet.Range("J25").Value Is Nothing Then
                                        tmp.SupplierApproveUser = ""
                                    Else
                                        tmp.SupplierApproveUser = Microsoft.VisualBasic.Left(ExcelSheet.Range("J25").Value.ToString & "", 15)
                                    End If
                                    If ExcelSheet.Range("J27").Value Is Nothing Then
                                        tmp.Remarks = ""
                                    Else
                                        tmp.Remarks = ExcelSheet.Range("J27").Value.ToString()
                                    End If

                                    autoApprove = clsTmpDB.CekPOAutoApprover("dbo.PO_Master", tmp.PONo, tmp.SupplierID, statusApprove)
                                    If autoApprove = 0 Then
                                        If inputMaster3 = False Then
                                            clsTmpDB.insertMasterPOUpload(tmp, SQLCom, "dbo.PO_MasterUpload", templateCode)
                                            clsTmpDB.UpdatePOMaster(tmp, SQLCom, statusApprove, "dbo.PO_Master")
                                            clsTmpDB.UpdatePOMasterUpload(tmp, SQLCom, "dbo.PO_MasterUpload", templateCode)
                                            inputMaster3 = True
                                        End If

                                        If statusApprove = "0" Then
                                            clsTmpDB.insertDetailPOUpload(tmp, _
                                            ExcelSheet.Range("AX" & i + 1).Value, ExcelSheet.Range("AX" & i).Value, ExcelSheet.Range("AZ" & i + 1).Value, ExcelSheet.Range("AZ" & i).Value, ExcelSheet.Range("BB" & i + 1).Value, ExcelSheet.Range("BB" & i).Value, ExcelSheet.Range("BD" & i + 1).Value, ExcelSheet.Range("BD" & i).Value, _
                                            ExcelSheet.Range("BF" & i + 1).Value, ExcelSheet.Range("BF" & i).Value, ExcelSheet.Range("BH" & i + 1).Value, ExcelSheet.Range("BH" & i).Value, ExcelSheet.Range("BJ" & i + 1).Value, ExcelSheet.Range("BJ" & i).Value, ExcelSheet.Range("BL" & i + 1).Value, ExcelSheet.Range("BL" & i).Value, _
                                            ExcelSheet.Range("BN" & i + 1).Value, ExcelSheet.Range("BN" & i).Value, ExcelSheet.Range("BP" & i + 1).Value, ExcelSheet.Range("BP" & i).Value, ExcelSheet.Range("BR" & i + 1).Value, ExcelSheet.Range("BR" & i).Value, ExcelSheet.Range("BT" & i + 1).Value, ExcelSheet.Range("BT" & i).Value, _
                                            ExcelSheet.Range("BV" & i + 1).Value, ExcelSheet.Range("BV" & i).Value, ExcelSheet.Range("BX" & i + 1).Value, ExcelSheet.Range("BX" & i).Value, ExcelSheet.Range("BZ" & i + 1).Value, ExcelSheet.Range("BZ" & i).Value, ExcelSheet.Range("CB" & i + 1).Value, ExcelSheet.Range("CB" & i).Value, _
                                            ExcelSheet.Range("CD" & i + 1).Value, ExcelSheet.Range("CD" & i).Value, ExcelSheet.Range("CF" & i + 1).Value, ExcelSheet.Range("CF" & i).Value, ExcelSheet.Range("CH" & i + 1).Value, ExcelSheet.Range("CH" & i).Value, ExcelSheet.Range("CJ" & i + 1).Value, ExcelSheet.Range("CJ" & i).Value, _
                                            ExcelSheet.Range("CL" & i + 1).Value, ExcelSheet.Range("CL" & i).Value, ExcelSheet.Range("CN" & i + 1).Value, ExcelSheet.Range("CN" & i).Value, ExcelSheet.Range("CP" & i + 1).Value, ExcelSheet.Range("CP" & i).Value, ExcelSheet.Range("CR" & i + 1).Value, ExcelSheet.Range("CR" & i).Value, _
                                            ExcelSheet.Range("CT" & i + 1).Value, ExcelSheet.Range("CT" & i).Value, ExcelSheet.Range("CV" & i + 1).Value, ExcelSheet.Range("CV" & i).Value, ExcelSheet.Range("CX" & i + 1).Value, ExcelSheet.Range("CX" & i).Value, ExcelSheet.Range("CZ" & i + 1).Value, ExcelSheet.Range("CZ" & i).Value, _
                                            ExcelSheet.Range("DB" & i + 1).Value, ExcelSheet.Range("DB" & i).Value, ExcelSheet.Range("DD" & i + 1).Value, ExcelSheet.Range("DD" & i).Value, ExcelSheet.Range("DF" & i + 1).Value, ExcelSheet.Range("DF" & i).Value, "dbo.PO_DetailUpload", templateCode)
                                        ElseIf statusApprove = "1" Then
                                            If ExcelSheet.Range("AG" & i).Value.ToString & "" = "YES" Then
                                                clsTmpDB.insertDetailPOUpload(tmp, _
                                                ExcelSheet.Range("AX" & i + 1).Value, ExcelSheet.Range("AX" & i).Value, ExcelSheet.Range("AZ" & i + 1).Value, ExcelSheet.Range("AZ" & i).Value, ExcelSheet.Range("BB" & i + 1).Value, ExcelSheet.Range("BB" & i).Value, ExcelSheet.Range("BD" & i + 1).Value, ExcelSheet.Range("BD" & i).Value, _
                                                ExcelSheet.Range("BF" & i + 1).Value, ExcelSheet.Range("BF" & i).Value, ExcelSheet.Range("BH" & i + 1).Value, ExcelSheet.Range("BH" & i).Value, ExcelSheet.Range("BJ" & i + 1).Value, ExcelSheet.Range("BJ" & i).Value, ExcelSheet.Range("BL" & i + 1).Value, ExcelSheet.Range("BL" & i).Value, _
                                                ExcelSheet.Range("BN" & i + 1).Value, ExcelSheet.Range("BN" & i).Value, ExcelSheet.Range("BP" & i + 1).Value, ExcelSheet.Range("BP" & i).Value, ExcelSheet.Range("BR" & i + 1).Value, ExcelSheet.Range("BR" & i).Value, ExcelSheet.Range("BT" & i + 1).Value, ExcelSheet.Range("BT" & i).Value, _
                                                ExcelSheet.Range("BV" & i + 1).Value, ExcelSheet.Range("BV" & i).Value, ExcelSheet.Range("BX" & i + 1).Value, ExcelSheet.Range("BX" & i).Value, ExcelSheet.Range("BZ" & i + 1).Value, ExcelSheet.Range("BZ" & i).Value, ExcelSheet.Range("CB" & i + 1).Value, ExcelSheet.Range("CB" & i).Value, _
                                                ExcelSheet.Range("CD" & i + 1).Value, ExcelSheet.Range("CD" & i).Value, ExcelSheet.Range("CF" & i + 1).Value, ExcelSheet.Range("CF" & i).Value, ExcelSheet.Range("CH" & i + 1).Value, ExcelSheet.Range("CH" & i).Value, ExcelSheet.Range("CJ" & i + 1).Value, ExcelSheet.Range("CJ" & i).Value, _
                                                ExcelSheet.Range("CL" & i + 1).Value, ExcelSheet.Range("CL" & i).Value, ExcelSheet.Range("CN" & i + 1).Value, ExcelSheet.Range("CN" & i).Value, ExcelSheet.Range("CP" & i + 1).Value, ExcelSheet.Range("CP" & i).Value, ExcelSheet.Range("CR" & i + 1).Value, ExcelSheet.Range("CR" & i).Value, _
                                                ExcelSheet.Range("CT" & i + 1).Value, ExcelSheet.Range("CT" & i).Value, ExcelSheet.Range("CV" & i + 1).Value, ExcelSheet.Range("CV" & i).Value, ExcelSheet.Range("CX" & i + 1).Value, ExcelSheet.Range("CX" & i).Value, ExcelSheet.Range("CZ" & i + 1).Value, ExcelSheet.Range("CZ" & i).Value, _
                                                ExcelSheet.Range("DB" & i + 1).Value, ExcelSheet.Range("DB" & i).Value, ExcelSheet.Range("DD" & i + 1).Value, ExcelSheet.Range("DD" & i).Value, ExcelSheet.Range("DF" & i + 1).Value, ExcelSheet.Range("DF" & i).Value, "dbo.PO_DetailUpload", templateCode)
                                            ElseIf ExcelSheet.Range("AG" & i).Value.ToString & "" = "NO" Then
                                                clsTmpDB.insertDetailPOUpload(tmp, _
                                                0, ExcelSheet.Range("AX" & i).Value, 0, ExcelSheet.Range("AZ" & i).Value, 0, ExcelSheet.Range("BB" & i).Value, 0, ExcelSheet.Range("BD" & i).Value, _
                                                0, ExcelSheet.Range("BF" & i).Value, 0, ExcelSheet.Range("BH" & i).Value, 0, ExcelSheet.Range("BJ" & i).Value, 0, ExcelSheet.Range("BL" & i).Value, _
                                                0, ExcelSheet.Range("BN" & i).Value, 0, ExcelSheet.Range("BP" & i).Value, 0, ExcelSheet.Range("BR" & i).Value, 0, ExcelSheet.Range("BT" & i).Value, _
                                                0, ExcelSheet.Range("BV" & i).Value, 0, ExcelSheet.Range("BX" & i).Value, 0, ExcelSheet.Range("BZ" & i).Value, 0, ExcelSheet.Range("CB" & i).Value, _
                                                0, ExcelSheet.Range("CD" & i).Value, 0, ExcelSheet.Range("CF" & i).Value, 0, ExcelSheet.Range("CH" & i).Value, 0, ExcelSheet.Range("CJ" & i).Value, _
                                                0, ExcelSheet.Range("CL" & i).Value, 0, ExcelSheet.Range("CN" & i).Value, 0, ExcelSheet.Range("CP" & i).Value, 0, ExcelSheet.Range("CR" & i).Value, _
                                                0, ExcelSheet.Range("CT" & i).Value, 0, ExcelSheet.Range("CV" & i).Value, 0, ExcelSheet.Range("CX" & i).Value, 0, ExcelSheet.Range("CZ" & i).Value, _
                                                0, ExcelSheet.Range("DB" & i).Value, 0, ExcelSheet.Range("DD" & i).Value, 0, ExcelSheet.Range("DF" & i).Value, "dbo.PO_DetailUpload", templateCode)
                                            End If
                                        ElseIf statusApprove = "2" Then
                                            clsTmpDB.insertDetailPOUpload(tmp, _
                                                0, ExcelSheet.Range("AX" & i).Value, 0, ExcelSheet.Range("AZ" & i).Value, 0, ExcelSheet.Range("BB" & i).Value, 0, ExcelSheet.Range("BD" & i).Value, _
                                                0, ExcelSheet.Range("BF" & i).Value, 0, ExcelSheet.Range("BH" & i).Value, 0, ExcelSheet.Range("BJ" & i).Value, 0, ExcelSheet.Range("BL" & i).Value, _
                                                0, ExcelSheet.Range("BN" & i).Value, 0, ExcelSheet.Range("BP" & i).Value, 0, ExcelSheet.Range("BR" & i).Value, 0, ExcelSheet.Range("BT" & i).Value, _
                                                0, ExcelSheet.Range("BV" & i).Value, 0, ExcelSheet.Range("BX" & i).Value, 0, ExcelSheet.Range("BZ" & i).Value, 0, ExcelSheet.Range("CB" & i).Value, _
                                                0, ExcelSheet.Range("CD" & i).Value, 0, ExcelSheet.Range("CF" & i).Value, 0, ExcelSheet.Range("CH" & i).Value, 0, ExcelSheet.Range("CJ" & i).Value, _
                                                0, ExcelSheet.Range("CL" & i).Value, 0, ExcelSheet.Range("CN" & i).Value, 0, ExcelSheet.Range("CP" & i).Value, 0, ExcelSheet.Range("CR" & i).Value, _
                                                0, ExcelSheet.Range("CT" & i).Value, 0, ExcelSheet.Range("CV" & i).Value, 0, ExcelSheet.Range("CX" & i).Value, 0, ExcelSheet.Range("CZ" & i).Value, _
                                                0, ExcelSheet.Range("DB" & i).Value, 0, ExcelSheet.Range("DD" & i).Value, 0, ExcelSheet.Range("DF" & i).Value, "dbo.PO_DetailUpload", templateCode)
                                        End If

                                        i = i + 1

                                    Else
                                        Application.DoEvents()
                                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo, True)

                                        Application.DoEvents()
                                        msgInfo = "read file upload process..."
                                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                                        Me.Cursor = Cursors.Default
                                        Exit For

                                    End If

                                ElseIf templateCode = "POR" Then 'PO REVISION

                                    tmp.POSeqNo = clsTmpDB.POSeqNo(tmp.PORevNo)
                                    tmp.PartNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                    If ExcelSheet.Range("AC" & i + 2).Value = ExcelSheet.Range("AC" & i + 1).Value Then
                                        tmp.DifferentCls = "0"
                                    Else
                                        tmp.DifferentCls = "1"
                                    End If
                                    tmp.POKanbanCls = IIf(ExcelSheet.Range("Q" & i).Value.ToString = "YES", "1", "0")
                                    tmp.POQty = Val(ExcelSheet.Range("AX" & i + 2).Value) + Val(ExcelSheet.Range("AZ" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("BB" & i + 2).Value) + Val(ExcelSheet.Range("BD" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("BF" & i + 2).Value) + Val(ExcelSheet.Range("BH" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("BJ" & i + 2).Value) + Val(ExcelSheet.Range("BL" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("BN" & i + 2).Value) + Val(ExcelSheet.Range("BP" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("BR" & i + 2).Value) + Val(ExcelSheet.Range("BT" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("BV" & i + 2).Value) + Val(ExcelSheet.Range("BX" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("BZ" & i + 2).Value) + Val(ExcelSheet.Range("CB" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("CD" & i + 2).Value) + Val(ExcelSheet.Range("CF" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("CH" & i + 2).Value) + Val(ExcelSheet.Range("CJ" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("CL" & i + 2).Value) + Val(ExcelSheet.Range("CN" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("CP" & i + 2).Value) + Val(ExcelSheet.Range("CR" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("CT" & i + 2).Value) + Val(ExcelSheet.Range("CV" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("CX" & i + 2).Value) + Val(ExcelSheet.Range("CZ" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("DB" & i + 2).Value) + Val(ExcelSheet.Range("DD" & i + 2).Value) _
                                                + Val(ExcelSheet.Range("DF" & i + 2).Value)

                                    tmp.POQtyOld = Val(ExcelSheet.Range("AX" & i + 1).Value) + Val(ExcelSheet.Range("AZ" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BB" & i + 1).Value) + Val(ExcelSheet.Range("BD" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BF" & i + 1).Value) + Val(ExcelSheet.Range("BH" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BJ" & i + 1).Value) + Val(ExcelSheet.Range("BL" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BN" & i + 1).Value) + Val(ExcelSheet.Range("BP" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BR" & i + 1).Value) + Val(ExcelSheet.Range("BT" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BV" & i + 1).Value) + Val(ExcelSheet.Range("BX" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("BZ" & i + 1).Value) + Val(ExcelSheet.Range("CB" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CD" & i + 1).Value) + Val(ExcelSheet.Range("CF" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CH" & i + 1).Value) + Val(ExcelSheet.Range("CJ" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CL" & i + 1).Value) + Val(ExcelSheet.Range("CN" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CP" & i + 1).Value) + Val(ExcelSheet.Range("CR" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CT" & i + 1).Value) + Val(ExcelSheet.Range("CV" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("CX" & i + 1).Value) + Val(ExcelSheet.Range("CZ" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("DB" & i + 1).Value) + Val(ExcelSheet.Range("DD" & i + 1).Value) _
                                                + Val(ExcelSheet.Range("DF" & i + 1).Value)

                                    tmp.CurrCls = ""
                                    tmp.Price = 0
                                    tmp.Amount = 0
                                    If ExcelSheet.Range("J25").Value Is Nothing Then
                                        tmp.SupplierApproveUser = ""
                                    Else
                                        tmp.SupplierApproveUser = Microsoft.VisualBasic.Left(ExcelSheet.Range("J25").Value.ToString & "", 15)
                                    End If
                                    If ExcelSheet.Range("J27").Value Is Nothing Then
                                        tmp.Remarks = ""
                                    Else
                                        tmp.Remarks = ExcelSheet.Range("J27").Value.ToString()
                                    End If

                                    autoApprove = clsTmpDB.CekPOAutoApprover("dbo.PO_Master", tmp.PONo, tmp.SupplierID, statusApprove)
                                    If autoApprove = 0 Then
                                        If inputMaster4 = False Then
                                            clsTmpDB.insertMasterPOUpload(tmp, SQLCom, "dbo.PORev_MasterUpload", templateCode)
                                            clsTmpDB.UpdatePOMaster(tmp, SQLCom, statusApprove, "dbo.PORev_Master")
                                            clsTmpDB.UpdatePOMasterUpload(tmp, SQLCom, "dbo.PORev_MasterUpload", templateCode)
                                            inputMaster4 = True
                                        End If
                                        If statusApprove = "0" Then
                                            clsTmpDB.insertDetailPOUpload(tmp, _
                                            ExcelSheet.Range("AX" & i + 2).Value, ExcelSheet.Range("AX" & i + 1).Value, ExcelSheet.Range("AZ" & i + 2).Value, ExcelSheet.Range("AZ" & i + 1).Value, ExcelSheet.Range("BB" & i + 2).Value, ExcelSheet.Range("BB" & i + 1).Value, ExcelSheet.Range("BD" & i + 2).Value, ExcelSheet.Range("BD" & i + 1).Value, _
                                            ExcelSheet.Range("BF" & i + 2).Value, ExcelSheet.Range("BF" & i + 1).Value, ExcelSheet.Range("BH" & i + 2).Value, ExcelSheet.Range("BH" & i + 1).Value, ExcelSheet.Range("BJ" & i + 2).Value, ExcelSheet.Range("BJ" & i + 1).Value, ExcelSheet.Range("BL" & i + 2).Value, ExcelSheet.Range("BL" & i + 1).Value, _
                                            ExcelSheet.Range("BN" & i + 2).Value, ExcelSheet.Range("BN" & i + 1).Value, ExcelSheet.Range("BP" & i + 2).Value, ExcelSheet.Range("BP" & i + 1).Value, ExcelSheet.Range("BR" & i + 2).Value, ExcelSheet.Range("BR" & i + 1).Value, ExcelSheet.Range("BT" & i + 2).Value, ExcelSheet.Range("BT" & i + 1).Value, _
                                            ExcelSheet.Range("BV" & i + 2).Value, ExcelSheet.Range("BV" & i + 1).Value, ExcelSheet.Range("BX" & i + 2).Value, ExcelSheet.Range("BX" & i + 1).Value, ExcelSheet.Range("BZ" & i + 2).Value, ExcelSheet.Range("BZ" & i + 1).Value, ExcelSheet.Range("CB" & i + 2).Value, ExcelSheet.Range("CB" & i + 1).Value, _
                                            ExcelSheet.Range("CD" & i + 2).Value, ExcelSheet.Range("CD" & i + 1).Value, ExcelSheet.Range("CF" & i + 2).Value, ExcelSheet.Range("CF" & i + 1).Value, ExcelSheet.Range("CH" & i + 2).Value, ExcelSheet.Range("CH" & i + 1).Value, ExcelSheet.Range("CJ" & i + 2).Value, ExcelSheet.Range("CJ" & i + 1).Value, _
                                            ExcelSheet.Range("CL" & i + 2).Value, ExcelSheet.Range("CL" & i + 1).Value, ExcelSheet.Range("CN" & i + 2).Value, ExcelSheet.Range("CN" & i + 1).Value, ExcelSheet.Range("CP" & i + 2).Value, ExcelSheet.Range("CP" & i + 1).Value, ExcelSheet.Range("CR" & i + 2).Value, ExcelSheet.Range("CR" & i + 1).Value, _
                                            ExcelSheet.Range("CT" & i + 2).Value, ExcelSheet.Range("CT" & i + 1).Value, ExcelSheet.Range("CV" & i + 2).Value, ExcelSheet.Range("CV" & i + 1).Value, ExcelSheet.Range("CX" & i + 2).Value, ExcelSheet.Range("CX" & i + 1).Value, ExcelSheet.Range("CZ" & i + 2).Value, ExcelSheet.Range("CZ" & i + 1).Value, _
                                            ExcelSheet.Range("DB" & i + 2).Value, ExcelSheet.Range("DB" & i + 1).Value, ExcelSheet.Range("DD" & i + 2).Value, ExcelSheet.Range("DD" & i + 1).Value, ExcelSheet.Range("DF" & i + 2).Value, ExcelSheet.Range("DF" & i + 1).Value, "dbo.PORev_DetailUpload", templateCode)

                                        ElseIf statusApprove = "1" Then
                                            If ExcelSheet.Range("AG" & i).Value.ToString & "" = "YES" Then
                                                clsTmpDB.insertDetailPOUpload(tmp, _
                                                ExcelSheet.Range("AX" & i + 2).Value, ExcelSheet.Range("AX" & i + 1).Value, ExcelSheet.Range("AZ" & i + 2).Value, ExcelSheet.Range("AZ" & i + 1).Value, ExcelSheet.Range("BB" & i + 2).Value, ExcelSheet.Range("BB" & i + 1).Value, ExcelSheet.Range("BD" & i + 2).Value, ExcelSheet.Range("BD" & i + 1).Value, _
                                                ExcelSheet.Range("BF" & i + 2).Value, ExcelSheet.Range("BF" & i + 1).Value, ExcelSheet.Range("BH" & i + 2).Value, ExcelSheet.Range("BH" & i + 1).Value, ExcelSheet.Range("BJ" & i + 2).Value, ExcelSheet.Range("BJ" & i + 1).Value, ExcelSheet.Range("BL" & i + 2).Value, ExcelSheet.Range("BL" & i + 1).Value, _
                                                ExcelSheet.Range("BN" & i + 2).Value, ExcelSheet.Range("BN" & i + 1).Value, ExcelSheet.Range("BP" & i + 2).Value, ExcelSheet.Range("BP" & i + 1).Value, ExcelSheet.Range("BR" & i + 2).Value, ExcelSheet.Range("BR" & i + 1).Value, ExcelSheet.Range("BT" & i + 2).Value, ExcelSheet.Range("BT" & i + 1).Value, _
                                                ExcelSheet.Range("BV" & i + 2).Value, ExcelSheet.Range("BV" & i + 1).Value, ExcelSheet.Range("BX" & i + 2).Value, ExcelSheet.Range("BX" & i + 1).Value, ExcelSheet.Range("BZ" & i + 2).Value, ExcelSheet.Range("BZ" & i + 1).Value, ExcelSheet.Range("CB" & i + 2).Value, ExcelSheet.Range("CB" & i + 1).Value, _
                                                ExcelSheet.Range("CD" & i + 2).Value, ExcelSheet.Range("CD" & i + 1).Value, ExcelSheet.Range("CF" & i + 2).Value, ExcelSheet.Range("CF" & i + 1).Value, ExcelSheet.Range("CH" & i + 2).Value, ExcelSheet.Range("CH" & i + 1).Value, ExcelSheet.Range("CJ" & i + 2).Value, ExcelSheet.Range("CJ" & i + 1).Value, _
                                                ExcelSheet.Range("CL" & i + 2).Value, ExcelSheet.Range("CL" & i + 1).Value, ExcelSheet.Range("CN" & i + 2).Value, ExcelSheet.Range("CN" & i + 1).Value, ExcelSheet.Range("CP" & i + 2).Value, ExcelSheet.Range("CP" & i + 1).Value, ExcelSheet.Range("CR" & i + 2).Value, ExcelSheet.Range("CR" & i + 1).Value, _
                                                ExcelSheet.Range("CT" & i + 2).Value, ExcelSheet.Range("CT" & i + 1).Value, ExcelSheet.Range("CV" & i + 2).Value, ExcelSheet.Range("CV" & i + 1).Value, ExcelSheet.Range("CX" & i + 2).Value, ExcelSheet.Range("CX" & i + 1).Value, ExcelSheet.Range("CZ" & i + 2).Value, ExcelSheet.Range("CZ" & i + 1).Value, _
                                                ExcelSheet.Range("DB" & i + 2).Value, ExcelSheet.Range("DB" & i + 1).Value, ExcelSheet.Range("DD" & i + 2).Value, ExcelSheet.Range("DD" & i + 1).Value, ExcelSheet.Range("DF" & i + 2).Value, ExcelSheet.Range("DF" & i + 1).Value, "dbo.PORev_DetailUpload", templateCode)
                                            ElseIf ExcelSheet.Range("AG" & i).Value.ToString & "" = "NO" Then
                                                clsTmpDB.insertDetailPOUpload(tmp, _
                                                0, ExcelSheet.Range("AX" & i + 1).Value, 0, ExcelSheet.Range("AZ" & i + 1).Value, 0, ExcelSheet.Range("BB" & i + 1).Value, 0, ExcelSheet.Range("BD" & i + 1).Value, _
                                                0, ExcelSheet.Range("BF" & i + 1).Value, 0, ExcelSheet.Range("BH" & i + 1).Value, 0, ExcelSheet.Range("BJ" & i + 1).Value, 0, ExcelSheet.Range("BL" & i + 1).Value, _
                                                0, ExcelSheet.Range("BN" & i + 1).Value, 0, ExcelSheet.Range("BP" & i + 1).Value, 0, ExcelSheet.Range("BR" & i + 1).Value, 0, ExcelSheet.Range("BT" & i + 1).Value, _
                                                0, ExcelSheet.Range("BV" & i + 1).Value, 0, ExcelSheet.Range("BX" & i + 1).Value, 0, ExcelSheet.Range("BZ" & i + 1).Value, 0, ExcelSheet.Range("CB" & i + 1).Value, _
                                                0, ExcelSheet.Range("CD" & i + 1).Value, 0, ExcelSheet.Range("CF" & i + 1).Value, 0, ExcelSheet.Range("CH" & i + 1).Value, 0, ExcelSheet.Range("CJ" & i + 1).Value, _
                                                0, ExcelSheet.Range("CL" & i + 1).Value, 0, ExcelSheet.Range("CN" & i + 1).Value, 0, ExcelSheet.Range("CP" & i + 1).Value, 0, ExcelSheet.Range("CR" & i + 1).Value, _
                                                0, ExcelSheet.Range("CT" & i + 1).Value, 0, ExcelSheet.Range("CV" & i + 1).Value, 0, ExcelSheet.Range("CX" & i + 1).Value, 0, ExcelSheet.Range("CZ" & i + 1).Value, _
                                                0, ExcelSheet.Range("DB" & i + 1).Value, 0, ExcelSheet.Range("DD" & i + 1).Value, 0, ExcelSheet.Range("DF" & i + 1).Value, "dbo.PORev_DetailUpload", templateCode)
                                            End If
                                        ElseIf statusApprove = "2" Then
                                            clsTmpDB.insertDetailPOUpload(tmp, _
                                               0, ExcelSheet.Range("AX" & i + 1).Value, 0, ExcelSheet.Range("AZ" & i + 1).Value, 0, ExcelSheet.Range("BB" & i + 1).Value, 0, ExcelSheet.Range("BD" & i + 1).Value, _
                                               0, ExcelSheet.Range("BF" & i + 1).Value, 0, ExcelSheet.Range("BH" & i + 1).Value, 0, ExcelSheet.Range("BJ" & i + 1).Value, 0, ExcelSheet.Range("BL" & i + 1).Value, _
                                               0, ExcelSheet.Range("BN" & i + 1).Value, 0, ExcelSheet.Range("BP" & i + 1).Value, 0, ExcelSheet.Range("BR" & i + 1).Value, 0, ExcelSheet.Range("BT" & i + 1).Value, _
                                               0, ExcelSheet.Range("BV" & i + 1).Value, 0, ExcelSheet.Range("BX" & i + 1).Value, 0, ExcelSheet.Range("BZ" & i + 1).Value, 0, ExcelSheet.Range("CB" & i + 1).Value, _
                                               0, ExcelSheet.Range("CD" & i + 1).Value, 0, ExcelSheet.Range("CF" & i + 1).Value, 0, ExcelSheet.Range("CH" & i + 1).Value, 0, ExcelSheet.Range("CJ" & i + 1).Value, _
                                               0, ExcelSheet.Range("CL" & i + 1).Value, 0, ExcelSheet.Range("CN" & i + 1).Value, 0, ExcelSheet.Range("CP" & i + 1).Value, 0, ExcelSheet.Range("CR" & i + 1).Value, _
                                               0, ExcelSheet.Range("CT" & i + 1).Value, 0, ExcelSheet.Range("CV" & i + 1).Value, 0, ExcelSheet.Range("CX" & i + 1).Value, 0, ExcelSheet.Range("CZ" & i + 1).Value, _
                                               0, ExcelSheet.Range("DB" & i + 1).Value, 0, ExcelSheet.Range("DD" & i + 1).Value, 0, ExcelSheet.Range("DF" & i + 1).Value, "dbo.PORev_DetailUpload", templateCode)
                                        End If

                                        i = i + 2

                                    Else
                                        Application.DoEvents()
                                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo, True)

                                        Application.DoEvents()
                                        msgInfo = "read file upload process..."
                                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                                        Me.Cursor = Cursors.Default
                                        Exit For
                                    End If

                                ElseIf templateCode = "POEM" Then 'PO MONTHLY
                                    statusApprove = 0
                                    If Trim(tmp.ForwarderID) <> "SEIWA" Then
                                        tmp.PONo = Trim(tmp.PONo)
                                        tmp.PartNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                        tmp.POQty = Val(ExcelSheet.Range("AE" & i).Value)
                                        tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                        tmp.ETDSplit1 = ExcelSheet.Range("Z" & i).Value & ""
                                        tmp.QtySplit1 = Val(ExcelSheet.Range("AE" & i).Value)

                                        If z = 0 Then
                                            If IsNothing(ExcelSheet.Range("AE" & i).Value) = True And postatusExportExist = False Then GoTo NextProcess
                                            tmp.POQty = Val(ExcelSheet.Range("AE" & i).Value)
                                            tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                            tmp.ETDSplit = ExcelSheet.Range("Z" & i).Value & ""
                                            tmp.QtySplit = Val(ExcelSheet.Range("AE" & i).Value)

                                            postatusExportExist = True : AdaQty = True : jmlSplit = 1
                                        ElseIf z = 1 Then
                                            If IsNothing(ExcelSheet.Range("AO" & i).Value) = True And postatusExportExist = False Then Exit For
                                            If IsNothing(ExcelSheet.Range("AO" & i).Value) = True And postatusExportExist = True Then GoTo NextProcess

                                            tmp.POQty = Val(ExcelSheet.Range("AO" & i).Value)
                                            tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                            tmp.ETDSplit = ExcelSheet.Range("AJ" & i).Value & ""
                                            tmp.QtySplit = Val(ExcelSheet.Range("AO" & i).Value)

                                            If AdaQty = False Then
                                                If postatusExportExist = True And jmlSplit > 0 Then
                                                    tmp.OrderNo = tmp.PONo & "-" & jmlSplit
                                                    AdaQty = True
                                                    jmlSplit = jmlSplit + 1
                                                    inputMaster3 = False
                                                Else
                                                    jmlSplit = 1
                                                    tmp.OrderNo = tmp.PONo
                                                    AdaQty = True
                                                End If
                                            Else
                                                If postatusExportExist = True Then
                                                    If jmlSplit - 1 = 0 Then
                                                        tmp.OrderNo = tmp.PONo
                                                    Else
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit - 1
                                                    End If
                                                    AdaQty = True
                                                Else
                                                    tmp.OrderNo = tmp.PONo
                                                    AdaQty = True
                                                End If
                                            End If

                                        ElseIf z = 2 Then
                                            If IsNothing(ExcelSheet.Range("AY" & i).Value) = True And postatusExportExist = False Then Exit For
                                            If IsNothing(ExcelSheet.Range("AY" & i).Value) = True And postatusExportExist = True Then GoTo NextProcess
                                            tmp.POQty = Val(ExcelSheet.Range("AY" & i).Value)
                                            tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                            tmp.ETDSplit = ExcelSheet.Range("AT" & i).Value & ""
                                            tmp.QtySplit = Val(ExcelSheet.Range("AY" & i).Value)
                                            If AdaQty = False Then
                                                If jmlSplit = z Then jmlSplit = jmlSplit - 1
                                                If postatusExportExist = True And jmlSplit > 0 Then
                                                    tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                    inputMaster3 = False
                                                    AdaQty = True
                                                    jmlSplit = jmlSplit + 1
                                                Else
                                                    jmlSplit = 1
                                                    tmp.OrderNo = tmp.PONo
                                                    AdaQty = True
                                                End If
                                            Else
                                                If postatusExportExist = True Then
                                                    If jmlSplit - 1 = 0 Then
                                                        tmp.OrderNo = tmp.PONo
                                                    Else
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit
                                                    End If
                                                    AdaQty = True
                                                Else
                                                    tmp.OrderNo = tmp.PONo
                                                    AdaQty = True
                                                End If
                                            End If
                                        ElseIf z = 3 Then
                                            If IsNothing(ExcelSheet.Range("BI" & i).Value) = True And postatusExportExist = False Then Exit For
                                            If IsNothing(ExcelSheet.Range("BI" & i).Value) = True And postatusExportExist = True Then GoTo NextProcess
                                            tmp.POQty = Val(ExcelSheet.Range("BI" & i).Value)
                                            tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                            tmp.ETDSplit = ExcelSheet.Range("BD" & i).Value & ""
                                            tmp.QtySplit = Val(ExcelSheet.Range("BI" & i).Value)

                                            If AdaQty = False Then
                                                If jmlSplit = z Then jmlSplit = jmlSplit - 1
                                                If postatusExportExist = True And jmlSplit > 0 Then
                                                    jmlSplit = jmlSplit
                                                    tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                    inputMaster3 = False
                                                    AdaQty = True
                                                    jmlSplit = jmlSplit
                                                Else
                                                    jmlSplit = 1
                                                    tmp.OrderNo = tmp.PONo
                                                    AdaQty = True
                                                End If
                                            Else
                                                If postatusExportExist = True Then
                                                    If jmlSplit - 1 = 0 Then
                                                        tmp.OrderNo = tmp.PONo
                                                    Else
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit
                                                    End If
                                                    AdaQty = True
                                                Else
                                                    tmp.OrderNo = tmp.PONo
                                                    AdaQty = True
                                                End If
                                            End If
                                        ElseIf z = 4 Then
                                            If IsNothing(ExcelSheet.Range("BS" & i).Value) = True And postatusExportExist = False Then Exit For
                                            If IsNothing(ExcelSheet.Range("BS" & i).Value) = True And postatusExportExist = True Then GoTo NextProcess

                                            tmp.POQty = Val(ExcelSheet.Range("BS" & i).Value)
                                            tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                            tmp.ETDSplit = ExcelSheet.Range("BN" & i).Value & ""
                                            tmp.QtySplit = Val(ExcelSheet.Range("BS" & i).Value)
                                            If AdaQty = False Then
                                                If jmlSplit = z Then jmlSplit = jmlSplit - 1
                                                If postatusExportExist = True And jmlSplit > 0 Then
                                                    jmlSplit = jmlSplit
                                                    tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                    inputMaster3 = True
                                                    AdaQty = True
                                                    jmlSplit = jmlSplit
                                                Else
                                                    jmlSplit = 1
                                                    tmp.OrderNo = tmp.PONo
                                                    AdaQty = True
                                                End If
                                            Else
                                                If postatusExportExist = True Then
                                                    If jmlSplit - 1 = 0 Then
                                                        tmp.OrderNo = tmp.PONo
                                                    Else
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit
                                                    End If
                                                    AdaQty = True
                                                Else
                                                    tmp.OrderNo = tmp.PONo
                                                    AdaQty = True
                                                End If
                                            End If
                                        End If

                                        tmp.Forecast1 = Val(ExcelSheet.Range("BX" & i).Value)
                                        tmp.Forecast2 = Val(ExcelSheet.Range("CB" & i).Value)
                                        tmp.Forecast3 = Val(ExcelSheet.Range("CF" & i).Value)

                                        If (postatusExportExist = True And tmp.POQty <> 0 And jmlSplit > 0) Or (postatusExportExist = False) Or (postatusExportExist = True And jmlSplit = 0) Then
                                            autoApprove = clsTmpDB.CekPOMonthlyAutoApprover("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, statusApprove)
                                            poExportStatus = clsTmpDB.CekPOMonthlyStatus("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, statusApprove)
                                            ii = 0
                                            iSplit = z + 1
                                            OrderSplit = ""

                                            If autoApprove = 0 And poExportStatus = 0 Then
                                                If inputMaster3 = False Then
                                                    clsTmpDB.insertMasterPOMonthlyUpload(tmp, SQLCom, "dbo.PO_MasterUpload_Export", templateCode)
                                                    clsTmpDB.UpdatePOMonthlyMaster(tmp, SQLCom, statusApprove, "dbo.PO_Master_Export")
                                                    'clsTmpDB.UpdatePOMonthlyMasterUpload(tmp, SQLCom, "dbo.PO_MasterUpload_Export", templateCode)
                                                    'upload ke PO_master jika belum ada
                                                    Status = 0
                                                    Status = clsTmpDB.CekPOMonthlyEXIST("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, statusApprove)
                                                    If Status = 0 Then
                                                        clsTmpDB.insertMasterPOMonthlyAfterUpload(tmp, SQLCom, "dbo.PO_Master_Export", templateCode, iSplit)
                                                        clsTmpDB.UpdatePOMonthlyMaster(tmp, SQLCom, statusApprove, "dbo.PO_Master_Export")
                                                    End If
                                                    inputMaster3 = True
                                                End If

                                                If statusApprove = "0" Then
                                                    clsTmpDB.insertDetailPOMonthlyUpload(tmp, _
                                                    ExcelSheet.Range("I9" & i).Value, ExcelSheet.Range("I16" & i).Value, ExcelSheet.Range("I11" & i).Value, ExcelSheet.Range("AE13" & i).Value, ExcelSheet.Range("D41" & i).Value, "", "", "", _
                                                    "", "", "", "", "", "", "", ExcelSheet.Range("AC41" & i).Value, ExcelSheet.Range("AC41" & i).Value, "", ExcelSheet.Range("AG41" & i).Value, ExcelSheet.Range("AK41" & i).Value, ExcelSheet.Range("AO41" & i).Value, _
                                                    "dbo.PO_DetailUpload_Export", templateCode)
                                                    'upload ke PO_master jika belum ada
                                                    clsTmpDB.insertDetailPOMonthlyAfterUpload(tmp, _
                                                    ExcelSheet.Range("I9" & i).Value, ExcelSheet.Range("I16" & i).Value, ExcelSheet.Range("I11" & i).Value, ExcelSheet.Range("AE13" & i).Value, ExcelSheet.Range("D41" & i).Value, "", "", "", _
                                                    "", "", "", "", "", "", "", ExcelSheet.Range("AC41" & i).Value, ExcelSheet.Range("AC41" & i).Value, "", ExcelSheet.Range("AG41" & i).Value, ExcelSheet.Range("AK41" & i).Value, ExcelSheet.Range("AO41" & i).Value, _
                                                    "dbo.PO_Detail_Export", templateCode)
                                                    postatusExportExist = True
                                                End If
                                            Else
                                                Application.DoEvents()
                                                gridProcess(rtbProcess, 1, 0, "End " & msgInfo, True)

                                                Application.DoEvents()
                                                msgInfo = "read file upload process..."
                                                gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                                                Me.Cursor = Cursors.Default
                                                Exit For

                                            End If
                                        End If
                                    End If
                                ElseIf templateCode = "POEE" Then 'PO EMERGENCY
                                    statusApprove = 0

                                    tmp.PartNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                    tmp.POQty = Val(ExcelSheet.Range("AD" & i).Value)
                                    tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)

                                    autoApprove = clsTmpDB.CekPOMonthlyAutoApprover("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, statusApprove)
                                    If autoApprove = 0 Then
                                        If inputMaster3 = False Then
                                            clsTmpDB.insertMasterPOMonthlyUpload(tmp, SQLCom, "dbo.PO_MasterUpload_Export", templateCode)
                                            clsTmpDB.UpdatePOMonthlyMaster(tmp, SQLCom, statusApprove, "dbo.PO_Master_Export")
                                            clsTmpDB.UpdatePOMonthlyMasterUpload(tmp, SQLCom, "dbo.PO_MasterUpload_Export", templateCode)
                                            inputMaster3 = True
                                        End If

                                        If statusApprove = "0" Then
                                            clsTmpDB.insertDetailPOMonthlyUpload(tmp, _
                                            ExcelSheet.Range("I9" & i).Value, ExcelSheet.Range("I16" & i).Value, ExcelSheet.Range("I11" & i).Value, ExcelSheet.Range("AE13" & i).Value, ExcelSheet.Range("D41" & i).Value, "", "", "", _
                                            "", "", "", "", "", "", "", ExcelSheet.Range("W41" & i).Value, ExcelSheet.Range("AG41" & i).Value, "", "", "", "", _
                                            "dbo.PO_DetailUpload_Export", templateCode)
                                        End If

                                        'i = i + 1

                                    Else
                                        Application.DoEvents()
                                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo, True)

                                        Application.DoEvents()
                                        msgInfo = "read file upload process..."
                                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                                        Me.Cursor = Cursors.Default
                                        Exit For

                                    End If

                                ElseIf templateCode = "INV" Then 'INVOICE

                                    tmp.SuratJalanNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                    tmp.PONo = ExcelSheet.Range("J" & i).Value.ToString & ""
                                    tmp.KanbanNo = ExcelSheet.Range("Q" & i).Value.ToString & ""
                                    tmp.PartNo = ExcelSheet.Range("U" & i).Value.ToString & ""
                                    tmp.InvoiceQty = ExcelSheet.Range("AU" & i).Value
                                    If ExcelSheet.Range("AX" & i).Value Is Nothing Then
                                        tmp.InvCurrCls = "03"
                                    Else
                                        tmp.InvCurrCls = clsTmpDB.CurrCls(ExcelSheet.Range("AX" & i).Value.ToString & "")
                                    End If

                                    tmp.InvPrice = ExcelSheet.Range("AZ" & i).Value
                                    tmp.InvAmount = ExcelSheet.Range("BD" & i).Value
                                    ds = clsTmpDB.StatusDelivery(tmp.PONo)

                                    If ds Is Nothing Then
                                    Else
                                        tmp.ReceiveCurrCls = tmp.InvCurrCls 'clsTmpDB.CurrCls(ds.Tables(0).Rows(0)("DeliveryByPASICls"))
                                        tmp.ReceivePrice = clsTmpDB.GetPrice(tmp.AffiliateID, tmp.PartNo, tmp.ReceiveCurrCls)
                                        tmp.ReceiveAmount = ds.Tables(0).Rows(0)("AmountT")

                                        If ds.Tables(0).Rows(0)("DeliveryByPASICls") = "0" Then 'Affiliate
                                            ds1 = clsTmpDB.GetRecValue(tmp.PONo, "ReceiveAffiliate_Detail", tmp.SuratJalanNo, tmp.SupplierID _
                                                                        , tmp.AffiliateID, tmp.KanbanNo, tmp.PartNo)
                                            tmp.ReceiveQty = ds1.Tables(0).Rows(0)("ReqQty")
                                        Else ' PASI
                                            ds1 = clsTmpDB.GetRecValue(tmp.PONo, "ReceivePASI_Detail", tmp.SuratJalanNo, tmp.SupplierID _
                                                                        , tmp.AffiliateID, tmp.KanbanNo, tmp.PartNo)
                                            tmp.ReceiveQty = ds1.Tables(0).Rows(0)("GoodRecQty")
                                        End If
                                    End If

                                    If inputMaster5 = False Then
                                        clsTmpDB.insertMasterInvoice(tmp, SQLCom)
                                        inputMaster5 = True
                                    End If
                                    clsTmpDB.insertDetailInvoice(tmp)
                                ElseIf templateCode = "DO-EX" Then 'Dian Export
                                    If Trim(tmp.ForwarderID) <> "SEIWA" Then
                                        If ExcelSheet.Range("B" & i).Value.ToString <> "E" Then
                                            tmp.PartNo = Trim(ExcelSheet.Range("I" & i).Value.ToString) & ""
                                            tmp.PONo = Trim(ExcelSheet.Range("AE" & 13).Value.ToString) & ""

                                            If ExcelSheet.Range("AE" & 15).Value IsNot Nothing Then
                                                tmp.OrderNo = Trim(ExcelSheet.Range("AE" & 15).Value.ToString) & ""
                                            Else
                                                tmp.OrderNo = ""
                                            End If
                                            If tmp.OrderNo = "" Then
                                                tmp.OrderNo = tmp.PONo
                                            End If

                                            tmp.UnitCls = clsTmpDB.UnitCls(ExcelSheet.Range("AA" & i).Value.ToString & "")
                                            tmp.DOQty = IIf(IsNumeric(ExcelSheet.Range("AM" & i).Value) = False, 0, ExcelSheet.Range("AM" & i).Value)
                                            If tmp.DOQty > 0 Then
                                                If inputMaster2 = False Then
                                                    clsTmpDB.insertMasterDOEX(tmp, SQLCom)
                                                    inputMaster2 = True
                                                End If

                                                'looping Box No
                                                Dim i_loopBox As Integer = 0
                                                Dim startBox As Integer = 0
                                                Dim ls_boxno As String = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 8)

                                                If Microsoft.VisualBasic.Mid(Trim(ExcelSheet.Range("W" & i).Value.ToString), 2, 1) <> "0" Then
                                                    ls_boxno = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 2)
                                                Else
                                                    ls_boxno = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 1)
                                                End If

                                                tmp.m_BoxNo = Trim(ExcelSheet.Range("W" & i).Value.ToString) & ""
                                                Boxdetail = clsTmpDB.CekDOBoxDetail("dbo.DOSupplier_DetailBox_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, tmp.PartNo)
                                                If Trim(Boxdetail) = 0 Then
                                                    Boxdetail = Trim(Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 8))
                                                Else
                                                    Boxdetail = ls_boxno + Microsoft.VisualBasic.Right("000000" & Trim(Boxdetail), 6)
                                                End If
                                                tmp.m_JmlBox = CDbl(ExcelSheet.Range("AM" & i).Value.ToString) / CDbl(ExcelSheet.Range("AC" & i).Value.ToString)
                                                '(Microsoft.VisualBasic.Right(Trim(ExcelSheet.Range("W" & i).Value.ToString), 5) - Microsoft.VisualBasic.Mid(Trim(ExcelSheet.Range("W" & i).Value.ToString), 2, 5)) + 1

                                                clsTmpDB.insertDetailDOEX(tmp)

                                                If IsNewRecEx = True Then
                                                    For i_loopBox = 0 To tmp.m_JmlBox - 1
                                                        clsTmpDB.insertDetailDOBoxEX(tmp, ls_boxno + Microsoft.VisualBasic.Right("000000" & (Microsoft.VisualBasic.Right(Boxdetail, 6) + i_loopBox), 6))
                                                    Next
                                                End If
                                                'looping Box No
                                            End If
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf templateCode = "REC-EX" Then 'Dian Export

                                    tmp.PartNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                    tmp.PONo = tmp.PONo
                                    tmp.OrderNo = tmp.PONo

                                    If IsNewRecEx = False Then
                                        tmp.GoodRecQty = IIf(IsNumeric(ExcelSheet.Range("AH" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AH" & i).Value) * CDbl(ExcelSheet.Range("X" & i).Value))
                                        tmp.DefectRecQty = IIf(IsNumeric(ExcelSheet.Range("AL" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AL" & i).Value) * CDbl(ExcelSheet.Range("X" & i).Value))
                                        tmp.UnitCls = clsTmpDB.UnitCls(ExcelSheet.Range("V" & i).Value.ToString & "")
                                    Else
                                        If recStatus = True Then
                                            clsTmpDB.DeleteRecEX(tmp, SQLCom)
                                        End If
                                        tmp.GoodRecQty = IIf(IsNumeric(ExcelSheet.Range("AL" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AL" & i).Value) * CDbl(ExcelSheet.Range("AB" & i).Value))
                                        tmp.DefectRecQty = IIf(IsNumeric(ExcelSheet.Range("AP" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AP" & i).Value) * CDbl(ExcelSheet.Range("AB" & i).Value))
                                        tmp.UnitCls = clsTmpDB.UnitCls(ExcelSheet.Range("Z" & i).Value.ToString & "")
                                    End If
                                    If tmp.GoodRecQty > 0 Or tmp.DefectRecQty > 0 Then
                                        Dim Xstatus As String = ""
                                        Dim totBox As Long = 0
                                        If tmp.GoodRecQty > 0 Then
                                            Xstatus = "G"
                                            totBox = IIf(IsNumeric(ExcelSheet.Range("AL" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AL" & i).Value))
                                        Else
                                            Xstatus = "D"
                                            totBox = IIf(IsNumeric(ExcelSheet.Range("AP" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AP" & i).Value))
                                        End If

                                        If inputMaster2 = False Then
                                            clsTmpDB.insertMasterRecevingEX(tmp, SQLCom)
                                            inputMaster2 = True
                                        End If
                                        clsTmpDB.insertDetailReceivingEX(tmp, Xstatus)

                                        If IsNewRecEx = True Then
                                            Dim i_rec1 As Integer = Microsoft.VisualBasic.Right(Trim(ExcelSheet.Range("R" & i).Value), 6)
                                            Dim i_rec2 As Integer = Microsoft.VisualBasic.Right(Trim(ExcelSheet.Range("V" & i).Value), 6)
                                            Dim i_LabelNo As String = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("V" & i).Value), 2)
                                            Dim i_LabelNo2 As String = ""
                                            Dim L1 As String = i_LabelNo & Microsoft.VisualBasic.Right(("000000" & i_rec1), 6)
                                            Dim L2 As String = i_LabelNo & Microsoft.VisualBasic.Right(("000000" & i_rec2), 6)


                                            clsTmpDB.insertDetailReceivingEX_BOX(tmp, L1, L2, Xstatus, totBox)
                                            clsTmpDB.RemainingReceiveExport(tmp)

                                            For i_rec1 = i_rec1 To i_rec2
                                                i_LabelNo2 = i_LabelNo & Microsoft.VisualBasic.Right(("000000" & i_rec1), 6)
                                                clsTmpDB.UpdateLabelPrint_RecEX(tmp, i_LabelNo2, Xstatus)
                                            Next
                                        End If
                                    End If
                                ElseIf templateCode = "INV-EX" Then 'DIAN EXPORT

                                    tmp.SuratJalanNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                    tmp.PONo = ExcelSheet.Range("J" & i).Value.ToString & ""
                                    tmp.OrderNo = ExcelSheet.Range("J" & i).Value.ToString & ""
                                    tmp.PartNo = ExcelSheet.Range("Q" & i).Value.ToString & ""
                                    tmp.InvoiceQty = ExcelSheet.Range("AM" & i).Value

                                    If ExcelSheet.Range("AX" & i).Value Is Nothing Then
                                        tmp.InvCurrCls = "03"
                                    Else
                                        tmp.InvCurrCls = clsTmpDB.CurrCls(ExcelSheet.Range("AX" & i).Value.ToString & "")
                                    End If

                                    tmp.InvPrice = ExcelSheet.Range("AR" & i).Value
                                    tmp.InvAmount = ExcelSheet.Range("AV" & i).Value
                                    ds = clsTmpDB.StatusDelivery(tmp.PONo)

                                    If inputMaster5 = False Then
                                        clsTmpDB.insertMasterInvoiceEx(tmp, SQLCom)
                                        inputMaster5 = True
                                    End If
                                    clsTmpDB.insertDetailInvoiceEx(tmp)
                                ElseIf templateCode = "TALLY" Then 'DIAN Export
                                    tmp.PalletNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                    tmp.OrderNo = ExcelSheet.Range("I" & i).Value.ToString & ""
                                    tmp.PartNo = ExcelSheet.Range("O" & i).Value.ToString & ""
                                    tmp.BoxNo = ExcelSheet.Range("AE" & i).Value
                                    tmp.Length = ExcelSheet.Range("AK" & i).Value
                                    tmp.Width = ExcelSheet.Range("AN" & i).Value
                                    tmp.Height = ExcelSheet.Range("AQ" & i).Value
                                    tmp.M3 = ExcelSheet.Range("AT" & i).Value
                                    tmp.WeightPallet = ExcelSheet.Range("AW" & i).Value

                                    clsTmpDB.insertMasterTaily(tmp)
                                    clsTmpDB.insertDetailTaily(tmp)
                                End If
NextProcess:
                            Next
                        Next
                        'SEND NOTIFICATION
                        Application.DoEvents()
                        msgInfo = "send email process " & tmpAllNo
                        gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

                        up_SendEmail(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                                     tmp.PONo, tmp.Period, tmp.AffiliateName, tmp.PORevNo, tmp.ShipCls, _
                                     tmp.DeliveryLocation, tmp.KanbanDate, tmp.Remarks, tmp.CommercialCls, _
                                     tmp.InvoiceNo, tmp.PartNo, tmp.KanbanNo, tmp.SuratJalanNo)

                        Application.DoEvents()
                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()

                        SQLTrans.Commit()

                        Application.DoEvents()
                        msgInfo = "move file upload process " & tmpAllNo
                        gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

                        'MOVE FILE
                        up_XMLFile_Copy(txtpath.Text & "\", txtPathBackup.Text & "\", excelName)

                        Application.DoEvents()
                        gridProcess(rtbProcess, 1, 0, "1 move file upload", True)
                        gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                        Me.Cursor = Cursors.Default

                        inputMaster = False : inputMaster2 = False : inputMaster3 = False
                        inputMaster4 = False : inputMaster5 = False

                    End Using
                Catch ex As Exception

                    Application.DoEvents()
                    If msgInfo <> "" Then
                        gridProcess(rtbProcess, 2, 0, "Failed " & msgInfo & vbCrLf, True)
                    Else
                        gridProcess(rtbProcess, 2, 0, "Failed " & ex.Message.ToString & vbCrLf, True)
                    End If

                    Me.Cursor = Cursors.Default
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If
                    'xlApp.Quit()
                    inputMaster = False : inputMaster2 = False : inputMaster3 = False
                    inputMaster4 = False

                    Application.DoEvents()
                    msgInfo = "move data corrupt process " & tmpAllNo
                    gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

                    'MOVE FILE
                    up_XMLFile_Copy(txtpath.Text & "\", tmpPathError & "\", excelName)

                End Try

                Application.DoEvents()
                msgInfo = "read file upload process..."
                gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

                'Next

            Catch ex As Exception

                Application.DoEvents()
                If msgInfo <> "" Then
                    gridProcess(rtbProcess, 2, 0, "Failed " & msgInfo & vbCrLf, True)
                Else
                    gridProcess(rtbProcess, 2, 0, "Failed " & ex.Message.ToString & vbCrLf, True)
                End If

                If Not IsNothing(ExcelBook) Then
                    'ExcelBook.Save()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                End If

                Application.DoEvents()
                msgInfo = "move file corrupt process " & tmpAllNo
                gridProcess(rtbProcess, 1, 0, "Start " & msgInfo, True)

                'MOVE FILE
                up_XMLFile_Copy(txtpath.Text & "\", tmpPathError & "\", excelName)

                Application.DoEvents()
                gridProcess(rtbProcess, 1, 0, "1 move file corrupt", True)
                gridProcess(rtbProcess, 1, 0, "End " & msgInfo & vbCrLf, True)

            Finally
                Me.Cursor = Cursors.Default

            End Try
        Next

        If Not IsNothing(ExcelBook) Then
            Do While Marshal.ReleaseComObject(xlApp) > 0
            Loop
            If Not IsNothing(ExcelBook) Then
                Do While Marshal.ReleaseComObject(xlApp) > 0
                Loop
            End If
            xlApp = Nothing
            ExcelBook = Nothing
            GC.GetTotalMemory(False)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.GetTotalMemory(True)
        End If

        Exit Sub
Err_Handler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    End Sub

    Private Sub up_XMLFile_Copy(ByVal pPathSource As String, ByVal pPathDestination As String, ByVal excelName As String)
        Try
            Dim di As New IO.DirectoryInfo(pPathSource)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.xlsm")
            Dim fi As IO.FileInfo = Nothing

            If Not System.IO.Directory.Exists(pPathSource) Then
                System.IO.Directory.CreateDirectory(pPathSource)
            End If

            'For Each fi In aryFi
            'My.Computer.FileSystem.MoveFile(pPathSource & fi.Name, pPathDestination & fi.Name, True)
            My.Computer.FileSystem.MoveFile(pPathSource & excelName, pPathDestination & excelName, True)
            'Next

        Catch ex As Exception
            txtMsg.ForeColor = Color.Red
            txtMsg.Text = "Error move file"
        End Try
    End Sub

    Private Function up_SendEmail(ByVal pConstr As String, ByVal pAffiliateID As String, ByVal pSupplierID As String, _
                                  ByVal pTemplateCode As String, ByVal pPONo As String, ByVal pPeriod As String, _
                                  ByVal pAffiliateName As String, ByVal pPONoRev As String, ByVal pShip As String, _
                                  ByVal pDeliveryLoc As String, ByVal pKanbanDate As Date, ByVal pRemarks As String, _
                                  ByVal pCommercialCls As String, ByVal pInvoiceNo As String, ByVal pPartNo As String, _
                                  ByVal pKanbanNo As String, ByVal pSuratJalanNo As String)

        Dim i As Integer, kirimKe As Integer = 0
        Dim pURLPASI As String = "", pURLAffiliate As String = "", pURLSupplier As String = ""
        Dim kirimEmail As Integer = 0
        Dim pPASI As Boolean = False, pAffiliate As Boolean = False, pSupplier As Boolean = False
        Dim EmailSetting As clsEmail, EmailPASI As clsEmail, EmailAffiliate As clsEmail, EmailSupplier As clsEmail, SubjectBody As clsEmail
        Dim ReceipientEmail As String, ReceipientArray() As String
        Dim CCEmail As String, CCArray() As String, retMessage As String = ""
        Dim pReceipient As String = "", pCC As String = ""
        Dim ReceipientList As New List(Of String)
        Dim CCList As New List(Of String)

        Try
            Me.Cursor = Cursors.WaitCursor

            EmailSetting = clsEmailDB.GetEmailSetting(pConstr)
            If EmailSetting Is Nothing Then
                Return ""
            End If
            EmailPASI = clsEmailDB.GetEmailPASI(pConstr, pTemplateCode)
            If EmailPASI Is Nothing Then
                Return ""
            End If
            EmailAffiliate = clsEmailDB.GetEmailAffiliate(pConstr, pAffiliateID, pTemplateCode)
            If EmailAffiliate Is Nothing Then
                Return ""
            End If
            EmailSupplier = clsEmailDB.GetEmailSupplier(pConstr, pSupplierID, pTemplateCode)
            If EmailSupplier Is Nothing Then
                Return ""
            End If

            If templateCode = "PO" Then
                '=========================== EMAIL KE PASI ===========================
                'URL AFFILIATE PO APPROVAL DETAIL (PASI SYSTEM)
                pURLPASI = _
                        "http://" & clsNotification.pub_ServerNamePasi & "/AffiliateOrder/AffiliateOrderAppDetail.aspx?id2=" & _
                        clsNotification.EncryptURL(pPONo.Trim) & _
                        "&t1=" & clsNotification.EncryptURL(pAffiliateID) & _
                        "&t2=" & clsNotification.EncryptURL(pAffiliateName.Trim) & _
                        "&t3=" & clsNotification.EncryptURL(pPeriod) & _
                        "&t4=" & clsNotification.EncryptURL(pSupplierID.Trim) & _
                        "&t5=" & clsNotification.EncryptURL(pRemarks.Trim) & _
                        "&t6=" & clsNotification.EncryptURL("1") & _
                        "&t7=" & clsNotification.EncryptURL("1") & _
                        "&t8=" & clsNotification.EncryptURL(pShip) & _
                        "&t9=" & clsNotification.EncryptURL(pCommercialCls.Trim) & _
                        "&t10=" & clsNotification.EncryptURL("") & _
                        "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderAppList.aspx")

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pPONo)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailPASI.EmailPASITo
                'ReceipientEmail = "edi@tos.co.id;dian@tos.co.id"
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailPASI.EmailPASICC
                'CCEmail = "kristriyana@tos.co.id"
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next
                'EmailPASI.EmailPASITo = "hadi@tos.co.id"
                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

                '=====================================================================
                '========================= EMAIL KE SUPPLIER =========================
                'TIDAK PAKAI URL UNTUK SUPPLIER
                pURLPASI = ""

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pPONo)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailSupplier.EmailSupplierTo
                'ReceipientEmail = "edi@tos.co.id;dian@tos.co.id"
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailSupplier.EmailSupplierCC
                'CCEmail = "kristriyana@tos.co.id"
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next
                'EmailPASI.EmailPASITo = "hadi@tos.co.id"
                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

            ElseIf templateCode = "POR" Then
                '=========================== EMAIL KE PASI ===========================
                'URL AFFILIATE PO REVISION DETAIL (PASI SYSTEM)
                pURLPASI = _
                            "http://" & clsNotification.pub_ServerNamePasi & "/AffiliateRevision/AffiliateOrderRevAppDetail.aspx?id2=" & _
                            clsNotification.EncryptURL(pPONo.Trim) & _
                            "&t1=" & clsNotification.EncryptURL(pPeriod) & _
                            "&t2=" & clsNotification.EncryptURL(pPONoRev.Trim) & _
                            "&t3=" & clsNotification.EncryptURL(pPONo.Trim) & _
                            "&t4=" & clsNotification.EncryptURL("") & _
                            "&t5=" & clsNotification.EncryptURL(pAffiliateID.Trim) & _
                            "&t6=" & clsNotification.EncryptURL("") & _
                            "&t7=" & clsNotification.EncryptURL("") & _
                            "&t8=" & clsNotification.EncryptURL("") & _
                            "&t9=" & clsNotification.EncryptURL("2") & _
                            "&t10=" & clsNotification.EncryptURL(pRemarks.Trim) & _
                            "&Session=" & clsNotification.EncryptURL("~/AffiliateRevision/AffiliateOrderRevAppList.aspx")

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pPONoRev)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailPASI.EmailPASITo '& ";" & EmailPASI.EmailPASICC
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailPASI.EmailPASICC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next

                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

                '=====================================================================
                '========================= EMAIL KE SUPPLIER =========================
                'TIDAK PAKAI URL UNTUK SUPPLIER
                pURLPASI = ""

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pPONoRev)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailSupplier.EmailSupplierTo  '& ";" & EmailPASI.EmailPASICC
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailSupplier.EmailSupplierCC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next

                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

            ElseIf templateCode = "KB" Then
                pKanbanDate = Format(pKanbanDate, "dd MMM yyyy")
                '=========================== EMAIL KE AFFILIATE ===========================
                'URL KANBAN ENTRY (AFFILIATE SYSTEM)
                pURLPASI = _
                        "http://" & clsNotification.pub_ServerNameAffiliate & "/Kanban/KanbanCreate.aspx?id2=URL" & _
                        "&t0=" & clsNotification.EncryptURL(pKanbanDate) & _
                        "&t1=" & clsNotification.EncryptURL(pSupplierID) & _
                        "&t2=" & clsNotification.EncryptURL(pDeliveryLoc) & _
                        "&Session=" & clsNotification.EncryptURL("~/Kanban/KanbanList.aspx")

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pKanbanNo)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailAffiliate.EmailAffiliateTo
                'ReceipientEmail = "edi@tos.co.id"
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailAffiliate.EmailAffiliateTo & ";" & EmailAffiliate.EmailAffiliateCC
                'CCEmail = "kristriyana@tos.co.id;dian@tos.co.id"
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next
                'EmailPASI.EmailPASITo = "hadi@tos.co.id"
                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()
                '=====================================================================
                '=========================== EMAIL KE PASI ===========================
                'AFFILIATE KANBAN DETAIL (PASI SYSTEM)
                pURLPASI = _
                        "http://" & clsNotification.pub_ServerNamePasi & "/AffKanban/AffKanbanCreate.aspx?id2=URL" & _
                        "&t0=" & clsNotification.EncryptURL(pKanbanDate) & _
                        "&t1=" & clsNotification.EncryptURL(pSupplierID) & _
                        "&t2=" & clsNotification.EncryptURL(pDeliveryLoc.Trim) & _
                        "&Session=" & clsNotification.EncryptURL("~/AffKanban/AffKanbanList.aspx")

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pKanbanNo)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailPASI.EmailPASITo
                'ReceipientEmail = "edi@tos.co.id"
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailPASI.EmailPASICC
                'CCEmail = "kristriyana@tos.co.id;dian@tos.co.id"
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next
                'EmailPASI.EmailPASITo = "hadi@tos.co.id"
                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

                '=========================================================================
                '=========================== EMAIL KE SUPPLIER ===========================
                'TIDAK PAKAI URL UNTUK SUPPLIER
                pURLPASI = ""

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pKanbanNo)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailSupplier.EmailSupplierTo  '& ";" & EmailPASI.EmailPASICC
                'ReceipientEmail = "edi@tos.co.id;dian@tos.co.id"
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailSupplier.EmailSupplierCC
                'CCEmail = "kristriyana@tos.co.id"
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next
                'EmailPASI.EmailPASITo = "hadi@tos.co.id"
                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

            ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then
                '=========================== EMAIL TO ALL ===========================
                'DO TIDAK PAKAI URL HANYA NOTIFIKASI
                pURLPASI = ""

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pSuratJalanNo)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailSupplier.EmailSupplierTo
                'ReceipientEmail = "edi@tos.co.id;dian@tos.co.id"
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailSupplier.EmailSupplierCC & ";" & EmailPASI.EmailPASICC & ";" & EmailAffiliate.EmailAffiliateTo & ";" & EmailAffiliate.EmailAffiliateCC
                'CCEmail = "kristriyana@tos.co.id"
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next
                'EmailPASI.EmailPASITo = "hadi@tos.co.id"
                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

            ElseIf templateCode = "INV" Then
                '=========================== EMAIL KE PASI ===========================
                'URL INVOICE FROM SUPPLIER DETAIL (PASI SYSTEM)
                pURLPASI = ""

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pInvoiceNo)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailPASI.EmailPASITo '& ";" & EmailPASI.EmailPASICC
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailPASI.EmailPASICC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next

                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

                '=====================================================================
                '========================= EMAIL KE SUPPLIER =========================
                'TIDAK PAKAI URL UNTUK SUPPLIER
                pURLPASI = ""

                SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pInvoiceNo)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailSupplier.EmailSupplierTo  '& ";" & EmailPASI.EmailPASICC
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    ReceipientList.Add(ReceipientArray(i))
                Next
                CCEmail = EmailSupplier.EmailSupplierCC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    CCList.Add(CCArray(i))
                Next

                retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

                ReceipientList.Clear()
                CCList.Clear()

            End If

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            txtMsg.Text = "Error send Email"
            Me.Cursor = Cursors.Default
        End Try


        ''================================= 
        'Dim i As Integer
        'Try
        '    Me.Cursor = Cursors.WaitCursor
        '    Dim EmailSetting As clsEmail = clsEmailDB.GetEmailSetting(pConstr)
        '    If EmailSetting Is Nothing Then
        '        Return ""
        '    End If
        '    Dim EmailPASI As clsEmail = clsEmailDB.GetEmailPASI(pConstr, pTemplateCode)
        '    If EmailPASI Is Nothing Then
        '        Return ""
        '    End If
        '    Dim EmailAffiliate As clsEmail = clsEmailDB.GetEmailAffiliate(pConstr, pAffiliateID, pTemplateCode)
        '    If EmailAffiliate Is Nothing Then
        '        Return ""
        '    End If
        '    Dim EmailSupplier As clsEmail = clsEmailDB.GetEmailSupplier(pConstr, pSupplierID, pTemplateCode)
        '    If EmailSupplier Is Nothing Then
        '        Return ""
        '    End If
        '    Dim SubjectBody As clsEmail = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, "")
        '    If SubjectBody Is Nothing Then
        '        Return ""
        '    End If

        '    Dim ReceipientEmail As String = EmailSupplier.EmailSupplierTo '"edi@tos.co.id" '
        '    Dim ReceipientArray() As String = Split(ReceipientEmail, ";")
        '    Dim ReceipientList As New List(Of String)
        '    For i = 0 To UBound(ReceipientArray)
        '        ReceipientList.Add(ReceipientArray(i))
        '    Next
        '    Dim CCEmail As String = EmailAffiliate.EmailAffiliateTo & ";" & EmailAffiliate.EmailAffiliateCC & ";" & _
        '                            EmailPASI.EmailPASICC & ";" & EmailSupplier.EmailSupplierCC

        '    'CCEmail = "herfin@tos.co.id;hadi@tos.co.id"
        '    Dim CCArray() As String = Split(CCEmail, ";")
        '    Dim CCList As New List(Of String)
        '    For i = 0 To UBound(CCArray)
        '        CCList.Add(CCArray(i))
        '    Next

        '    Dim retMessage As String = ""

        '    retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
        '             SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.SMTPServer, EmailSetting.Port)

        '    Return retMessage

        '    Me.Cursor = Cursors.Default
        'Catch ex As Exception
        '    txtMsg.Text = ex.Message.ToString
        '    Me.Cursor = Cursors.Default
        'End Try
    End Function

    Function SendEmail(ByVal Recipients As List(Of String), _
                        ByVal CC As List(Of String), _
                        ByVal FromAddress As String, _
                        ByVal Subject As String, _
                        ByVal Body As String, _
                        ByVal UserName As String, _
                        ByVal Password As String, _
                        ByVal EnableSSL As Boolean, _
                        ByVal Server As String, _
                        ByVal Port As Integer) As String

        Dim Email As New MailMessage()
        Dim retMsg As String

        Try
            Dim SMTPServer As New SmtpClient

            Email.From = New MailAddress(FromAddress) '("hadi@tos.co.id") '
            For Each Recipient As String In Recipients
                Email.To.Add(Recipient)
            Next
            For Each CC1 As String In CC
                If CC1 <> "" Then
                    Email.CC.Add(CC1)
                    'Email.To.Add(CC1)
                End If
            Next
            Email.Subject = Subject
            Email.Body = Body
            SMTPServer.Host = Server '"tos-is.com" '
            SMTPServer.Port = Port
            SMTPServer.Credentials = New System.Net.NetworkCredential(UserName, Password)
            SMTPServer.EnableSsl = EnableSSL
            SMTPServer.Send(Email)
            Email.Dispose()
            Return ""

        Catch ex As SmtpException
            Email.Dispose()
            retMsg = "Sending Email Failed. Smtp Error."
        Catch ex As ArgumentOutOfRangeException
            Email.Dispose()
            retMsg = "Sending Email Failed. Check Port Number."
        Catch Ex As InvalidOperationException
            Email.Dispose()
            retMsg = "Sending Email Failed. Check Port Number."
        End Try
        Return retMsg

    End Function

    Private Sub gridProcess(ByVal rtbox As RichTextBox, ByVal pColor As Integer, ByVal pPos As Integer, ByVal pMsg As String, Optional ByVal UseTime As Boolean = False)
        With rtbox
            Select Case pColor
                Case 1 : .SelectionColor = Color.Black   ' Hitam
                Case 2 : .SelectionColor = Color.Red     ' Merah
                Case 3 : .SelectionColor = Color.Green   ' Hijau
                Case 4 : .SelectionColor = Color.Blue    ' Biru
                Case 5 : .SelectionColor = Color.Gray    ' Abu
                Case 6 : .SelectionColor = Color.Yellow  ' kuning
                Case 7 : .SelectionColor = Color.Purple  ' Ungu
            End Select
            Dim CurrentTime As String = Format(Date.Now, "yyyy-MM-dd  HH:mm:ss")
            Dim Message As String
            If UseTime Then
                Message = "[" & CurrentTime & "] " & pMsg
            Else
                Message = pMsg
            End If
            .AppendText(Space(pPos) & Message & vbCrLf)
            .ScrollToCaret()
        End With
    End Sub

    Private Function EmailToEmailCCPOMonthly(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""
        MdlConn.ReadConnection()
        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                " select 'AFF' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailAffiliate where AffiliateID='" & pAfffCode & "'" & vbCrLf & _
                " union all " & vbCrLf & _
                " --PASI TO -CC " & vbCrLf & _
                " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                " union all " & vbCrLf & _
                " --Supplier TO- CC " & vbCrLf & _
                " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail='' from ms_emailSupplier where SupplierID='" & Trim(pSupplierID) & "'"
        Dim ds As New DataSet
        ds = uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        End If
    End Function

    Private Sub sendEmailPOtoSupllierMonthly(ByVal pFileName As String, ByVal pAffiliate As String, ByVal ls_Message As String)
        Try
            Dim TempFilePath As String
            Dim TempFileName As String
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            TempFileName = pFileName

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCCPOMonthly(pAffiliate, "PASI", "")
            'To Supplier, CC Supplier
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
                Exit Sub
            End If
            If fromEmail = "" Then
                Exit Sub
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "Notification Error in PO Confirmation"

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
            GetSettingEmail("PO")
            'uf_GetNotification("11")
            'ls_Body = pLine1 & vbCr & pLine2 & vbCr & "PO No:" & pPONo & vbCr & pLine3 & vbCr & pLine4 & vbCr & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
            'mailMessage.Body = ls_Body
            'ls_Body = clsNotification.GetNotification("11", "", pPONo.Trim)
            mailMessage.Body = ls_Message
            Dim filename As String = TempFilePath & TempFileName
            mailMessage.Attachments.Add(New Attachment(filename))
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            'smtp.Host = "smtp.atisicloud.com"
            'smtp.Host = "mail.fast.net.id"
            smtp.Host = smtpClient
            If smtp.UseDefaultCredentials = True Then
                smtp.EnableSsl = True
            Else
                smtp.EnableSsl = False
                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
                smtp.Credentials = myCredential
            End If

            smtp.Port = portClient
            smtp.Send(mailMessage)

        Catch ex As Exception

        End Try

    End Sub

    Private Sub GetSettingEmail(ByVal ls_Value As String)
        Dim ls_SQL As String = ""
        MdlConn.ReadConnection()
        ls_SQL = "SELECT * FROM dbo.Ms_EmailSetting"
        Dim ds As New DataSet
        ds = uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            smtpClient = Trim(ds.Tables(0).Rows(0)("SMTP"))
            portClient = Trim(ds.Tables(0).Rows(0)("PORTSMTP"))
            usernameSMTP = If(IsDBNull(ds.Tables(0).Rows(0)("usernameSMTP")), "", ds.Tables(0).Rows(0)("usernameSMTP"))
            PasswordSMTP = If(IsDBNull(ds.Tables(0).Rows(0)("passwordSMTP")), "", ds.Tables(0).Rows(0)("passwordSMTP"))
        End If
    End Sub

    Private Sub UploadProcess()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""

        Try
            timerProcess.Enabled = False
            processTime = True

            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Batch Process", rtbProcess)

            '01. Upload data PO
            up_UploadProcess()

            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Batch Process", rtbProcess)

            timerProcess.Enabled = True

            txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            intervalpro = TimeSpan.FromSeconds(CDbl(txtTime.Text))
            Dim Last As Date = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")
            processTime = False
        Catch ex As Exception
            timerProcess.Enabled = True
            txtLast.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            intervalpro = TimeSpan.FromSeconds(CDbl(txtTime.Text))
            Dim Last As Date = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + intervalpro, "HH:mm:ss")
            processTime = False
        End Try

    End Sub

    Private Sub up_UploadProcess()
        Dim ErrMsg As String = ""
        Dim errSummary As String = ""
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application

        Dim ls_file As String

        Dim sheetNumber As Integer = 1
        Dim startTime As DateTime = Now

        Try
            screenName = "UploadProcess"

            Application.DoEvents()
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "Start Upload Process", rtbProcess)

            Log.WriteToProcessLog(startTime, screenName, "Start Upload Process")

            'clsPO.up_SendPODomestic(cfg, Log, cls, rtbProcess, txtAttachmentDOM.Text.Trim, txtSaveAsDOM.Text.Trim, screenName, ErrMsg, errSummary)
            Thread.Sleep(500)
            Dim di As New IO.DirectoryInfo(txtpath.Text)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.xlsm")
            Dim fi As IO.FileInfo

            Dim jmlFile As Integer = aryFi.Length

            For Each fi In aryFi
                ls_file = txtpath.Text & "\" & fi.Name

                ExcelBook = xlApp.Workbooks.Open(ls_file)
                ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                templateCode = ExcelSheet.Range("H1").Value.ToString & ""

                If templateCode = "POEM" Then
                    clsPOEM.up_Upload(ExcelSheet, cfg, Log, cls, rtbProcess, txtPathBackup.Text.Trim, txtBackupErrorFile.Text.Trim, templateCode, ErrMsg, errSummary)
                End If
            Next

            'If ErrMsg = "-" Then
            '    ErrMsg = "There is No PO data to process."
            'End If

            'If ErrMsg <> "" Then
            '    If ErrMsg = "There is No PO data to process." Then
            '        clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ErrMsg, rtbProcess)
            '        Log.WriteToProcessLog(startTime, screenName, ErrMsg)
            '    Else
            '        clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.ErrorMsg, ErrMsg, rtbProcess)
            '        Log.WriteToErrorLog(screenName, ErrMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            '        Log.WriteToProcessLog(startTime, screenName, ErrMsg, , , clsLog.ErrSeverity.ERR)
            '    End If
            'End If

        Catch ex As Exception
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, ex.Message, rtbProcess)
            Log.WriteToErrorLog(Me.Tag, ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            Log.WriteToProcessLog(startTime, screenName, ex.Message, , , clsLog.ErrSeverity.ERR)
        Finally
            clsGeneral.up_displayLog(clsGlobal.MsgTypeEnum.InformationMsg, "End Upload Process", rtbProcess)
            Log.WriteToProcessLog(startTime, screenName, "End Upload Process")
            Thread.Sleep(500)
        End Try
    End Sub
#End Region

#Region "Control Event"
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnAuto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuto.Click
        Me.Cursor = Cursors.WaitCursor
        UploadProcess()
        Me.Cursor = Cursors.Default
    End Sub
#End Region

End Class
