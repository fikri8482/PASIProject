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
Imports System.Web.Services.Protocols
Imports System.Management
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates

Imports System.Data.OleDb

Public Class frmUpload

#Region "Declaration"
    Dim stream As FileStream
    Dim excelReader As IDataReader
    Dim i As Integer
    Dim ii As Integer
    Dim Conn As String

    Dim templateCode As String = ""
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
    Dim tmpPathError As String = "D:\PASI EBWEB\02. TEMPORARY EMAIL ATTACHMENT\BACKUP ERROR FILE"

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String
    Dim iSplit As Integer
    Dim OrderSplit As String
    Dim statusSPlit As Boolean = False
#End Region

#Region "DECRALATION ON OFF"
    Dim IsNewRecEx As Boolean = True
    Dim NewDN As Boolean = True
#End Region

#Region "Initialization"
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim tmp As New clsTmp
        Label12.Text = Application.ProductVersion
        Try
            MdlConn.ReadConnection()
            'Conn = "Data Source=" & gs_DBserver & ";Initial Catalog=" & gs_DBdatabase & ";User ID=" & gs_DBuser & "" & ";pwd=" & gs_DBpass & "" & ";Connection Timeout=" & gi_connectionTimeOut & "" & ""
            lblDB.Text = "Server [" & Trim(gs_DBserver) & "], Database [" & gs_DBdatabase & "]"
            clrScreen()
            ds = clsTmpDB.Attachment(tmp)
            If ds Is Nothing Then
                txtpath.Text = "D:\Hadi\PASI\PASI\Template\Temporary Folder"
                txtPathBackup.Text = "D:\Hadi\PASI\PASI\Template\Backup Folder"
                txtTime.Text = 60
            Else
                txtpath.Text = Trim(ds.Tables(0).Rows(0)("AttachmentFolder"))
                txtPathBackup.Text = Trim(ds.Tables(0).Rows(0)("AttachmentBackupFolder"))
                'txtpath.Text = "D:\PASI EBWEB\02. TEMPORARY ATT"
                'txtPathBackup.Text = "D:\PASI EBWEB\02. TEMPORARY ATT\BACKUP"
                txtTime.Text = ds.Tables(0).Rows(0)("Interval")
            End If

            'txtpath.Text = "\\192.168.0.5\d\PASI EBWEB\TEMPORARY ATT"
            'txtPathBackup.Text = "\\192.168.0.5\d\PASI EBWEB\TEMPORARY ATT\BACKUP"
            'tmpPathError = txtpath.Text & "\\192.168.0.5\d\PASI EBWEB\TEMPORARY ATT\BACKUP ERROR FILE"

            'txtTime.Text = 15        
            StartScheduler()
        Catch ex As Exception
            txtMsg.ForeColor = Color.Red
            txtMsg.Text = "Error load data"
        End Try

    End Sub
#End Region

#Region "Procedures"
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
        Dim tmpAllNo As String = "", autoApprove As Integer, Status As Integer, Boxdetail As String, isSplit As Integer
        Dim AllfileNameValid As String = "", AllfileNameInvalid As String = ""
        Dim poExportStatus As Integer
        Dim AdaQty As Boolean = False
        Dim postatusExportExist As Boolean = False
        Dim ExistPart As Boolean = False
        Dim jmlSplit As Integer = 0
        Dim StatusSplit As Boolean = False
        Dim pub_SuratJalanNo As Boolean = True

        statusSave = True
        MsgstatusSave = ""

        Application.DoEvents()
        msgInfo = "search file upload process (file format .xlsm)"
        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

        Dim di As New IO.DirectoryInfo(txtpath.Text)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xlsm")
        Dim fi As IO.FileInfo

        Dim jmlFile As Integer = aryFi.Length

        Application.DoEvents()
        gridProcess(Rtb1, 1, 0, jmlFile & " file upload", True)
        gridProcess(Rtb1, 1, 0, "End search file upload process (file format .xlsm)" & vbCrLf, True)

        Dim xlApp = New Excel.Application
        Dim ExcelBook As Excel.Workbook = Nothing
        Dim ExcelSheet As Excel.Worksheet = Nothing

        For Each fi In aryFi

            Dim StatusKirim As Boolean = True

            Dim tmp As New clsTmp
            Dim i As Integer
            Dim z As Integer
            Dim startRow As Long = 0
            Dim inputMaster As Boolean = False, inputMaster2 As Boolean = False, _
                inputMaster3 As Boolean = False, inputMaster4 As Boolean = False, inputMaster5 As Boolean = False
            Dim statusApprove As String = ""
            Dim Jumlah As Integer = 0
            Dim ls_file As String = txtpath.Text & "\" & fi.Name

            msgInfo = "open file upload process..."
            gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

            Try
                ExcelBook = xlApp.Workbooks.Open(ls_file)
                Jumlah = ExcelBook.Worksheets.Count
            Catch ex As Exception
                gridProcess(Rtb1, 1, 0, "Failed " & msgInfo & vbCrLf, True)

                'MOVE FILE
                msgInfo = "move data Failed Open " & fi.Name
                gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                up_XMLFile_Copy(txtpath.Text & "\", tmpPathError & "\", fi.Name)
                gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                Application.DoEvents()
                LogError(ex, fi.Name, "File Cannot Open", "")
                Exit Sub
            End Try

            gridProcess(Rtb1, 1, 0, Jumlah & " Sheet Found In " & fi.Name, True)

            gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

            'Looping Jumlah Sheet
            msgInfo = "read file upload process..."
            gridProcess(Rtb1, 1, 0, "Start " & msgInfo & vbCrLf, True)

            For sheetNumber = 1 To Jumlah

                Try
                    statusSave = True
                    tmp = New clsTmp

                    MsgstatusSave = ""

                    Me.Cursor = Cursors.WaitCursor
                    txtMsg.Text = ""
                    txtMsg.ForeColor = Color.Red

                    pub_SuratJalanNo = True

                    excelName = fi.Name

                    Application.DoEvents()

                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                    Application.DoEvents()

                    Dim ds As New DataSet
                    Dim ds1 As New DataSet

                    Application.DoEvents()

                    templateCode = ExcelSheet.Range("H1").Value.ToString & ""
                    tmp.AffiliateID = ExcelSheet.Range("H3").Value.ToString & ""

                    If templateCode = "REC-EX" Or templateCode = "TALLY" Then
                        tmp.AffiliateID = ExcelSheet.Range("H3").Value.ToString & ""
                        tmp.AffiliateID = clsTmpDB.getAffiliateID(tmp.AffiliateID)
                    End If

                    If templateCode <> "TALLY" Then
                        tmp.SupplierID = ExcelSheet.Range("H5").Value.ToString & ""
                    End If

                    '=========== KHUSUS RECEIVING EXPORT ===============
                    Dim recStatus As Boolean = False

                    If IsNewRecEx = True Then
                        If templateCode = "REC-EX" Then
                            If ExcelSheet.Range("I28").Value Is Nothing Then
                                tmp.SuratJalanNo = ""
                                pub_SuratJalanNo = False
                                GoTo suratJalanKeluar
                            Else
                                tmp.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I28").Value.ToString & "", 20)
                                tmp.SuratJalanNo = tmp.SuratJalanNo.Trim.Replace(vbCr, "").Replace(vbLf, "")
                            End If

                            If ExcelSheet.Range("AE15").Value Is Nothing Then
                                tmp.PONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                            Else
                                tmp.PONo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE15").Value.ToString & "", 20)
                            End If

                            If ExcelSheet.Range("AE13").Value Is Nothing Then
                                tmp.OrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                            Else
                                tmp.OrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                            End If

                            tmp.ForwarderID = Microsoft.VisualBasic.Left(ExcelSheet.Range("H4").Value.ToString & "", 20)

                            If clsTmpDB.CekRecEX(tmp) = True Then
                                recStatus = True
                            Else
                                recStatus = False
                            End If

                            If clsTmpDB.CekPOEX(tmp) = False Then
                                statusSave = False
                                MsgstatusSave = "PO"
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
                        tmpAllNo = "KANBAN NO: " & tmp.KanbanNo & " In Sheet " & sheetNumber

                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

                    ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then

                        startRow = 39
                        If ExcelSheet.Range("J27").Value Is Nothing Then
                            tmp.SuratJalanNo = ""
                            pub_SuratJalanNo = False
                            GoTo suratJalanKeluar
                        Else
                            tmp.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("J27").Value.ToString & "", 20)
                            tmp.SuratJalanNo = tmp.SuratJalanNo.Trim.Replace(vbCr, "").Replace(vbLf, "")
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
                        If ExcelSheet.Range("X29").Value Is Nothing Then
                            tmp.TotalBox = 0
                        Else
                            tmp.TotalBox = ExcelSheet.Range("X29").Value.ToString & ""
                        End If

                        Application.DoEvents()
                        If templateCode = "DO" Then
                            tmpAllNo = "DO NO: " & Trim(tmp.SuratJalanNo) & " In Sheet " & sheetNumber
                        Else
                            tmpAllNo = "DO (PO non Kanban) NO: " & Trim(tmp.SuratJalanNo) & " In Sheet " & sheetNumber
                        End If

                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

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
                        tmp.Period = tempDate
                        tmp.PONo = ExcelSheet.Range("I9").Value.ToString & ""

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
                            tmpAllNo = "PO NO: " & Trim(tmp.PONo) & " In Sheet " & sheetNumber
                            msgInfo = "processing upload " & tmpAllNo
                            gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                        Else
                            tmpAllNo = "POREV NO: " & Trim(tmp.PORevNo) & " In Sheet " & sheetNumber
                            msgInfo = "processing upload " & tmpAllNo
                            gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                        End If

                    ElseIf templateCode = "POEM" Then 'PO Monthly

                        startRow = 37
                        tmp.OrderNo = ExcelSheet.Range("I9").Value.ToString & ""

                        If ExcelSheet.Range("I11").Value IsNot Nothing Then
                            If ExcelSheet.Range("I11").Value.ToString & "" = "" Then
                                tmp.PONo = ExcelSheet.Range("I9").Value.ToString & ""
                            Else
                                tmp.PONo = ExcelSheet.Range("I11").Value.ToString & ""
                            End If
                        Else
                            tmp.PONo = ExcelSheet.Range("I9").Value.ToString & ""
                        End If

                        tmp.Period = Format(ExcelSheet.Range("AE9").Value, "yyyy-MM-dd")
                        'Affiliate ID
                        If ExcelSheet.Range("I16").Value Is Nothing Then
                            tmp.AffiliateName = ""
                        Else
                            tmp.AffiliateName = ExcelSheet.Range("I16").Value.ToString & ""
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
                            tmpAllNo = "ORDER NO: " & Trim(tmp.PONo) & " In Sheet " & sheetNumber
                            msgInfo = "processing upload " & tmpAllNo
                            gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                        End If

                        postatusExportExist = False
                        isSplit = clsTmpDB.CekSplitPO("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.AffiliateID)

                    ElseIf templateCode = "POEE" Then 'PO Emergency

                        startRow = 37
                        tmp.PONo = ExcelSheet.Range("S1").Value.ToString & ""
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
                            tmpAllNo = "ORDER NO: " & Trim(tmp.PONo) & " In Sheet " & sheetNumber
                            msgInfo = "processing upload " & tmpAllNo
                            gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                        End If
                    ElseIf templateCode = "INV" Then

                        startRow = 36
                        tmp.InvoiceNo = ExcelSheet.Range("I11").Value.ToString & ""
                        tmp.InvoiceNo = tmp.InvoiceNo.Trim.Replace(vbCr, "").Replace(vbLf, "")

                        Try
                            tmp.InvoiceDate = Format(ExcelSheet.Range("W11").Value, "yyyy-MM-dd")
                        Catch ex As Exception
                            statusSave = False
                        End Try

                        If ExcelSheet.Range("AM11").Value Is Nothing Then
                            tmp.PaymentItem = ""
                        Else
                            tmp.PaymentItem = ExcelSheet.Range("AM11").Value.ToString & ""
                        End If
                        Try
                            tmp.DueDate = Format(ExcelSheet.Range("BA11").Value, "yyyy-MM-dd")
                        Catch ex As Exception
                            statusSave = False
                        End Try

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
                        tmpAllNo = "INVOICE NO: " & Trim(tmp.InvoiceNo) & " In Sheet " & sheetNumber
                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

                    ElseIf templateCode = "DO-EX" Then 'DIAN EXPORT

                        startRow = 34
                        If ExcelSheet.Range("I28").Value Is Nothing Then
                            tmp.SuratJalanNo = ""
                            pub_SuratJalanNo = False
                            GoTo suratJalanKeluar
                        Else
                            tmp.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I28").Value.ToString & "", 20)
                            tmp.SuratJalanNo = tmp.SuratJalanNo.Trim.Replace(vbCr, "").Replace(vbLf, "")
                        End If

                        If ExcelSheet.Range("AE13").Value Is Nothing Then
                            tmp.OrderNo = ""
                        Else
                            tmp.OrderNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("AE13").Value.ToString & "", 20)
                        End If

                        If ExcelSheet.Range("AP11").Value Is Nothing Then
                            tmp.CommercialCls = "1"
                        Else
                            tmp.CommercialCls = Trim(ExcelSheet.Range("AP11").Value.ToString)
                            tmp.CommercialCls = IIf(tmp.CommercialCls.ToUpper = "YES", "1", "0")
                        End If

                        'cek apakah SJ sudah ada atau belum
                        If clsTmpDB.CekData(tmp, "DOSupplier_Master_Export") = 1 Then
                            statusSave = False
                            MsgstatusSave = "ALREADY"
                        Else
                            statusSave = True
                            MsgstatusSave = ""
                        End If
                        'cek apakah SJ sudah ada atau belum

                        Application.DoEvents()
                        If templateCode = "DO-EX" Then
                            tmpAllNo = "DO NO: " & Trim(tmp.SuratJalanNo) & " In Sheet " & sheetNumber
                        End If

                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                    ElseIf templateCode = "REC-EX" Then 'DIAN EXPORT
                        startRow = 34

                        If ExcelSheet.Range("I28").Value Is Nothing Then
                            tmp.SuratJalanNo = ""
                        Else
                            tmp.SuratJalanNo = Microsoft.VisualBasic.Left(ExcelSheet.Range("I28").Value.ToString & "", 20)
                            tmp.SuratJalanNo = tmp.SuratJalanNo.Trim.Replace(vbCr, "").Replace(vbLf, "")
                        End If

                        tmp.ForwarderID = Microsoft.VisualBasic.Left(ExcelSheet.Range("H4").Value.ToString & "", 20)

                        'cek apakah SJ sudah ada atau belum
                        If clsTmpDB.CekData2(tmp, "ReceiveForwarder_Master") = 1 Then
                            statusSave = False
                            MsgstatusSave = "ALREADY"
                        End If

                        'CEK APAKAH SJ SUDAH ADA DI DN SUPPLIER
                        If clsTmpDB.CekData(tmp, "DOSupplier_Master_Export") = 1 Then
                            statusDNSupp = False
                        Else
                            statusSave = False
                            MsgstatusSave = "DATA DN NO EXIST"
                            statusDNSupp = True
                        End If

                        Application.DoEvents()
                        If templateCode = "REC-EX" Then
                            tmpAllNo = "SJ NO: " & Trim(tmp.SuratJalanNo) & " In Sheet " & sheetNumber
                        End If

                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                    ElseIf templateCode = "INV-EX" Then 'DIAN Export
                        startRow = 36
                        tmp.ForwarderID = ExcelSheet.Range("H4").Value.ToString.Replace(vbCr, "").Replace(vbLf, "") & ""
                        tmp.InvoiceNo = ExcelSheet.Range("I11").Value.ToString.Replace(vbCr, "").Replace(vbLf, "") & ""
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
                        tmpAllNo = "INVOICE NO: " & Trim(tmp.InvoiceNo) & " In Sheet " & sheetNumber
                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                    ElseIf templateCode = "TALLY" Then 'DIAN Export
                        startRow = 23
                        tmp.ForwarderID = ExcelSheet.Range("H4").Value.ToString & ""

                        tmp.InvoiceNo = IIf(IsNothing(ExcelSheet.Range("I8").Value), "", ExcelSheet.Range("I8").Value)
                        tmp.Vassel = "" 'IIf(IsNothing(ExcelSheet.Range("AA8").Value), "", ExcelSheet.Range("AA8").Value)
                        'Request Pak Andik, Vessel Name di pindah ke VesselNo/Voyyage
                        tmp.NamaKapal = IIf(IsNothing(ExcelSheet.Range("AA8").Value), "", ExcelSheet.Range("AA8").Value) 'IIf(IsNothing(ExcelSheet.Range("AQ8").Value), "", ExcelSheet.Range("AQ8").Value)
                        tmp.StuffingDate = IIf(IsNothing(ExcelSheet.Range("AQ10").Value), "", ExcelSheet.Range("AQ10").Value)

                        tmp.ContainerNo = IIf(IsNothing(ExcelSheet.Range("I10").Value), "", ExcelSheet.Range("I10").Value)
                        tmp.DONo = ""

                        tmp.SealNo = IIf(IsNothing(ExcelSheet.Range("I12").Value), "", ExcelSheet.Range("I12").Value)
                        tmp.SizeContainer = Replace(IIf(IsDBNull(ExcelSheet.Range("AA10").Value), "", ExcelSheet.Range("AA10").Value), "'", " ")

                        tmp.Tare = IIf(IsNothing(ExcelSheet.Range("I14").Value), 0, ExcelSheet.Range("I14").Value)
                        tmp.ETDJakarta = IIf(IsNothing(ExcelSheet.Range("AA12").Value), "", ExcelSheet.Range("AA12").Value)

                        tmp.Gross = IIf(IsNothing(ExcelSheet.Range("I16").Value), 0, ExcelSheet.Range("I16").Value)
                        tmp.ShippingLine = IIf(IsNothing(ExcelSheet.Range("AA14").Value), "", ExcelSheet.Range("AA14").Value)

                        tmp.TotalCarton = IIf(IsNothing(ExcelSheet.Range("I18").Value), 0, ExcelSheet.Range("I18").Value)
                        tmp.DestinationPort = IIf(IsNothing(ExcelSheet.Range("AA16").Value), "", ExcelSheet.Range("AA16").Value)

                        Application.DoEvents()
                        tmpAllNo = "Tally: " & Trim(tmp.InvoiceNo) & " In Sheet " & sheetNumber
                        msgInfo = "processing upload " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)
                    End If

                    Try
                        Using SQLCon As New SqlConnection(uf_GetConString)
                            SQLCon.Open()

                            Dim SQLCom As SqlCommand = SQLCon.CreateCommand
                            Dim SQLTrans As SqlTransaction
                            Dim StartCol As Integer
                            Dim write As Boolean

                            write = True

                            SQLTrans = SQLCon.BeginTransaction
                            SQLCom.Connection = SQLCon
                            SQLCom.Transaction = SQLTrans
                            z = 0
                            If templateCode <> "POEM" Then StartCol = 4

                            For z = StartCol To 4
                                AdaQty = False
                                StatusSplit = False
                                For i = startRow To 10000
                                    Dim Nomor As String = IIf(IsNothing(ExcelSheet.Range("B" & i).Value), "", ExcelSheet.Range("B" & i).Value)
                                    If Nomor = "E" Or Nomor = "" Then
                                        Application.DoEvents()
                                        If write = True Then
                                            gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)
                                            write = False
                                        End If

                                        Application.DoEvents()
                                        If write = True Then
                                            msgInfo = "read file upload process Sheet " & sheetNumber
                                            gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)
                                            write = False
                                        End If

                                        'EDIT FOR GOOD RECEIVE EXPORT
                                        If templateCode = "REC-EX" Then
                                            clsTmpDB.updateExcelMasterRecevingEX(tmp, SQLCom)
                                            If statusDNSupp = True Then
                                                Dim xlRange1 As Excel.Range = Nothing
                                                xlRange1 = CType(ExcelSheet.Rows(i), Excel.Range)
                                                xlRange1.Delete()
                                            End If
                                        End If

                                        Me.Cursor = Cursors.Default
                                        Exit For
                                    ElseIf ExcelSheet.Range("D" & i).Value Is Nothing And templateCode <> "TALLY" Then
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
                                            gridProcess(Rtb1, 1, 0, "End " & msgInfo, True)

                                            Application.DoEvents()
                                            msgInfo = "read file upload process Sheet " & sheetNumber & " ..."
                                            gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                                            Me.Cursor = Cursors.Default
                                            Exit For
                                        End If

                                    ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then
                                        Dim Moq, QtyBox As Integer
                                        Moq = 0 : QtyBox = 0
                                        tmp.PartNo = ExcelSheet.Range("P" & i).Value.ToString & ""
                                        tmp.PONo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                        tmp.UnitCls = clsTmpDB.UnitCls(ExcelSheet.Range("AD" & i).Value.ToString & "")
                                        tmp.KanbanNo = ExcelSheet.Range("L" & i).Value.ToString & ""
                                        tmp.DOQty = IIf(IsNumeric(ExcelSheet.Range("AO" & i).Value) = False, 0, ExcelSheet.Range("AO" & i).Value)
                                        Moq = clsTmpDB.uf_GetMOQ(1, tmp.PONo, tmp.PartNo, tmp.SupplierID, tmp.AffiliateID)
                                        QtyBox = clsTmpDB.uf_GetQtybox(1, tmp.PONo, tmp.PartNo, tmp.SupplierID, tmp.AffiliateID)
                                        'MsgBox(tmp.DOQty.ToString())
                                        If tmp.DOQty > 0 Then
                                            Dim PlanQty As Double = IIf(IsNumeric(ExcelSheet.Range("AK" & i).Value) = False, 0, ExcelSheet.Range("AK" & i).Value)
                                            'MsgBox(PlanQty.ToString())
                                            If tmp.DOQty > PlanQty Then
                                                statusSave = False
                                                MsgstatusSave = "OVER"
                                                Exit For
                                            End If
                                            'MsgBox(MsgstatusSave)
                                            If clsTmpDB.POReceive(tmp.PONo, tmp.PartNo, tmp.AffiliateID, tmp.SupplierID, tmp.DOQty) = True Then
                                                statusSave = False
                                                MsgstatusSave = "OVER"
                                                Exit For
                                            End If
                                            'MsgBox(MsgstatusSave)
                                            tmp.POKanbanCls = clsTmpDB.POKanbanCls(tmp.PONo, tmp.PartNo, tmp.AffiliateID, tmp.SupplierID)
                                            If inputMaster2 = False Then
                                                clsTmpDB.insertMasterDO(tmp, SQLCom)
                                                inputMaster2 = True
                                            End If
                                            clsTmpDB.insertDetailDO(tmp, Moq, QtyBox, SQLCom)
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
                                            gridProcess(Rtb1, 1, 0, "End " & msgInfo, True)

                                            Application.DoEvents()
                                            msgInfo = "read file upload process Sheet " & sheetNumber & " ..."
                                            gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

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

                                        autoApprove = clsTmpDB.CekPOAutoApprover("dbo.PORev_Master", tmp.PONo, tmp.SupplierID, statusApprove)
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
                                            'Application.DoEvents()
                                            'gridProcess(Rtb1, 1, 0, "End " & msgInfo, True)

                                            Application.DoEvents()
                                            msgInfo = "read file upload process Sheet " & sheetNumber & " ..."
                                            gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

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
                                                If ExcelSheet.Range("AE" & i).Value Is Nothing And postatusExportExist = False Then GoTo NextProcess
                                                tmp.POQty = Val(ExcelSheet.Range("AE" & i).Value)
                                                tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                                tmp.ETDSplit = ExcelSheet.Range("Z" & i).Value & ""
                                                tmp.QtySplit = Val(ExcelSheet.Range("AE" & i).Value)

                                                postatusExportExist = True : AdaQty = True : jmlSplit = 1
                                            ElseIf z = 1 Then
                                                If ExcelSheet.Range("AO" & i).Value Is Nothing And postatusExportExist = False Then GoTo NextProcess
                                                If ExcelSheet.Range("AO" & i).Value Is Nothing And postatusExportExist = True Then GoTo NextProcess

                                                tmp.POQty = Val(ExcelSheet.Range("AO" & i).Value)
                                                tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                                tmp.ETDSplit = ExcelSheet.Range("AJ" & i).Value & ""
                                                tmp.QtySplit = Val(ExcelSheet.Range("AO" & i).Value)

                                                If AdaQty = False Then
                                                    If postatusExportExist = True Then
                                                        jmlSplit = clsTmpDB.CekSplitPO("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.AffiliateID)
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                        StatusSplit = True
                                                        AdaQty = True
                                                    Else
                                                        tmp.OrderNo = tmp.PONo
                                                        AdaQty = True
                                                    End If
                                                Else
                                                    If postatusExportExist = True And StatusSplit = True Then
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                        AdaQty = True
                                                    Else
                                                        tmp.OrderNo = tmp.PONo
                                                        AdaQty = True
                                                    End If
                                                End If

                                            ElseIf z = 2 Then
                                                If ExcelSheet.Range("AY" & i).Value Is Nothing And postatusExportExist = False Then Exit For
                                                If ExcelSheet.Range("AY" & i).Value Is Nothing And postatusExportExist = True Then GoTo NextProcess
                                                tmp.POQty = Val(ExcelSheet.Range("AY" & i).Value)
                                                tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                                tmp.ETDSplit = ExcelSheet.Range("AT" & i).Value & ""
                                                tmp.QtySplit = Val(ExcelSheet.Range("AY" & i).Value)
                                                If AdaQty = False Then
                                                    If postatusExportExist = True Then
                                                        jmlSplit = clsTmpDB.CekSplitPO("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.AffiliateID)
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                        StatusSplit = True
                                                        AdaQty = True
                                                    Else
                                                        tmp.OrderNo = tmp.PONo
                                                        AdaQty = True
                                                    End If
                                                Else
                                                    If postatusExportExist = True And StatusSplit = True Then
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                        AdaQty = True
                                                    Else
                                                        tmp.OrderNo = tmp.PONo
                                                        AdaQty = True
                                                    End If
                                                End If
                                            ElseIf z = 3 Then
                                                If ExcelSheet.Range("BI" & i).Value Is Nothing And postatusExportExist = False Then Exit For
                                                If ExcelSheet.Range("BI" & i).Value Is Nothing And postatusExportExist = True Then GoTo NextProcess
                                                tmp.POQty = Val(ExcelSheet.Range("BI" & i).Value)
                                                tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                                tmp.ETDSplit = ExcelSheet.Range("BD" & i).Value & ""
                                                tmp.QtySplit = Val(ExcelSheet.Range("BI" & i).Value)

                                                If AdaQty = False Then
                                                    If postatusExportExist = True Then
                                                        jmlSplit = clsTmpDB.CekSplitPO("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.AffiliateID)
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                        StatusSplit = True
                                                        AdaQty = True
                                                    Else
                                                        tmp.OrderNo = tmp.PONo
                                                        AdaQty = True
                                                    End If
                                                Else
                                                    If postatusExportExist = True And StatusSplit = True Then
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                        AdaQty = True
                                                    Else
                                                        tmp.OrderNo = tmp.PONo
                                                        AdaQty = True
                                                    End If
                                                End If
                                            ElseIf z = 4 Then
                                                If ExcelSheet.Range("BS" & i).Value Is Nothing And postatusExportExist = False Then Exit For
                                                If ExcelSheet.Range("BS" & i).Value Is Nothing And postatusExportExist = True Then GoTo NextProcess

                                                tmp.POQty = Val(ExcelSheet.Range("BS" & i).Value)
                                                tmp.POQtyOld = Val(ExcelSheet.Range("U" & i).Value)
                                                tmp.ETDSplit = ExcelSheet.Range("BN" & i).Value & ""
                                                tmp.QtySplit = Val(ExcelSheet.Range("BS" & i).Value)
                                                If AdaQty = False Then
                                                    If postatusExportExist = True Then
                                                        jmlSplit = clsTmpDB.CekSplitPO("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.AffiliateID)
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
                                                        StatusSplit = True
                                                        AdaQty = True
                                                    Else
                                                        tmp.OrderNo = tmp.PONo
                                                        AdaQty = True
                                                    End If
                                                Else
                                                    If postatusExportExist = True And StatusSplit = True Then
                                                        tmp.OrderNo = tmp.PONo & "-" & jmlSplit + 1
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

                                            ExistPart = clsTmpDB.CekExistsPart("dbo.PO_Detail_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, tmp.PartNo)
                                            If ExistPart = True Then
                                                If (postatusExportExist = True And tmp.POQty <> 0 And jmlSplit > 0) Or (postatusExportExist = False) Or (postatusExportExist = True And jmlSplit = 0) Then
                                                    autoApprove = clsTmpDB.CekPOMonthlyAutoApprover("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, statusApprove)
                                                    poExportStatus = clsTmpDB.CekPOMonthlyStatus("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, statusApprove)
                                                    ii = 0
                                                    iSplit = z + 1
                                                    OrderSplit = ""

                                                    If autoApprove = 0 And poExportStatus = 0 Then
                                                        clsTmpDB.insertMasterPOMonthlyUpload(tmp, SQLCom, "dbo.PO_MasterUpload_Export", templateCode)
                                                        clsTmpDB.UpdatePOMonthlyMaster(tmp, SQLCom, statusApprove, "dbo.PO_Master_Export")
                                                        Status = 0
                                                        Status = clsTmpDB.CekPOMonthlyEXIST("dbo.PO_Master_Export", tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, statusApprove)
                                                        If Status = 0 Then
                                                            clsTmpDB.insertMasterPOMonthlyAfterUpload(tmp, SQLCom, "dbo.PO_Master_Export", templateCode, iSplit)
                                                            clsTmpDB.UpdatePOMonthlyMaster(tmp, SQLCom, statusApprove, "dbo.PO_Master_Export")
                                                        End If
                                                        inputMaster3 = True

                                                        If statusApprove = "0" Then
                                                            clsTmpDB.insertDetailPOMonthlyUpload(tmp, _
                                                            ExcelSheet.Range("I9" & i).Value, ExcelSheet.Range("I16" & i).Value, ExcelSheet.Range("I11" & i).Value, ExcelSheet.Range("AE13" & i).Value, ExcelSheet.Range("D41" & i).Value, "", "", "", _
                                                            "", "", "", "", "", "", "", ExcelSheet.Range("AC41" & i).Value, ExcelSheet.Range("AC41" & i).Value, "", ExcelSheet.Range("AG41" & i).Value, ExcelSheet.Range("AK41" & i).Value, ExcelSheet.Range("AO41" & i).Value, _
                                                            "dbo.PO_DetailUpload_Export", templateCode, SQLCom)
                                                            'upload ke PO_master jika belum ada
                                                            If StatusSplit = True Then
                                                                clsTmpDB.insertDetailPOMonthlyAfterUpload(tmp, _
                                                                ExcelSheet.Range("I9" & i).Value, ExcelSheet.Range("I16" & i).Value, ExcelSheet.Range("I11" & i).Value, ExcelSheet.Range("AE13" & i).Value, ExcelSheet.Range("D41" & i).Value, "", "", "", _
                                                                "", "", "", "", "", "", "", ExcelSheet.Range("AC41" & i).Value, ExcelSheet.Range("AC41" & i).Value, "", ExcelSheet.Range("AG41" & i).Value, ExcelSheet.Range("AK41" & i).Value, ExcelSheet.Range("AO41" & i).Value, _
                                                                "dbo.PO_Detail_Export", templateCode, SQLCom, StatusSplit)
                                                            End If
                                                            postatusExportExist = True
                                                        End If
                                                    Else
                                                        Application.DoEvents()
                                                        'gridProcess(Rtb1, 1, 0, "End " & msgInfo, True)

                                                        Application.DoEvents()
                                                        'msgInfo = "read file upload process Sheet " & sheetNumber & " ..."
                                                        'gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                                                        Me.Cursor = Cursors.Default
                                                        Exit For

                                                    End If
                                                End If
                                            End If
                                        End If
                                    ElseIf templateCode = "POEE" Then 'PO EMERGENCY
                                        statusApprove = 0

                                        tmp.PartNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                        tmp.POQty = Val(ExcelSheet.Range("AE" & i).Value)
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
                                                "dbo.PO_DetailUpload_Export", templateCode, SQLCom)
                                            End If

                                            'i = i + 1

                                        Else
                                            Application.DoEvents()
                                            'gridProcess(Rtb1, 1, 0, "End " & msgInfo, True)

                                            Application.DoEvents()
                                            'msgInfo = "read file upload process Sheet " & sheetNumber & " ..."
                                            'gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                                            Me.Cursor = Cursors.Default
                                            Exit For

                                        End If

                                    ElseIf templateCode = "INV" Then 'INVOICE
                                        If statusSave = True Then
                                            tmp.SuratJalanNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                            tmp.PONo = ExcelSheet.Range("J" & i).Value.ToString & ""
                                            tmp.KanbanNo = ExcelSheet.Range("Q" & i).Value.ToString & ""
                                            tmp.PartNo = ExcelSheet.Range("U" & i).Value.ToString & ""
                                            tmp.InvoiceQty = ExcelSheet.Range("AQ" & i).Value
                                            If ExcelSheet.Range("AT" & i).Value Is Nothing Then
                                                tmp.InvCurrCls = "03"
                                            Else
                                                tmp.InvCurrCls = clsTmpDB.CurrCls(ExcelSheet.Range("AT" & i).Value.ToString & "")
                                            End If

                                            tmp.InvPrice = ExcelSheet.Range("AV" & i).Value
                                            tmp.InvAmount = ExcelSheet.Range("AZ" & i).Value
                                            ds = clsTmpDB.StatusDelivery(tmp.PONo)

                                            If ds Is Nothing Then

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
                                            clsTmpDB.insertDetailInvoice(tmp, SQLCom)
                                        End If
                                    ElseIf templateCode = "DO-EX" Then 'Dian Export
                                        If statusSave = True Then
                                            If Trim(tmp.ForwarderID) <> "SEIWA" Then
                                                If ExcelSheet.Range("B" & i).Value.ToString <> "E" Then
                                                    tmp.PartNo = Trim(ExcelSheet.Range("I" & i).Value.ToString) & ""

                                                    If NewDN = False Then
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
                                                            Boxdetail = Trim(Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 8))
                                                            tmp.m_JmlBox = CDbl(ExcelSheet.Range("AM" & i).Value.ToString) / CDbl(ExcelSheet.Range("AC" & i).Value.ToString)

                                                            clsTmpDB.insertDetailDOEX(tmp, "1", SQLCom)

                                                            If IsNewRecEx = True Then
                                                                For i_loopBox = 0 To tmp.m_JmlBox - 1
                                                                    clsTmpDB.insertDetailDOBoxEX(tmp, ls_boxno + Microsoft.VisualBasic.Right("0000000" & (Microsoft.VisualBasic.Right(Boxdetail, 7) + i_loopBox), 7), 1, SQLCom)
                                                                Next
                                                            End If
                                                            'looping Box No
                                                        End If

                                                    Else 'DN BARU
                                                        tmp.OrderNo = Trim(ExcelSheet.Range("AE" & 13).Value.ToString) & ""

                                                        If ExcelSheet.Range("AE" & 15).Value IsNot Nothing Then
                                                            tmp.PONo = Trim(ExcelSheet.Range("AE" & 15).Value.ToString) & ""
                                                        Else
                                                            tmp.PONo = ""
                                                        End If
                                                        If tmp.PONo = "" Then
                                                            tmp.PONo = tmp.OrderNo
                                                        End If

                                                        tmp.UnitCls = clsTmpDB.UnitCls(ExcelSheet.Range("AC" & i).Value.ToString & "")
                                                        tmp.DOQty = IIf(IsNumeric(ExcelSheet.Range("AO" & i).Value) = False, 0, ExcelSheet.Range("AO" & i).Value)
                                                        'tmp.QtyBox = clsTmpDB.QtyBox(tmp.SupplierID, tmp.AffiliateID, tmp.PartNo)
                                                        tmp.QtyBox = clsTmpDB.uf_GetQtybox(2, tmp.PONo, tmp.PartNo, tmp.SupplierID, tmp.AffiliateID)

                                                        If tmp.DOQty > 0 Then
                                                            If (tmp.DOQty Mod tmp.QtyBox) <> 0 Then
                                                                statusSave = False
                                                                MsgstatusSave = "QTY MOD QTYBOX, NOT MATCH"
                                                                Exit For
                                                            End If

                                                            If inputMaster2 = False Then
                                                                clsTmpDB.insertMasterDOEX(tmp, SQLCom)
                                                                inputMaster2 = True
                                                            End If

                                                            'looping Box No
                                                            Dim i_loopBox As Integer = 0
                                                            Dim startBox As Integer = Microsoft.VisualBasic.Right(Trim(ExcelSheet.Range("W" & i).Value.ToString), 7)
                                                            Dim ls_boxno As String = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 8)

                                                            If Microsoft.VisualBasic.Mid(Trim(ExcelSheet.Range("W" & i).Value.ToString), 2, 1) <> "0" Then
                                                                ls_boxno = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 2)
                                                            Else
                                                                ls_boxno = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 1)
                                                            End If

                                                            'cek apakah part dan Box sudah pernah diupload
                                                            tmp.m_BoxNo = Trim(ExcelSheet.Range("W" & i).Value.ToString) & ""
                                                            Boxdetail = Trim(Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("W" & i).Value.ToString), 9))
                                                            tmp.m_JmlBox = CDbl(ExcelSheet.Range("AO" & i).Value.ToString) / CDbl(ExcelSheet.Range("AE" & i).Value.ToString)

                                                            Dim ls_SeqNo As Integer = clsTmpDB.CekDataDetailDNSupplier(tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, tmp.PartNo, (ls_boxno + Microsoft.VisualBasic.Right("000000" & (startBox + i_loopBox), 7)), tmp.SuratJalanNo)

                                                            clsTmpDB.insertDetailDOEX(tmp, ls_SeqNo, SQLCom)

                                                            If IsNewRecEx = True Then
                                                                For i_loopBox = 0 To tmp.m_JmlBox - 1
                                                                    If clsTmpDB.CekBoxNoDNSupplier(tmp.PartNo, ls_boxno + Microsoft.VisualBasic.Right("000000" & (startBox + i_loopBox), 7), tmp.OrderNo) = 0 Then
                                                                        statusSave = False
                                                                        MsgstatusSave = "BOX NO NOT MATCH"
                                                                        Exit For
                                                                    End If
                                                                    'Cek Box Sudah pernah Input
                                                                    If Trim(ExcelSheet.Range("A7").Value.ToString) = "REMAINING DELIVERY NOTE (EXPORT)" Then
                                                                        'DN Remaining
                                                                        If clsTmpDB.CekDataDNremaining(tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, tmp.PartNo, (ls_boxno + Microsoft.VisualBasic.Right("000000" & (startBox + i_loopBox), 7))) = 1 Then
                                                                            statusSave = False
                                                                            MsgstatusSave = "BOX ALREADY"
                                                                            Exit For
                                                                        End If
                                                                    Else
                                                                        'DN Normal
                                                                        If clsTmpDB.CekDataDNSupplier(tmp.PONo, tmp.OrderNo, tmp.SupplierID, tmp.AffiliateID, tmp.PartNo, (ls_boxno + Microsoft.VisualBasic.Right("000000" & (startBox + i_loopBox), 7))) = 1 Then
                                                                            statusSave = False
                                                                            MsgstatusSave = "BOX ALREADY"
                                                                            Exit For
                                                                        End If
                                                                    End If

                                                                    clsTmpDB.insertDetailDOBoxEX(tmp, ls_boxno + Microsoft.VisualBasic.Right("000000" & (startBox + i_loopBox), 7), ls_SeqNo, SQLCom)

                                                                    'UPDATE SURAT JALAN FORWARDER BILA GOOD RECEIVING SUDAH ADA (NO DN)
                                                                    If clsTmpDB.CekReceiveBox(tmp.SuratJalanNo, tmp.PartNo, tmp.PONo, ls_boxno + Microsoft.VisualBasic.Right("000000" & (startBox + i_loopBox), 7)) = 1 Then
                                                                        clsTmpDB.UpdateRECEX(tmp, SQLCom, ls_boxno + Microsoft.VisualBasic.Right("000000" & (startBox + i_loopBox), 7))
                                                                    End If

                                                                Next
                                                            End If

                                                            'CEK APAKAH SJ SUDAH ADA DI DN SUPPLIER
                                                            If clsTmpDB.CekData(tmp, "ReceiveForwarder_Master") = 1 Then
                                                                clsTmpDB.updateExcelDOEX(tmp, SQLCom)
                                                            End If

                                                        End If
                                                    End If
                                                Else
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    ElseIf templateCode = "REC-EX" Then 'Dian Export
                                        If statusSave = True Then
                                            tmp.PartNo = ExcelSheet.Range("D" & i).Value.ToString & ""
                                            tmp.PONo = tmp.PONo
                                            tmp.OrderNo = tmp.OrderNo

                                            If recStatus = True Then
                                                clsTmpDB.DeleteRecEX(tmp, SQLCom)
                                            End If

                                            'tmp.QtyBox = clsTmpDB.QtyBox(tmp.SupplierID, tmp.AffiliateID, tmp.PartNo)
                                            tmp.QtyBox = clsTmpDB.uf_GetQtybox(2, tmp.PONo, tmp.PartNo, tmp.SupplierID, tmp.AffiliateID)
                                            tmp.GoodRecQty = IIf(IsNumeric(ExcelSheet.Range("AL" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AL" & i).Value) * CDbl(tmp.QtyBox))
                                            tmp.DefectRecQty = IIf(IsNumeric(ExcelSheet.Range("AP" & i).Value) = False, 0, CDbl(ExcelSheet.Range("AP" & i).Value) * CDbl(tmp.QtyBox))
                                            tmp.UnitCls = clsTmpDB.UnitClsPart(tmp.PartNo)

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
                                                clsTmpDB.insertDetailReceivingEX(tmp, Xstatus, SQLCom)

                                                If IsNewRecEx = True Then
                                                    Dim i_rec1 As Integer = Microsoft.VisualBasic.Right(Trim(ExcelSheet.Range("R" & i).Value), 7)
                                                    Dim i_rec2 As Integer
                                                    Try
                                                        i_rec2 = Microsoft.VisualBasic.Right(Trim(ExcelSheet.Range("V" & i).Value), 7)
                                                    Catch ex As Exception
                                                        i_rec2 = i_rec1
                                                    End Try

                                                    Dim i_LabelNo As String = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("R" & i).Value), 2)
                                                    Dim i_LabelNo2 As String = ""
                                                    Dim L1 As String = i_LabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7)
                                                    Dim L2 As String = i_LabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec2), 7)

                                                    If ((i_rec2 - i_rec1) + 1) <> totBox Then
                                                        statusSave = False
                                                        MsgstatusSave = "QTY"
                                                        Exit For
                                                    End If

                                                    'Cek apakah Label terdaftar atau tidak
                                                    If clsTmpDB.CekBox2(tmp, L1, L2, Xstatus, totBox, SQLCom) = 1 Then
                                                        statusSave = False
                                                        MsgstatusSave = "ALREADY"
                                                        Exit For
                                                    End If
                                                    'Cek apakah Label terdaftar atau tidak

                                                    'CEK APAKAH BOX SESUAI DENGAN DN SUPPLIER
                                                    If statusDNSupp = False Then
                                                        If clsTmpDB.CekBoxDNSupplier(tmp, L1, L2, Xstatus, totBox, SQLCom) = 0 Then
                                                            statusSave = False
                                                            MsgstatusSave = "ALREADY"
                                                            Exit For
                                                        End If
                                                    End If

                                                    If clsTmpDB.insertDetailReceivingEX_BOX(tmp, L1, L2, Xstatus, totBox, SQLCom) = 0 Then
                                                        'Delete Data
                                                        statusSave = False
                                                        MsgstatusSave = "ALREADY"
                                                        Exit For
                                                    End If
                                                    clsTmpDB.RemainingReceiveExport(tmp, SQLCom)

                                                    For i_rec1 = i_rec1 To i_rec2
                                                        i_LabelNo2 = i_LabelNo & Microsoft.VisualBasic.Right(("0000000" & i_rec1), 7)
                                                        clsTmpDB.UpdateLabelPrint_RecEX(tmp, i_LabelNo2, Xstatus, SQLCom)
                                                    Next
                                                End If
                                            End If
                                            'EDIT FOR GOOD RECEIVE REPORT
                                            If statusDNSupp = True Then
                                                ExcelSheet.Range("R" & i & ":AP" & i).Interior.Color = Color.White
                                                ExcelSheet.Range("I" & i).Value = uf_GetPartName(tmp.PartNo)
                                            End If
                                        End If
                                    ElseIf templateCode = "INV-EX" Then 'DIAN EXPORT

                                        tmp.SuratJalanNo = ExcelSheet.Range("D" & i).Value.ToString.Replace(vbCr, "").Replace(vbLf, "") & ""
                                        tmp.PONo = ExcelSheet.Range("J" & i).Value.ToString.Replace(vbCr, "").Replace(vbLf, "") & ""
                                        tmp.OrderNo = ExcelSheet.Range("M" & i).Value.ToString.Replace(vbCr, "").Replace(vbLf, "") & ""
                                        tmp.PartNo = ExcelSheet.Range("Q" & i).Value.ToString.Replace(vbCr, "").Replace(vbLf, "") & ""
                                        tmp.DOQty = ExcelSheet.Range("AI" & i).Value
                                        tmp.InvoiceQty = ExcelSheet.Range("AM" & i).Value
                                        tmp.InvPrice = ExcelSheet.Range("AR" & i).Value
                                        tmp.InvAmount = ExcelSheet.Range("AV" & i).Value

                                        If ExcelSheet.Range("AP" & i).Value Is Nothing Then
                                            tmp.InvCurrCls = "03"
                                        Else
                                            tmp.InvCurrCls = clsTmpDB.CurrCls(ExcelSheet.Range("AP" & i).Value.ToString.Replace(vbCr, "").Replace(vbLf, "") & "")
                                        End If

                                        If inputMaster5 = False Then
                                            clsTmpDB.insertMasterInvoiceEx(tmp, SQLCom)
                                            inputMaster5 = True
                                        End If
                                        clsTmpDB.insertDetailInvoiceEx(tmp, SQLCom)
                                    ElseIf templateCode = "TALLY" Then 'DIAN Export
                                        Dim i_rec1 As Integer = Microsoft.VisualBasic.Right(Trim(ExcelSheet.Range("AE" & i).Value), 7)
                                        Dim i_rec2 As Integer

                                        Try
                                            i_rec2 = Microsoft.VisualBasic.Right(Trim(ExcelSheet.Range("AK" & i).Value), 7)
                                        Catch ex As Exception
                                            i_rec2 = i_rec1
                                        End Try

                                        Dim i_LabelNo As String = Microsoft.VisualBasic.Left(Trim(ExcelSheet.Range("AE" & i).Value), 2)
                                        Dim i_LabelNo2 As String = ""
                                        Dim L1 As String = i_LabelNo & Microsoft.VisualBasic.Right(("000000" & i_rec1), 7)
                                        Dim L2 As String = i_LabelNo & Microsoft.VisualBasic.Right(("000000" & i_rec2), 7)

                                        tmp.OrderNo = ExcelSheet.Range("I" & i).Value.ToString & ""
                                        tmp.PartNo = ExcelSheet.Range("O" & i).Value.ToString & ""
                                        tmp.BoxNo = ExcelSheet.Range("AE" & i).Value

                                        Try
                                            tmp.BoxNo2 = ExcelSheet.Range("AK" & i).Value
                                        Catch ex As Exception
                                            tmp.BoxNo2 = ExcelSheet.Range("AE" & i).Value
                                        End Try

                                        tmp.TotBoxEx = (i_rec2 - i_rec1) + 1

                                        Try
                                            tmp.PalletNo = IIf(IsNothing(ExcelSheet.Range("D" & i).Value.ToString & ""), "", ExcelSheet.Range("D" & i).Value).ToString & ""
                                        Catch ex As Exception
                                            tmp.PalletNo = tmp.PalletNo
                                        End Try

                                        If IIf(IsNothing(ExcelSheet.Range("AQ" & i).Value), "", ExcelSheet.Range("AQ" & i).Value).ToString & "" <> "" And IIf(IsNothing(ExcelSheet.Range("AQ" & i).Value), 0, ExcelSheet.Range("AQ" & i).Value) <> 0 Then tmp.Length = ExcelSheet.Range("AQ" & i).Value
                                        If IIf(IsNothing(ExcelSheet.Range("AT" & i).Value), "", ExcelSheet.Range("AT" & i).Value).ToString & "" <> "" And IIf(IsNothing(ExcelSheet.Range("AT" & i).Value), 0, ExcelSheet.Range("AT" & i).Value) <> 0 Then tmp.Width = ExcelSheet.Range("AT" & i).Value
                                        If IIf(IsNothing(ExcelSheet.Range("AW" & i).Value), "", ExcelSheet.Range("AW" & i).Value).ToString & "" <> "" And IIf(IsNothing(ExcelSheet.Range("AW" & i).Value), 0, ExcelSheet.Range("AW" & i).Value) <> 0 Then tmp.Height = ExcelSheet.Range("AW" & i).Value
                                        If IIf(IsNothing(ExcelSheet.Range("AZ" & i).Value), "", ExcelSheet.Range("AZ" & i).Value).ToString & "" <> "" And IIf(IsNothing(ExcelSheet.Range("AZ" & i).Value), 0, ExcelSheet.Range("AZ" & i).Value) <> 0 Then tmp.M3 = ExcelSheet.Range("AZ" & i).Value
                                        If IIf(IsNothing(ExcelSheet.Range("BC" & i).Value), "", ExcelSheet.Range("BC" & i).Value).ToString & "" <> "" And IIf(IsNothing(ExcelSheet.Range("BC" & i).Value), 0, ExcelSheet.Range("BC" & i).Value) <> 0 Then tmp.WeightPallet = ExcelSheet.Range("BC" & i).Value

                                        clsTmpDB.insertMasterTaily(tmp, SQLCom)
                                        clsTmpDB.UpdateShippingInstruction(tmp, SQLCom)

                                        i_LabelNo2 = i_LabelNo & Microsoft.VisualBasic.Right(("000000" & i_rec1), 7)
                                        clsTmpDB.insertDetailTaily(tmp, SQLCom)

                                    End If
NextProcess:
                                Next
                                If templateCode = "REC-EX" Then
                                    clsTmpDB.updateRemainingLblEpt(tmp, SQLCom)
                                End If
                            Next 'Next khusus untuk POEM

                            'Delete data PO Detail yang Qty = 0
                            If templateCode = "POEM" Then clsTmpDB.DeletePODetail0(tmp, SQLCom, statusApprove, "dbo.PO_Detail_Export")

                            'EDIT EXCEL FOR GOOD RECEIVE REPORT
                            If templateCode = "REC-EX" Then
                                If statusDNSupp = True And statusSave = True Then
                                    Dim ls_sql As String = ""
                                    ls_sql = " select distinct Consignee = isnull(MA.ConsigneeCode,''), attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), period = PME.Period, DOM.SuratJalanNo, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') + ' ' + isnull(MS.City,'') + ' ' + isnull(MS.Postalcode,'') SUPPAddress,  " & vbCrLf & _
                                             " PME.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  + ' ' + isnull(MF.City,'') + ' ' + isnull(MF.PostalCode,'') as FWDAddress, ETDVendor1 as ETDVendor, ETDPort1 as ETDPort, ETAPort1 as ETAPort, ETAFactory1 as ETAFactory, " & vbCrLf & _
                                             " PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'')  + ' ' + isnull(MA.City,'') + ' ' + isnull(MA.PostalCode,'') as AFFAddress, ISNULL(DOM.MovingList,0) MovingList, ConsigneName = isnull(MA.ConsigneeName,''), ConsigneeAdd = Rtrim(Isnull(MA.ConsigneeAddress,'')), isnull(DOM.ExcelCls,0) ExcelCls, ISNULL(DOM.SplitReffPONo, '') SplitReffPONo " & vbCrLf & _
                                             " from receiveforwarder_master DOM With(NOLOCK) LEFT JOIN PO_Master_Export PME With(NOLOCK) " & vbCrLf & _
                                             " ON  DOM.PONo = PME.PONo and DOM.SupplierID = PME.SupplierID and DOM.AffiliateID = PME.AffiliateID AND DOM.OrderNo = PME.OrderNo1" & vbCrLf & _
                                             " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                                             " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID " & vbCrLf & _
                                             " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                                             " WHERE DOM.PONo = '" & Trim(tmp.PONo) & "' " & vbCrLf & _
                                             " AND DOM.AffiliateID = '" & Trim(tmp.AffiliateID) & "' " & vbCrLf & _
                                             " AND DOM.SupplierID = '" & Trim(tmp.SupplierID) & "' " & vbCrLf & _
                                             " AND DOM.OrderNo = '" & Trim(tmp.OrderNo) & "' "
                                    Dim dsSplit As New DataSet
                                    dsSplit = uf_GetDataSet(ls_sql)

                                    ExcelSheet.Range("I11:X11").Value = Trim(dsSplit.Tables(0).Rows(0)("suppliername"))
                                    ExcelSheet.Range("I12:X15").Value = Trim(dsSplit.Tables(0).Rows(0)("SuppAddress"))
                                    ExcelSheet.Range("I19:X19").Value = Trim(dsSplit.Tables(0).Rows(0)("ForwarderName"))
                                    ExcelSheet.Range("I20:X22").Value = Trim(dsSplit.Tables(0).Rows(0)("FWDAddress"))
                                    ExcelSheet.Range("I23:X23").Value = "ATTN : " & Trim(dsSplit.Tables(0).Rows(0)("attn")) & "     TELP : " & Trim(dsSplit.Tables(0).Rows(0)("telp"))

                                    ExcelSheet.Range("AE15:AE15").Value = Trim(dsSplit.Tables(0).Rows(0)("PONo"))
                                    ExcelSheet.Range("AE11:AE11").Value = Format(dsSplit.Tables(0).Rows(0)("Period"), "yyyy-MM-dd")

                                    ExcelSheet.Range("AP11:AT11").Value = Format((dsSplit.Tables(0).Rows(0)("ETDVendor")), "yyyy-MM-dd")
                                    ExcelSheet.Range("AP13:AT13").Value = Format((dsSplit.Tables(0).Rows(0)("ETDPort")), "yyyy-MM-dd")
                                    ExcelSheet.Range("AP15:AT15").Value = Format((dsSplit.Tables(0).Rows(0)("ETAPort")), "yyyy-MM-dd")
                                    ExcelSheet.Range("AP17:AT17").Value = Format((dsSplit.Tables(0).Rows(0)("ETAFactory")), "yyyy-MM-dd")

                                    ExcelSheet.Range("AE19:AT19").Value = Trim(dsSplit.Tables(0).Rows(0)("Consignename"))
                                    ExcelSheet.Range("AE20:AT20").Value = Trim(dsSplit.Tables(0).Rows(0)("ConsigneeAdd"))

                                    ExcelSheet.Range("AK15").Value = "ETA PORT"
                                    ExcelSheet.Range("AK15").Font.Bold = True
                                    ExcelSheet.Range("AK17").Value = "ETA FACTORY"
                                    ExcelSheet.Range("AK117").Font.Bold = True

                                    ExcelSheet.Range("Z17").Value = ""
                                    ExcelSheet.Range("AE17").Value = ""

                                    ExcelSheet.Range("AE17").Interior.Color = Color.White
                                    ExcelSheet.Range("AE17:AI17").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                                    ExcelSheet.Range("AE17:AI17").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                                    ExcelSheet.Range("AE17:AI17").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                                    ExcelSheet.Range("AE17:AI17").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone

                                    Dim xlRange1 As Excel.Range = Nothing
                                    xlRange1 = CType(ExcelSheet.Rows("1:5"), Excel.Range)
                                    xlRange1.Delete()

                                    Dim xlRange2 As Excel.Range = Nothing
                                    xlRange2 = CType(ExcelSheet.Rows("20:25"), Excel.Range)
                                    xlRange2.Delete()

                                    ExcelSheet.Range("A2").Value = "GOOD RECEIVING REPORT"

                                    Dim NameExcel As String = ""
                                    'NameExcel = txtPathBackup.Text & "\" & "Good Receiving Report-" & Trim(tmp.AffiliateID) & "-" & Trim(tmp.SupplierID) & "-" & Trim(tmp.PONo) & " " & Format(Now, "ddMMyyyy hhmm")

                                    Dim filename As String = "\" & "Good Receiving Report-" & Trim(tmp.AffiliateID) & "-" & Trim(tmp.SupplierID) & "-" & Trim(tmp.PONo) & " " & Format(Now, "ddMMyyyy hhmm")
                                    Dim filecheck As String = txtPathBackup.Text + filename
                                    NameExcel = filecheck
                                    Dim ii As Integer = 0
                                    For u = 0 To 999999
                                        'If i <> 0 Then fileFix = filecheck + "(" + i.ToString() + ")"
                                        If System.IO.File.Exists(NameExcel + ".xlsm") Then
                                            ii = ii + 1
                                            NameExcel = filecheck + "(" + ii.ToString() + ")"
                                        Else
                                            Exit For
                                        End If
                                    Next

                                    If Not IsNothing(xlApp) Then
                                        ExcelBook.SaveAs(NameExcel)
                                    End If

                                    ls_file = NameExcel & ".xlsm"

                                Else
                                    If Not IsNothing(xlApp) Then
                                        ExcelBook.Save()
                                    End If
                                End If
                            Else
                                If Not IsNothing(xlApp) Then
                                    ExcelBook.Save()
                                End If
                            End If
                            If Not IsNothing(xlApp) Then
                                ExcelBook.Save()
                            End If

                            If statusSave = True Then
                                SQLTrans.Commit()
                            End If

                        End Using
                    Catch ex As Exception

                        'jika error kirim email
                        'up_SendEmailEX(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                        '                 tmp.OrderNo, tmp.OrderNo1, tmp.SuratJalanNo, ls_file, "", False)
                        'jika error kirim email

                        Application.DoEvents()
                        If msgInfo <> "" Then
                            gridProcess(Rtb1, 2, 0, "Failed " & msgInfo & vbCrLf, True)
                        Else
                            gridProcess(Rtb1, 2, 0, "Failed " & ex.Message.ToString & vbCrLf, True)
                        End If

                        Me.Cursor = Cursors.Default
                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            ExcelBook.Close()
                            'xlApp.Quit()
                        End If
                        'xlApp.Quit()
                        inputMaster = False : inputMaster2 = False : inputMaster3 = False
                        inputMaster4 = False

                        Application.DoEvents()
                        msgInfo = "move data corrupt process " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

                        'MOVE FILE
                        up_XMLFile_Copy(txtpath.Text & "\", tmpPathError & "\", excelName)
                        gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                        StatusKirim = False
                        Exit For

                    End Try

                    'msgInfo = "read file upload process..."
                    'gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)
                    'Next
suratJalanKeluar:
                    If pub_SuratJalanNo = False Then
                        Application.DoEvents()

                        If Not IsNothing(ExcelBook) Then
                            ExcelBook.Save()
                            ExcelBook.Close()
                            'xlApp.Quit()
                        End If

                        Application.DoEvents()
                        msgInfo = "move file corrupt process " & tmpAllNo
                        gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

                        'MOVE FILE
                        up_XMLFile_Copy(txtpath.Text & "\", tmpPathError & "\", excelName)

                        Application.DoEvents()
                        gridProcess(Rtb1, 1, 0, "1 move file corrupt", True)
                        gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                        StatusKirim = False
                        Exit For
                    End If

                Catch ex As Exception

                    Application.DoEvents()
                    If msgInfo <> "" Then
                        gridProcess(Rtb1, 2, 0, "Failed " & msgInfo & vbCrLf, True)
                    Else
                        gridProcess(Rtb1, 2, 0, "Failed " & ex.Message.ToString & vbCrLf, True)
                    End If

                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        ExcelBook.Close()
                        'xlApp.Quit()
                    End If

                    Application.DoEvents()
                    msgInfo = "move file corrupt process " & tmpAllNo
                    gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

                    'MOVE FILE
                    up_XMLFile_Copy(txtpath.Text & "\", tmpPathError & "\", excelName)

                    Application.DoEvents()
                    gridProcess(Rtb1, 1, 0, "1 move file corrupt", True)
                    gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                    StatusKirim = False
                    Exit For

                Finally
                    Me.Cursor = Cursors.Default

                End Try

            Next
            'Setelah semua Sheet Terbaca baru excel Close
            Application.DoEvents()
            msgInfo = "read file upload process..."
            gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

            Try
                If Not IsNothing(xlApp) Then
                    If Not IsNothing(ExcelBook) Then
                        ExcelBook.Save()
                        ExcelBook.Close()
                        xlApp.Quit()
                    End If
                End If
            Catch ex As Exception
                LogError(ex, excelName, "Ada Data yang Kosong", "")
            End Try

            'Lalu Kirim Email File nya dn Status kirim nya True
            'SEND NOTIFICATION
            Application.DoEvents()
            tmpAllNo = Replace(tmpAllNo, "In Sheet " & Jumlah, "")
            msgInfo = "send email process " & tmpAllNo
            gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

            'StatusKirim = False
            If StatusKirim = True Then
                If templateCode = "REC-EX" Then
                    If statusSave = False Then
                        up_SendEmailEX(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                          tmp.OrderNo, tmp.OrderNo, tmp.SuratJalanNo, ls_file, tmp.ForwarderID, False)
                    Else
                        up_SendEmailEX(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                          tmp.OrderNo, tmp.OrderNo, tmp.SuratJalanNo, ls_file, tmp.ForwarderID, True)
                    End If
                ElseIf templateCode = "DO-EX" Or templateCode = "POEM" Or templateCode = "POEE" Then
                    If statusSave = False Then
                        up_SendEmailEX(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                          tmp.OrderNo, tmp.OrderNo, tmp.SuratJalanNo, ls_file, "", False)
                    Else
                        up_SendEmailEX(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                          tmp.OrderNo, tmp.OrderNo, tmp.SuratJalanNo, ls_file, "", True)
                    End If
                ElseIf templateCode = "INV-EX" Then
                    If statusSave = False Then
                        up_SendEmailEX(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                          tmp.OrderNo, tmp.OrderNo, tmp.InvoiceNo, ls_file, tmp.ForwarderID, False)
                    Else
                        up_SendEmailEX(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                          tmp.OrderNo, tmp.OrderNo, tmp.InvoiceNo, ls_file, tmp.ForwarderID, True)
                    End If
                ElseIf templateCode = "INV" Then
                    up_SendEmail(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                       tmp.PONo, tmp.Period, tmp.AffiliateName, tmp.PORevNo, tmp.ShipCls, _
                       tmp.DeliveryLocation, tmp.KanbanDate, tmp.Remarks, tmp.CommercialCls, _
                       tmp.InvoiceNo, tmp.PartNo, tmp.KanbanNo, tmp.SuratJalanNo, ls_file, statusSave)
                Else
                    up_SendEmail(uf_GetConString, tmp.AffiliateID, tmp.SupplierID, templateCode, _
                       tmp.PONo, tmp.Period, tmp.AffiliateName, tmp.PORevNo, tmp.ShipCls, _
                       tmp.DeliveryLocation, tmp.KanbanDate, tmp.Remarks, tmp.CommercialCls, _
                       tmp.InvoiceNo, tmp.PartNo, tmp.KanbanNo, tmp.SuratJalanNo, ls_file, statusSave)
                End If
SenderrNotif:
                Application.DoEvents()
                gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                Application.DoEvents()
                msgInfo = "move file upload process " & tmpAllNo
                gridProcess(Rtb1, 1, 0, "Start " & msgInfo, True)

                'MOVE FILE
                up_XMLFile_Copy(txtpath.Text & "\", txtPathBackup.Text & "\", excelName)

                Application.DoEvents()
                gridProcess(Rtb1, 1, 0, "1 move file upload", True)
                gridProcess(Rtb1, 1, 0, "End " & msgInfo & vbCrLf, True)

                Me.Cursor = Cursors.Default

                inputMaster = False : inputMaster2 = False : inputMaster3 = False
                inputMaster4 = False : inputMaster5 = False
                statusApprove = ""
            Else
                gridProcess(Rtb1, 1, 0, "Not " & msgInfo & vbCrLf, True)
            End If

        Next

        Do While Marshal.ReleaseComObject(xlApp) > 0
        Loop

        If Not IsNothing(ExcelBook) Then
            Do While Marshal.ReleaseComObject(xlApp) > 0
            Loop
        End If

        If Not IsNothing(xlApp) Then
            xlApp = Nothing
        End If

        If Not IsNothing(xlApp) Then
            ExcelBook = Nothing
        End If

        GC.GetTotalMemory(False)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.GetTotalMemory(True)

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

            My.Computer.FileSystem.MoveFile(pPathSource & excelName, pPathDestination & excelName, True)
        Catch ex As Exception
            txtMsg.ForeColor = Color.Red
            txtMsg.Text = "Error move file"
        End Try
    End Sub

    Private Sub StartScheduler()
        Try
            txtMsg.Text = ""
            If Trim(btnAuto.Text) = "&Manual Upload" Then
                If (CDbl(txtTime.Text)) = 0 Then
                    tmrCycle1.Interval = 100
                Else
                    tmrCycle1.Interval = CDbl(txtTime.Text) * 1000 '1 menit
                End If
                'gs_StepProcess = "[" & Format(Now, "HH:mm:ss") & "] Scheduler Started" & vbCrLf
                gridProcess(Rtb1, 1, 0, gs_StepProcess)
                txtLast.Text = " -"
                timeLast = Format(Now, "yyyy-MM-dd HH:mm:ss")
                IntervalProcess = TimeSpan.FromSeconds(CDbl(txtTime.Text))
                up_ShowNextProcess()
                EnableForm(False)
                'CountProcess = 0

                'btnAuto.Text = "&Stop Upload"
                tmrCycle1.Enabled = True
                'Else
                '    btnAuto.Text = "&Start Upload"
                '    tmrCycle1.Enabled = False
                '    ActiveCycle = 0
                '    txtLast.Text = ""
                '    txtNext.Text = ""
                '    gs_StepProcess = "[" & Format(Now, "HH:mm:ss") & "] Scheduler Stopped" & vbCrLf
                '    gridProcess(Rtb1, 1, 0, gs_StepProcess)
                '    EnableForm(True)
            End If
        Catch ex As Exception
            txtMsg.ForeColor = Color.Red
            txtMsg.Text = "Error scheduler"
        End Try
    End Sub

    Private Sub EnableForm(ByVal Enable As Boolean)
        btnManual.Enabled = Enable
        BtnBrowse.Enabled = Enable
        txtpath.Enabled = Enable
        txtTime.Enabled = Enable
        'btnExit.Enabled = Enable
        txtPathBackup.Enabled = Enable
        Button1.Enabled = Enable
    End Sub

    Private Sub up_ShowNextProcess()
        JamServer = Format(Now, "yyyy-MM-dd HH:mm:ss")

        Dim Last As Date = FormatDateTime(timeLast)
        txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + IntervalProcess, "HH:mm:ss")
        NextProcess = txtNext.Text
        tmrCycle1.Enabled = True
    End Sub

    Private Function up_SendEmailEX(ByVal pConstr As String, ByVal pAffiliateID As String, ByVal pSupplierID As String, _
                                 ByVal pTemplateCode As String, ByVal pPONo As String, ByVal pOrderNo As String, _
                                 ByVal pSuratJalanNo As String, ByVal pFileName As String, ByVal pFwd As String, ByVal sts As Boolean)

        Dim i As Integer, kirimKe As Integer = 0
        Dim pURLPASI As String = "", pURLAffiliate As String = "", pURLSupplier As String = ""
        Dim kirimEmail As Integer = 0
        Dim pPASI As Boolean = False, pAffiliate As Boolean = False, pSupplier As Boolean = False
        Dim EmailSetting As clsEmail, EmailPASI As clsEmail, EmailForwarder As clsEmail, EmailSupplier As clsEmail, SubjectBody As clsEmail
        Dim ReceipientEmail As String, ReceipientArray() As String
        Dim CCEmail As String, CCArray() As String, retMessage As String = ""
        Dim pReceipient As String = "", pCC As String = ""
        Dim ReceipientList As New List(Of String)
        Dim CCList As New List(Of String)
        Dim a As String = ""
        Try
            Me.Cursor = Cursors.WaitCursor

            EmailSetting = clsEmailDB.GetEmailSettingEx(pConstr)
            If EmailSetting Is Nothing Then
                Return ""
            End If
            EmailPASI = clsEmailDB.GetEmailPASIEx(pConstr, pTemplateCode)
            If EmailPASI Is Nothing Then
                Return ""
            End If
            If pTemplateCode = "REC-EX" Or pTemplateCode = "Tally" Then
                EmailForwarder = clsEmailDB.GetEmailFWD(pConstr, pFwd, pTemplateCode)
                If EmailForwarder Is Nothing Then
                    Return ""
                End If

                EmailSupplier = clsEmailDB.GetEmailSupplierEx(pConstr, pSupplierID, pTemplateCode)
                If EmailSupplier Is Nothing Then
                    Return ""
                End If
            Else
                EmailSupplier = clsEmailDB.GetEmailSupplierEx(pConstr, pSupplierID, pTemplateCode)
                If EmailSupplier Is Nothing Then
                    Return ""
                End If
            End If

            If templateCode = "REC-EX" Then
                SubjectBody = clsEmailDB.GetEmailSubjectBodyEx(pConstr, pTemplateCode, pURLPASI, pSuratJalanNo, sts)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailForwarder.EmailFWDTo
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                CCEmail = EmailForwarder.EmailFWDCC & ";" & EmailPASI.EmailPASICC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                If sts = False Then
                    retMessage = SendEmailEx(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                         SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, pFileName)
                Else
                    retMessage = SendEmailEx(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

                    'SEND EMAIL TO SUPPLIER IF NO DN
                    If statusDNSupp = True Then
                        SubjectBody = clsEmailDB.GetEmailSubjectBodyEx(pConstr, pTemplateCode, pURLPASI, pOrderNo, sts)
                        If SubjectBody Is Nothing Then
                            Return ""
                        End If
                        ReceipientEmail = EmailSupplier.EmailSupplierTo
                        ReceipientArray = Split(ReceipientEmail, ";")
                        For i = 0 To UBound(ReceipientArray)
                            If ReceipientArray(i) <> "" Then
                                ReceipientList.Add(ReceipientArray(i))
                            End If
                        Next
                        CCEmail = EmailSupplier.EmailSupplierCC & ";" & EmailPASI.EmailPASICC
                        CCArray = Split(CCEmail, ";")
                        For i = 0 To UBound(CCArray)
                            If CCArray(i) <> "" Then
                                CCList.Add(CCArray(i))
                            End If
                        Next

                        retMessage = SendEmailEx(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                         SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, pFileName)
                    End If
                End If
                ReceipientList.Clear()
                CCList.Clear()
            ElseIf templateCode = "DO-EX" Then
                SubjectBody = clsEmailDB.GetEmailSubjectBodyEx(pConstr, pTemplateCode, pURLPASI, pSuratJalanNo, sts)
                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailSupplier.EmailSupplierTo
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                CCEmail = EmailSupplier.EmailSupplierCC & ";" & EmailPASI.EmailPASICC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                If MsgstatusSave = "" Then
                    retMessage = SendEmailEx(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                         SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")
                Else
                    retMessage = SendEmailEx(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                         SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, pFileName)
                End If
                ReceipientList.Clear()
                CCList.Clear()
            Else 'POEM, POEX
                If templateCode = "INV-EX" Then
                    SubjectBody = clsEmailDB.GetEmailSubjectBodyEx(pConstr, pTemplateCode, pURLPASI, pSuratJalanNo, sts)
                Else
                    SubjectBody = clsEmailDB.GetEmailSubjectBodyEx(pConstr, pTemplateCode, pURLPASI, pOrderNo, sts)
                End If

                If SubjectBody Is Nothing Then
                    Return ""
                End If
                ReceipientEmail = EmailSupplier.EmailSupplierTo
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                CCEmail = EmailSupplier.EmailSupplierCC & ";" & EmailPASI.EmailPASICC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next

                retMessage = SendEmailEx(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                        SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

                ReceipientList.Clear()
                CCList.Clear()
            End If
            If retMessage <> "" Then LogErrorWithoutException(retMessage, pFileName, pSupplierID, a)

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            'txtMsg.Text = "Error send Email"
            LogError(ex, pFileName, pSupplierID, a)
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Function up_SendEmail(ByVal pConstr As String, ByVal pAffiliateID As String, ByVal pSupplierID As String, _
                                  ByVal pTemplateCode As String, ByVal pPONo As String, ByVal pPeriod As String, _
                                  ByVal pAffiliateName As String, ByVal pPONoRev As String, ByVal pShip As String, _
                                  ByVal pDeliveryLoc As String, ByVal pKanbanDate As Date, ByVal pRemarks As String, _
                                  ByVal pCommercialCls As String, ByVal pInvoiceNo As String, ByVal pPartNo As String, _
                                  ByVal pKanbanNo As String, ByVal pSuratJalanNo As String, Optional ByVal pFileName As String = "", Optional ByVal sts As Boolean = True)

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
        Dim a As String = ""

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
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                CCEmail = EmailPASI.EmailPASICC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next

                retMessage = clsSendEmail.SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

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
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                CCEmail = EmailSupplier.EmailSupplierCC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next

                retMessage = clsSendEmail.SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

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
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                CCEmail = EmailPASI.EmailPASICC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next

                retMessage = clsSendEmail.SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

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
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                CCEmail = EmailSupplier.EmailSupplierCC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next

                retMessage = clsSendEmail.SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

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
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                CCEmail = EmailAffiliate.EmailAffiliateTo & ";" & EmailAffiliate.EmailAffiliateCC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next

                retMessage = clsSendEmail.SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

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
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                CCEmail = EmailPASI.EmailPASICC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next

                retMessage = clsSendEmail.SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

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
                ReceipientArray = Split(ReceipientEmail, ";")
                For i = 0 To UBound(ReceipientArray)
                    If ReceipientArray(i) <> "" Then
                        ReceipientList.Add(ReceipientArray(i))
                    End If
                Next
                a = String.Join(",", ReceipientList)
                CCEmail = EmailSupplier.EmailSupplierCC
                CCArray = Split(CCEmail, ";")
                For i = 0 To UBound(CCArray)
                    If CCArray(i) <> "" Then
                        CCList.Add(CCArray(i))
                    End If
                Next

                retMessage = clsSendEmail.SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                     SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, "")

                ReceipientList.Clear()
                CCList.Clear()

            ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then
                If MsgstatusSave <> "" Then
                    '=========================== EMAIL TO ALL ===========================
                    'DO TIDAK PAKAI URL HANYA NOTIFIKASI
                    pURLPASI = ""

                    SubjectBody = clsEmailDB.GetEmailSubjectBody(pConstr, pTemplateCode, pURLPASI, pSuratJalanNo, sts)
                    If SubjectBody Is Nothing Then
                        Return ""
                    End If
                    ReceipientEmail = EmailSupplier.EmailSupplierTo
                    ReceipientArray = Split(ReceipientEmail, ";")
                    For i = 0 To UBound(ReceipientArray)
                        If ReceipientArray(i) <> "" Then
                            ReceipientList.Add(ReceipientArray(i))
                        End If
                    Next
                    a = String.Join(",", ReceipientList)
                    CCEmail = EmailPASI.EmailPASICC
                    CCArray = Split(CCEmail, ";")
                    For i = 0 To UBound(CCArray)
                        If CCArray(i) <> "" Then
                            CCList.Add(CCArray(i))
                        End If
                    Next

                    If MsgstatusSave = "OVER" Then
                        retMessage = clsSendEmail.SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                         SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port, pFileName)
                    Else
                        retMessage = SendEmail(ReceipientList, CCList, EmailPASI.EmailPASITo, SubjectBody.Subject, _
                             SubjectBody.Body, EmailSetting.SenderName, EmailSetting.Password, EmailSetting.EnableSSL, EmailSetting.DefaultCredentials, EmailSetting.SMTPServer, EmailSetting.Port)
                    End If

                    If retMessage <> "" Then LogErrorWithoutException(retMessage, pFileName, pSupplierID, a)
                    ReceipientList.Clear()
                    CCList.Clear()
                End If
            End If

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            LogError(ex, pFileName, pSupplierID, a)
            'txtMsg.Text = "Error send Email"
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Function SendEmail(ByVal Recipients As List(Of String), _
                        ByVal CC As List(Of String), _
                        ByVal FromAddress As String, _
                        ByVal Subject As String, _
                        ByVal Body As String, _
                        ByVal UserName As String, _
                        ByVal Password As String, _
                        ByVal EnableSSL As Boolean, _
                        ByVal DefaultCredentials As Boolean, _
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

            'SMTPServer.Host = Server '"tos-is.com" '
            SMTPServer.Host = Trim(Server)
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            SMTPServer.UseDefaultCredentials = DefaultCredentials
            SMTPServer.EnableSsl = EnableSSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(UserName), Trim(Password))
            SMTPServer.Credentials = myCredential

            'If SMTPServer.UseDefaultCredentials = True Then
            '    SMTPServer.EnableSsl = True
            'Else
            '    SMTPServer.Credentials = New System.Net.NetworkCredential(UserName, Password)
            'End If
            'SMTPServer.Credentials = New System.Net.NetworkCredential(UserName, Password)
            'SMTPServer.EnableSsl = EnableSSL

            SMTPServer.Port = Port
            SMTPServer.Send(Email)
            Email.Dispose()
            Return ""

        Catch ex As SmtpException
            Email.Dispose()
            retMsg = "Sending Email Failed. Smtp Error." & " - " & ex.Message
        Catch ex As ArgumentOutOfRangeException
            Email.Dispose()
            retMsg = "Sending Email Failed. Check Port Number." & " - " & ex.Message
        Catch Ex As InvalidOperationException
            Email.Dispose()
            retMsg = "Sending Email Failed. Check Port Number." & " - " & Ex.Message
        End Try
        Return retMsg

    End Function

    Function SendEmailEx(ByVal Recipients As List(Of String), _
                        ByVal CC As List(Of String), _
                        ByVal FromAddress As String, _
                        ByVal Subject As String, _
                        ByVal Body As String, _
                        ByVal UserName As String, _
                        ByVal Password As String, _
                        ByVal EnableSSL As Boolean, _
                        ByVal DefaultCredentials As Boolean, _
                        ByVal Server As String, _
                        ByVal Port As Integer, ByVal pFile As String) As String

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

            Dim filename As String = pFile
            If pFile <> "" Then Email.Attachments.Add(New Attachment(filename))

            'SMTPServer.Host = Trim(Server) '"tos-is.com" '
            SMTPServer.Host = Trim(Server)
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            SMTPServer.UseDefaultCredentials = DefaultCredentials
            SMTPServer.EnableSsl = EnableSSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(UserName), Trim(Password))
            SMTPServer.Credentials = myCredential

            SMTPServer.Port = Trim(Port)
            'SMTPServer.Credentials = New System.Net.NetworkCredential(Trim(UserName), Trim(Password))
            'SMTPServer.EnableSsl = EnableSSL

            SMTPServer.Send(Email)
            Email.Dispose()
            Return ""

        Catch ex As SmtpException
            Email.Dispose()
            retMsg = "Sending Email Failed. Smtp Error." & " - " & ex.Message
        Catch ex As ArgumentOutOfRangeException
            Email.Dispose()
            retMsg = "Sending Email Failed. Check Port Number." & " - " & ex.Message
        Catch Ex As InvalidOperationException
            Email.Dispose()
            retMsg = "Sending Email Failed. Check Port Number." & " - " & Ex.Message
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

    Private Function uf_GetPartName(ByVal pPartNo As String)
        Dim ls_SQL As String = ""
        MdlConn.ReadConnection()
        ls_SQL = " Select * From MS_Parts WITH(NOLOCK) Where PartNo = '" & pPartNo & "'"

        Dim ds As New DataSet
        ds = uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return Trim(ds.Tables(0).Rows(0)("PartName"))
        End If
    End Function
#End Region

#Region "Control Event"
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManual.Click
        UploadData()
    End Sub


    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub BtnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBrowse.Click
        Dim result As DialogResult

        result = fbd.ShowDialog()
        If result = DialogResult.Cancel Then
            Exit Sub
        Else
            txtpath.Text = fbd.SelectedPath
        End If
    End Sub

    Private Sub btnAuto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuto.Click
        Dim msgInfo As String = ""
        Me.Cursor = Cursors.WaitCursor
        ProcessStarted = True

        Try
            tmrCycle1.Enabled = False
            UploadData()
            tmrCycle1.Enabled = True
        Catch ex As Exception
            Application.DoEvents()
            If msgInfo <> "" Then
                gridProcess(Rtb1, 2, 0, "Failed " & msgInfo & vbCrLf, True)
            Else
                gridProcess(Rtb1, 2, 0, "Failed " & ex.Message.ToString & vbCrLf, True)
            End If
            txtMsg.Text = ex.Message.ToString
        Finally
            Me.Cursor = Cursors.Default

            IntervalProcess = TimeSpan.FromSeconds(CDbl(txtTime.Text))
            timeLast = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            txtLast.Text = Format(timeLast, "yyyy-MM-dd HH:mm:ss")
            Dim Last As Date = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + IntervalProcess, "HH:mm:ss")

            ProcessStarted = False
            IsProcessing = False
        End Try
    End Sub

    Private Sub tmrCycle1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrCycle1.Tick
        Dim Last As Date

        If Format(Now, "yyyy-MM-dd HH:mm:ss") >= NextProcess And ProcessStarted = False And IsProcessing = False Then
            'If clsTmpDB.BatchProcessStatus = "1" Then
            Me.Cursor = Cursors.WaitCursor
            ProcessStarted = True
            tmrCycle1.Enabled = False
            UploadData()
            tmrCycle1.Enabled = True
            Me.Cursor = Cursors.Default

            IntervalProcess = TimeSpan.FromSeconds(CDbl(txtTime.Text))
            timeLast = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            txtLast.Text = Format(timeLast, "yyyy-MM-dd HH:mm:ss")
            Last = FormatDateTime(txtLast.Text)
            txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + IntervalProcess, "HH:mm:ss")

            ProcessStarted = False
            IsProcessing = False

            'clsTmpDB.BatchProcessStatusUpdate()
            'Else
            'gridProcess(Rtb1, 1, 0, "Skip process to wait other process finished", True)
            'End If

            'IntervalProcess = TimeSpan.FromSeconds(CDbl(txtTime.Text))
            'timeLast = Format(Now, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            'txtLast.Text = Format(timeLast, "yyyy-MM-dd HH:mm:ss")
            'Last = FormatDateTime(txtLast.Text)
            'txtNext.Text = Format(Now, "yyyy-MM-dd") & " " & Format(Last + IntervalProcess, "HH:mm:ss")
        End If
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim result As DialogResult

        result = fbd.ShowDialog()
        If result = DialogResult.Cancel Then
            Exit Sub
        Else
            txtPathBackup.Text = fbd.SelectedPath
        End If
    End Sub

    Private Sub LogError(ex As Exception, File As String, pSupplierID As String, pAdress As String)
        Try
            Dim message As String = String.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
            message += Environment.NewLine
            message += "-----------------------------------------------------------"
            message += Environment.NewLine
            message += "Adress In [] " + pAdress
            message += Environment.NewLine
            message += pSupplierID
            message += Environment.NewLine
            message += String.Format("Message: {0}", ex.Message)
            message += Environment.NewLine
            message += String.Format("StackTrace: {0}", ex.StackTrace)
            message += Environment.NewLine
            message += String.Format("Source: {0}", ex.Source)
            message += Environment.NewLine
            message += String.Format("TargetSite: {0}", ex.TargetSite.ToString())
            message += Environment.NewLine
            message += "Error In File " + File + ""
            message += Environment.NewLine
            message += "-----------------------------------------------------------"
            message += Environment.NewLine

            Dim path As String = "D:\PASI EBWEB\Log\Error\AttachmentError_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".txt"
            If System.IO.File.Exists(path) = False Then
                System.IO.File.Create(path).Dispose()
            End If
            Dim Writer As New System.IO.StreamWriter(path, True)
            Writer.WriteLine(message)
            Writer.Close()
        Catch exx As Exception

        End Try
    End Sub

    Private Sub LogErrorWithoutException(ex As String, File As String, pSupplierID As String, pAdress As String)
        Try
            Dim message As String = String.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
            message += Environment.NewLine
            message += "-----------------------------------------------------------"
            message += Environment.NewLine
            message += "Adress In [] " + pAdress
            message += Environment.NewLine
            message += pSupplierID
            message += Environment.NewLine
            message += String.Format("Message: {0}", ex)
            'message += Environment.NewLine
            'message += String.Format("StackTrace: {0}", ex.StackTrace)
            'message += Environment.NewLine
            'message += String.Format("Source: {0}", ex.Source)
            'message += Environment.NewLine
            'message += String.Format("TargetSite: {0}", ex.TargetSite.ToString())
            message += Environment.NewLine
            message += "Error In File " + File + ""
            message += Environment.NewLine
            message += "-----------------------------------------------------------"
            message += Environment.NewLine

            Dim path As String = "D:\PASI EBWEB\Log\Error\AttachmentError_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".txt"
            If System.IO.File.Exists(path) = False Then
                System.IO.File.Create(path).Dispose()
            End If
            Dim Writer As New System.IO.StreamWriter(path, True)
            Writer.WriteLine(message)
            Writer.Close()
        Catch exx As Exception

        End Try
    End Sub
#End Region

End Class
