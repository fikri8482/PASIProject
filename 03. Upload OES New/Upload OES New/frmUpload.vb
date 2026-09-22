Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.VisualBasic

Public Class frmUpload
    Dim strHeaderID As String = ""
    Dim strDataID As String = ""
    Dim strAffiliateCode As String = ""
    Dim strPASICode As String = ""
    Dim strOrderType As String = ""
    Dim strOrderRequest As String = ""
    Dim strCommercialCls As String = ""
    Dim strFreight As String = ""
    Dim strOrderNo1 As String = ""
    Dim strETDVendor1 As String = ""
    Dim strETDPort1 As String = ""
    Dim strETAPort1 As String = ""
    Dim strETAFactory1 As String = ""
    Dim strDeliveryMonth1 As String = ""
    Dim strOrderNo2 As String = ""
    Dim strETDVendor2 As String = ""
    Dim strETDPort2 As String = ""
    Dim strETAPort2 As String = ""
    Dim strETAFactory2 As String = ""
    Dim strDeliveryMonth2 As String = ""
    Dim strOrderNo3 As String = ""
    Dim strETDVendor3 As String = ""
    Dim strETDPort3 As String = ""
    Dim strETAPort3 As String = ""
    Dim strETAFactory3 As String = ""
    Dim strDeliveryMonth3 As String = ""
    Dim strOrderNo4 As String = ""
    Dim strETDVendor4 As String = ""
    Dim strETDPort4 As String = ""
    Dim strETAPort4 As String = ""
    Dim strETAFactory4 As String = ""
    Dim strDeliveryMonth4 As String = ""
    Dim strOrderNo5 As String = ""
    Dim strETDVendor5 As String = ""
    Dim strETDPort5 As String = ""
    Dim strETAPort5 As String = ""
    Dim strETAFactory5 As String = ""
    Dim strDeliveryMonth5 As String = ""
    Dim strForecastDate1 As String = ""
    Dim strForecastDate2 As String = ""
    Dim strForecastDate3 As String = ""
    Dim strForecastDate4 As String = ""
    Dim strForecastDate5 As String = ""
    Dim strForecastDate6 As String = ""
    Dim strForecastDate7 As String = ""
    Dim strForecastDate8 As String = ""
    Dim strPartNo As String = ""
    Dim strUOM As String = ""
    Dim strOrderQty1 As String = ""
    Dim strOrderQty2 As String = ""
    Dim strOrderQty3 As String = ""
    Dim strOrderQty4 As String = ""
    Dim strOrderQty5 As String = ""
    Dim strPartDescription As String = ""
    Dim strForecastQty1 As String = ""
    Dim strForecastQty2 As String = ""
    Dim strForecastQty3 As String = ""
    Dim strForecastQty4 As String = ""
    Dim strForecastQty5 As String = ""
    Dim strForecastQty6 As String = ""
    Dim strForecastQty7 As String = ""
    Dim strForecastQty8 As String = ""

    Private Sub UpdateLog(ByVal strMessage As String)
        If txtLog.Text = "" Then
            txtLog.Text = Format(Date.Now, "[yyyy-MM-dd HH:mm:ss] ") & strMessage
        Else
            txtLog.Text = txtLog.Text & vbCrLf & Format(Date.Now, "[yyyy-MM-dd HH:mm:ss] ") & strMessage
        End If

        txtLog.SelectionStart = txtLog.TextLength
        txtLog.Focus()

    End Sub

    Private Sub UploadProcess()
        Dim strFile As String
        Dim dtStartProcessNow As DateTime = DateTime.Now
        Dim logFolder As String = "C:\EDI\OES\Log"
        Dim AdaProcess As Boolean = False

        Try
            strFile = Dir(gs_UploadPath & "\*.*")
            If strFile <> "" Then
                AdaProcess = True
                UpdateLog("Start processing file : " & strFile)

                up_Import(gs_UploadPath & "\" & strFile)

                UpdateLog("End processing file : " & strFile & vbCrLf)
            End If
        Catch ex As Exception
            UpdateLog(ex.Message)
        Finally
            If (AdaProcess) Then
                'Buat folder kalau belum ada
                If Not Directory.Exists(logFolder) Then
                    Directory.CreateDirectory(logFolder)
                End If

                Dim path_log As String = logFolder & "\Log_OES_" & dtStartProcessNow.ToString("yyyy-MM-dd") & ".txt"

                ' Menggunakan Microsoft Learn - WriteAllText untuk menulis konten
                File.WriteAllText(path_log, txtLog.Text)
            End If
        End Try
    End Sub

    Private Sub ClearH10()
        strOrderType = ""
        strOrderRequest = ""
        strCommercialCls = ""
        strFreight = ""
        strOrderNo1 = ""
        strETDVendor1 = ""
        strETDPort1 = ""
        strETAPort1 = ""
        strETAFactory1 = ""
        strDeliveryMonth1 = ""
        strOrderNo2 = ""
        strETDVendor2 = ""
        strETDPort2 = ""
        strETAPort2 = ""
        strETAFactory2 = ""
        strDeliveryMonth2 = ""
        strOrderNo3 = ""
        strETDVendor3 = ""
        strETDPort3 = ""
        strETAPort3 = ""
        strETAFactory3 = ""
        strDeliveryMonth3 = ""
        strOrderNo4 = ""
        strETDVendor4 = ""
        strETDPort4 = ""
        strETAPort4 = ""
        strETAFactory4 = ""
        strDeliveryMonth4 = ""
        strOrderNo5 = ""
        strETDVendor5 = ""
        strETDPort5 = ""
        strETAPort5 = ""
        strETAFactory5 = ""
        strDeliveryMonth5 = ""
        strForecastDate1 = ""
        strForecastDate2 = ""
        strForecastDate3 = ""
        strForecastDate4 = ""
        strForecastDate5 = ""
        strForecastDate6 = ""
        strForecastDate7 = ""
        strForecastDate8 = ""
        strPartNo = ""
        strUOM = ""
        strOrderQty1 = ""
        strOrderQty2 = ""
        strOrderQty3 = ""
        strOrderQty4 = ""
        strOrderQty5 = ""
        strPartDescription = ""
        strForecastQty1 = ""
        strForecastQty2 = ""
        strForecastQty3 = ""
        strForecastQty4 = ""
        strForecastQty5 = ""
        strForecastQty6 = ""
        strForecastQty7 = ""
        strForecastQty8 = ""
    End Sub

    Private Sub ClearH20()
        strOrderNo1 = ""
        strETDVendor1 = ""
        strETDPort1 = ""
        strETAPort1 = ""
        strETAFactory1 = ""
        strDeliveryMonth1 = ""
        strOrderNo2 = ""
        strETDVendor2 = ""
        strETDPort2 = ""
        strETAPort2 = ""
        strETAFactory2 = ""
        strDeliveryMonth2 = ""
        strOrderNo3 = ""
        strETDVendor3 = ""
        strETDPort3 = ""
        strETAPort3 = ""
        strETAFactory3 = ""
        strDeliveryMonth3 = ""
        strOrderNo4 = ""
        strETDVendor4 = ""
        strETDPort4 = ""
        strETAPort4 = ""
        strETAFactory4 = ""
        strDeliveryMonth4 = ""
        strOrderNo5 = ""
        strETDVendor5 = ""
        strETDPort5 = ""
        strETAPort5 = ""
        strETAFactory5 = ""
        strDeliveryMonth5 = ""
        strForecastDate1 = ""
        strForecastDate2 = ""
        strForecastDate3 = ""
        strForecastDate4 = ""
        strForecastDate5 = ""
        strForecastDate6 = ""
        strForecastDate7 = ""
        strForecastDate8 = ""
        strPartNo = ""
        strUOM = ""
        strOrderQty1 = ""
        strOrderQty2 = ""
        strOrderQty3 = ""
        strOrderQty4 = ""
        strOrderQty5 = ""
        strPartDescription = ""
        strForecastQty1 = ""
        strForecastQty2 = ""
        strForecastQty3 = ""
        strForecastQty4 = ""
        strForecastQty5 = ""
        strForecastQty6 = ""
        strForecastQty7 = ""
        strForecastQty8 = ""
    End Sub

    Private Sub ClearD10()
        strPartNo = ""
        strUOM = ""
        strOrderQty1 = "0"
        strOrderQty2 = "0"
        strOrderQty3 = "0"
        strOrderQty4 = "0"
        strOrderQty5 = "0"
        strPartDescription = ""
        strForecastQty1 = "0"
        strForecastQty2 = "0"
        strForecastQty3 = "0"
        strForecastQty4 = "0"
        strForecastQty5 = "0"
        strForecastQty6 = "0"
        strForecastQty7 = "0"
        strForecastQty8 = "0"
    End Sub

    Private Sub up_Import(ByVal pFile As String)
        Dim strMessage As String = ""
        Dim strSQL As String = ""
        Dim ds As DataSet = Nothing

        Dim strAffiliateID As String = ""
        Dim strForwarderID As String = ""
        Dim strSupplierID As String = ""

        Try

            Dim readText() As String = File.ReadAllLines(pFile)

            Using sqlConn As New SqlConnection(gs_ConString)
                sqlConn.Open()

                Using sr As StreamReader = New StreamReader(pFile)
                    Dim strText As String = ""
                    Dim intLineEnd = readText.Length

                    Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("UploadOES")
                        For intLine = 1 To intLineEnd
                            strText = sr.ReadLine

                            strHeaderID = Mid(strText, 1, 3).Trim
                            Select Case strHeaderID
                                Case "H00"
                                    strDataID = Mid(strText, 4, 8).Trim
                                    strAffiliateCode = Mid(strText, 12, 8).Trim
                                    strPASICode = Mid(strText, 20, 8).Trim

                                    If strDataID <> "VF30" Then
                                        strMessage = "Invalid Data ID. Please check template, Data ID should be 'VF30'" & Environment.NewLine & _
                                                    "File Name : " & Dir(pFile) & Environment.NewLine & _
                                                    "Line : " & intLine & Environment.NewLine & _
                                                    "Data : " & strDataID
                                        Exit For
                                    End If

                                    strSQL = "SELECT AffiliateID FROM MS_Affiliate WHERE AffiliateCode = '" & strAffiliateCode & "' "
                                    ds = uf_GetDataSet(strSQL)
                                    If ds.Tables(0).Rows.Count > 0 Then
                                        strAffiliateID = ds.Tables(0).Rows("0")("AffiliateID").ToString.Trim
                                    Else
                                        strAffiliateID = ""
                                    End If
                                    ds.Clear()

                                    If strAffiliateID = "" Then
                                        strMessage = "Invalid affiliate Code. Please check setting at Affiliate Master" & Environment.NewLine & _
                                                    "File Name : " & Dir(pFile) & Environment.NewLine & _
                                                    "Line : " & intLine & Environment.NewLine & _
                                                    "Data : " & strAffiliateCode
                                        Exit For
                                    End If

                                    If strPASICode <> "32G8" Then
                                        strMessage = "Invalid PASI Code. Please check template, PASI Code should be '32G8'" & Environment.NewLine & _
                                                    "File Name : " & Dir(pFile) & Environment.NewLine & _
                                                    "Line : " & intLine & Environment.NewLine & _
                                                    "Data : " & strPASICode
                                        Exit For
                                    End If
                                Case "H10"
                                    ClearH10()

                                    strOrderType = Mid(strText, 4, 2).Trim
                                    strOrderRequest = Mid(strText, 6, 1).Trim
                                    strCommercialCls = Mid(strText, 7, 1).Trim
                                    strFreight = Mid(strText, 8, 1).Trim

                                    If strOrderType <> "MM" And strOrderType <> "MW" And strOrderType <> "WW" Then
                                        strMessage = "Invalid Order Type. Please check template, Order Type should be 'MM', 'MW' or 'WW'" & Environment.NewLine & _
                                                    "File Name : " & Dir(pFile) & Environment.NewLine & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "Data : " & strOrderType
                                        Exit For
                                    End If

                                    If strOrderRequest <> "M" And strOrderRequest <> "E" Then
                                        strMessage = "Invalid Order Request. Please check template, Order Request should be 'M' or 'E'" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "Data : " & strOrderRequest
                                        Exit For
                                    End If

                                    If strCommercialCls <> "C" And strCommercialCls <> "N" Then
                                        strMessage = "Invalid Commercial Cls. Please check template, Commercial Cls should be 'C' or 'N'" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "Data : " & strCommercialCls
                                        Exit For
                                    End If

                                    If strFreight <> "A" And strFreight <> "B" And strFreight <> "T" And strFreight <> "H" Then
                                        strMessage = "Invalid Freight. Please check template, Freight should be 'A', 'B', 'T' or 'H'" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "Data : " & strFreight
                                        Exit For
                                    End If

                                    strSQL = "SELECT ForwarderID FROM MS_ForwarderMapping WHERE AffiliateID = '" & strAffiliateID & "' and ShipCls = '" & strFreight & "'"
                                    ds = uf_GetDataSet(strSQL)
                                    If ds.Tables(0).Rows.Count > 0 Then
                                        strForwarderID = ds.Tables(0).Rows("0")("ForwarderID").ToString.Trim
                                    Else
                                        strForwarderID = ""
                                    End If
                                    ds.Clear()

                                    If strForwarderID = "" Then
                                        strSQL = "SELECT ForwarderID FROM MS_Forwarder WHERE DefaultCls = '1'"
                                        ds = uf_GetDataSet(strSQL)
                                        If ds.Tables(0).Rows.Count > 0 Then
                                            strForwarderID = ds.Tables(0).Rows("0")("ForwarderID").ToString.Trim
                                        Else
                                            strForwarderID = "SLI1"
                                        End If
                                        ds.Clear()
                                    End If

                                    If strForwarderID = "" Then
                                        strMessage = "Can't get Forwarder ID. Please check data Forwarder Mapping and Forwarder Default" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "Affiliate ID : " & strAffiliateID & vbCrLf & _
                                                    "Freight : " & strFreight
                                        Exit For
                                    End If
                                Case "H20"
                                    ClearH20()

                                    strOrderNo1 = Mid(strText, 4, 15).Trim
                                    strETDVendor1 = Mid(strText, 19, 8).Trim
                                    strETDPort1 = Mid(strText, 27, 8).Trim
                                    strETAPort1 = Mid(strText, 35, 8).Trim
                                    strETAFactory1 = Mid(strText, 43, 8).Trim
                                    strDeliveryMonth1 = Mid(strText, 51, 4).Trim

                                    If Not IsDate(Mid(strETDVendor1, 1, 4) & "-" & Mid(strETDVendor1, 5, 2) & "-" & Mid(strETDVendor1, 7, 2)) Then
                                        strMessage = "ETD Vendor is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Vendor : " & strETDVendor1
                                        Exit For
                                    Else
                                        strETDVendor1 = Mid(strETDVendor1, 1, 4) & "-" & Mid(strETDVendor1, 5, 2) & "-" & Mid(strETDVendor1, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETDPort1, 1, 4) & "-" & Mid(strETDPort1, 5, 2) & "-" & Mid(strETDPort1, 7, 2)) Then
                                        strMessage = "ETD Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Port : " & strETDPort1
                                        Exit For
                                    Else
                                        strETDPort1 = Mid(strETDPort1, 1, 4) & "-" & Mid(strETDPort1, 5, 2) & "-" & Mid(strETDPort1, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAPort1, 1, 4) & "-" & Mid(strETAPort1, 5, 2) & "-" & Mid(strETAPort1, 7, 2)) Then
                                        strMessage = "ETA Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Port : " & strETAPort1
                                        Exit For
                                    Else
                                        strETAPort1 = Mid(strETAPort1, 1, 4) & "-" & Mid(strETAPort1, 5, 2) & "-" & Mid(strETAPort1, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAFactory1, 1, 4) & "-" & Mid(strETAFactory1, 5, 2) & "-" & Mid(strETAFactory1, 7, 2)) Then
                                        strMessage = "ETA Factory is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strETAFactory1
                                        Exit For
                                    Else
                                        strETAFactory1 = Mid(strETAFactory1, 1, 4) & "-" & Mid(strETAFactory1, 5, 2) & "-" & Mid(strETAFactory1, 7, 2)
                                    End If

                                    If Not IsDate("20" & Mid(strDeliveryMonth1, 1, 2) & "-" & Mid(strDeliveryMonth1, 3, 2) & "-01") Then
                                        strMessage = "Period/Delivery Month is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strDeliveryMonth1
                                        Exit For
                                    Else
                                        strDeliveryMonth1 = "20" & Mid(strDeliveryMonth1, 1, 2) & "-" & Mid(strDeliveryMonth1, 3, 2) & "-01"
                                    End If
                                Case "H21"
                                    strOrderNo2 = Mid(strText, 4, 15).Trim
                                    strETDVendor2 = Mid(strText, 19, 8).Trim
                                    strETDPort2 = Mid(strText, 27, 8).Trim
                                    strETAPort2 = Mid(strText, 35, 8).Trim
                                    strETAFactory2 = Mid(strText, 43, 8).Trim
                                    strDeliveryMonth2 = Mid(strText, 51, 4).Trim

                                    If Not IsDate(Mid(strETDVendor2, 1, 4) & "-" & Mid(strETDVendor2, 5, 2) & "-" & Mid(strETDVendor2, 7, 2)) Then
                                        strMessage = "ETD Vendor is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Vendor : " & strETDVendor2
                                        Exit For
                                    Else
                                        strETDVendor2 = Mid(strETDVendor2, 1, 4) & "-" & Mid(strETDVendor2, 5, 2) & "-" & Mid(strETDVendor2, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETDPort2, 1, 4) & "-" & Mid(strETDPort2, 5, 2) & "-" & Mid(strETDPort2, 7, 2)) Then
                                        strMessage = "ETD Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Port : " & strETDPort2
                                        Exit For
                                    Else
                                        strETDPort2 = Mid(strETDPort2, 1, 4) & "-" & Mid(strETDPort2, 5, 2) & "-" & Mid(strETDPort2, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAPort2, 1, 4) & "-" & Mid(strETAPort2, 5, 2) & "-" & Mid(strETAPort2, 7, 2)) Then
                                        strMessage = "ETA Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Port : " & strETAPort2
                                        Exit For
                                    Else
                                        strETAPort2 = Mid(strETAPort2, 1, 4) & "-" & Mid(strETAPort2, 5, 2) & "-" & Mid(strETAPort2, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAFactory2, 1, 4) & "-" & Mid(strETAFactory2, 5, 2) & "-" & Mid(strETAFactory2, 7, 2)) Then
                                        strMessage = "ETA Factory is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strETAFactory2
                                        Exit For
                                    Else
                                        strETAFactory2 = Mid(strETAFactory2, 1, 4) & "-" & Mid(strETAFactory2, 5, 2) & "-" & Mid(strETAFactory2, 7, 2)
                                    End If

                                    If Not IsDate("20" & Mid(strDeliveryMonth2, 1, 2) & "-" & Mid(strDeliveryMonth2, 3, 2) & "-01") Then
                                        strMessage = "Period/Delivery Month is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strDeliveryMonth2
                                        Exit For
                                    Else
                                        strDeliveryMonth2 = "20" & Mid(strDeliveryMonth2, 1, 2) & "-" & Mid(strDeliveryMonth2, 3, 2) & "-01"
                                    End If
                                Case "H22"
                                    strOrderNo3 = Mid(strText, 4, 15).Trim
                                    strETDVendor3 = Mid(strText, 19, 8).Trim
                                    strETDPort3 = Mid(strText, 27, 8).Trim
                                    strETAPort3 = Mid(strText, 35, 8).Trim
                                    strETAFactory3 = Mid(strText, 43, 8).Trim
                                    strDeliveryMonth3 = Mid(strText, 51, 4).Trim

                                    If Not IsDate(Mid(strETDVendor3, 1, 4) & "-" & Mid(strETDVendor3, 5, 2) & "-" & Mid(strETDVendor3, 7, 2)) Then
                                        strMessage = "ETD Vendor is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Vendor : " & strETDVendor3
                                        Exit For
                                    Else
                                        strETDVendor3 = Mid(strETDVendor3, 1, 4) & "-" & Mid(strETDVendor3, 5, 2) & "-" & Mid(strETDVendor3, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETDPort3, 1, 4) & "-" & Mid(strETDPort3, 5, 2) & "-" & Mid(strETDPort3, 7, 2)) Then
                                        strMessage = "ETD Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Port : " & strETDPort3
                                        Exit For
                                    Else
                                        strETDPort3 = Mid(strETDPort3, 1, 4) & "-" & Mid(strETDPort3, 5, 2) & "-" & Mid(strETDPort3, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAPort3, 1, 4) & "-" & Mid(strETAPort3, 5, 2) & "-" & Mid(strETAPort3, 7, 2)) Then
                                        strMessage = "ETA Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Port : " & strETAPort3
                                        Exit For
                                    Else
                                        strETAPort3 = Mid(strETAPort3, 1, 4) & "-" & Mid(strETAPort3, 5, 2) & "-" & Mid(strETAPort3, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAFactory3, 1, 4) & "-" & Mid(strETAFactory3, 5, 2) & "-" & Mid(strETAFactory3, 7, 2)) Then
                                        strMessage = "ETA Factory is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strETAFactory3
                                        Exit For
                                    Else
                                        strETAFactory3 = Mid(strETAFactory3, 1, 4) & "-" & Mid(strETAFactory3, 5, 2) & "-" & Mid(strETAFactory3, 7, 2)
                                    End If

                                    If Not IsDate("20" & Mid(strDeliveryMonth3, 1, 2) & "-" & Mid(strDeliveryMonth3, 3, 2) & "-01") Then
                                        strMessage = "Period/Delivery Month is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strDeliveryMonth3
                                        Exit For
                                    Else
                                        strDeliveryMonth3 = "20" & Mid(strDeliveryMonth3, 1, 2) & "-" & Mid(strDeliveryMonth3, 3, 2) & "-01"
                                    End If
                                Case "H23"
                                    strOrderNo4 = Mid(strText, 4, 15).Trim
                                    strETDVendor4 = Mid(strText, 19, 8).Trim
                                    strETDPort4 = Mid(strText, 27, 8).Trim
                                    strETAPort4 = Mid(strText, 35, 8).Trim
                                    strETAFactory4 = Mid(strText, 43, 8).Trim
                                    strDeliveryMonth4 = Mid(strText, 51, 4).Trim

                                    If Not IsDate(Mid(strETDVendor4, 1, 4) & "-" & Mid(strETDVendor4, 5, 2) & "-" & Mid(strETDVendor4, 7, 2)) Then
                                        strMessage = "ETD Vendor is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Vendor : " & strETDVendor4
                                        Exit For
                                    Else
                                        strETDVendor4 = Mid(strETDVendor4, 1, 4) & "-" & Mid(strETDVendor4, 5, 2) & "-" & Mid(strETDVendor4, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETDPort4, 1, 4) & "-" & Mid(strETDPort4, 5, 2) & "-" & Mid(strETDPort4, 7, 2)) Then
                                        strMessage = "ETD Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Port : " & strETDPort4
                                        Exit For
                                    Else
                                        strETDPort4 = Mid(strETDPort4, 1, 4) & "-" & Mid(strETDPort4, 5, 2) & "-" & Mid(strETDPort4, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAPort4, 1, 4) & "-" & Mid(strETAPort4, 5, 2) & "-" & Mid(strETAPort4, 7, 2)) Then
                                        strMessage = "ETA Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Port : " & strETAPort4
                                        Exit For
                                    Else
                                        strETAPort4 = Mid(strETAPort4, 1, 4) & "-" & Mid(strETAPort4, 5, 2) & "-" & Mid(strETAPort4, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAFactory4, 1, 4) & "-" & Mid(strETAFactory4, 5, 2) & "-" & Mid(strETAFactory4, 7, 2)) Then
                                        strMessage = "ETA Factory is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strETAFactory4
                                        Exit For
                                    Else
                                        strETAFactory4 = Mid(strETAFactory4, 1, 4) & "-" & Mid(strETAFactory4, 5, 2) & "-" & Mid(strETAFactory4, 7, 2)
                                    End If

                                    If Not IsDate("20" & Mid(strDeliveryMonth4, 1, 2) & "-" & Mid(strDeliveryMonth4, 3, 2) & "-01") Then
                                        strMessage = "Period/Delivery Month is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strDeliveryMonth4
                                        Exit For
                                    Else
                                        strDeliveryMonth4 = "20" & Mid(strDeliveryMonth4, 1, 2) & "-" & Mid(strDeliveryMonth4, 3, 2) & "-01"
                                    End If
                                Case "H24"
                                    strOrderNo5 = Mid(strText, 4, 15).Trim
                                    strETDVendor5 = Mid(strText, 19, 8).Trim
                                    strETDPort5 = Mid(strText, 27, 8).Trim
                                    strETAPort5 = Mid(strText, 35, 8).Trim
                                    strETAFactory5 = Mid(strText, 43, 8).Trim
                                    strDeliveryMonth5 = Mid(strText, 51, 4).Trim

                                    If Not IsDate(Mid(strETDVendor5, 1, 4) & "-" & Mid(strETDVendor5, 5, 2) & "-" & Mid(strETDVendor5, 7, 2)) Then
                                        strMessage = "ETD Vendor is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Vendor : " & strETDVendor5
                                        Exit For
                                    Else
                                        strETDVendor5 = Mid(strETDVendor5, 1, 4) & "-" & Mid(strETDVendor5, 5, 2) & "-" & Mid(strETDVendor5, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETDPort5, 1, 4) & "-" & Mid(strETDPort5, 5, 2) & "-" & Mid(strETDPort5, 7, 2)) Then
                                        strMessage = "ETD Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETD Port : " & strETDPort5
                                        Exit For
                                    Else
                                        strETDPort5 = Mid(strETDPort5, 1, 4) & "-" & Mid(strETDPort5, 5, 2) & "-" & Mid(strETDPort5, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAPort5, 1, 4) & "-" & Mid(strETAPort5, 5, 2) & "-" & Mid(strETAPort5, 7, 2)) Then
                                        strMessage = "ETA Port is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Port : " & strETAPort5
                                        Exit For
                                    Else
                                        strETAPort5 = Mid(strETAPort5, 1, 4) & "-" & Mid(strETAPort5, 5, 2) & "-" & Mid(strETAPort5, 7, 2)
                                    End If

                                    If Not IsDate(Mid(strETAFactory5, 1, 4) & "-" & Mid(strETAFactory5, 5, 2) & "-" & Mid(strETAFactory5, 7, 2)) Then
                                        strMessage = "ETA Factory is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strETAFactory5
                                        Exit For
                                    Else
                                        strETAFactory5 = Mid(strETAFactory5, 1, 4) & "-" & Mid(strETAFactory5, 5, 2) & "-" & Mid(strETAFactory5, 7, 2)
                                    End If

                                    If Not IsDate("20" & Mid(strDeliveryMonth5, 1, 2) & "-" & Mid(strDeliveryMonth5, 3, 2) & "-01") Then
                                        strMessage = "Period/Delivery Month is not a valid date. Please check template" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "ETA Factory : " & strDeliveryMonth5
                                        Exit For
                                    Else
                                        strDeliveryMonth5 = "20" & Mid(strDeliveryMonth5, 1, 2) & "-" & Mid(strDeliveryMonth5, 3, 2) & "-01"
                                    End If
                                Case "H30"
                                    strForecastDate1 = Mid(strText, 4, 8).Trim
                                    strForecastDate2 = Mid(strText, 12, 8).Trim
                                    strForecastDate3 = Mid(strText, 20, 8).Trim
                                    strForecastDate4 = Mid(strText, 28, 8).Trim
                                    strForecastDate5 = Mid(strText, 36, 8).Trim
                                    strForecastDate6 = Mid(strText, 44, 8).Trim
                                    strForecastDate7 = Mid(strText, 52, 8).Trim
                                    strForecastDate8 = Mid(strText, 60, 8).Trim
                                Case "D10"
                                    ClearD10()

                                    strPartNo = Mid(strText, 4, 25).Trim
                                    strUOM = Mid(strText, 29, 3).Trim
                                    strOrderQty1 = Val(Mid(strText, 32, 8).Trim)
                                    strOrderQty2 = Val(Mid(strText, 40, 8).Trim)
                                    strOrderQty3 = Val(Mid(strText, 48, 8).Trim)
                                    strOrderQty4 = Val(Mid(strText, 56, 8).Trim)
                                    strOrderQty5 = Val(Mid(strText, 64, 8).Trim)

                                    strSQL = "SELECT * FROM MS_PartMapping WHERE AffiliateID = '" & strAffiliateID & "' AND PartNo = '" & strPartNo & "'"
                                    ds = uf_GetDataSet(strSQL)
                                    If ds.Tables(0).Rows.Count > 0 Then
                                        strSupplierID = ds.Tables(0).Rows("0")("SupplierID").ToString.Trim
                                    Else
                                        strSupplierID = ""
                                    End If
                                    ds.Clear()

                                    If strSupplierID = "" Then
                                        strMessage = "Can't get Supplier ID. Please check data Part Mapping" & vbCrLf & _
                                                    "File Name : " & Dir(pFile) & vbCrLf & _
                                                    "Line : " & intLine & vbCrLf & _
                                                    "Affiliate ID : " & strAffiliateID & vbCrLf & _
                                                    "Part No : " & strPartNo
                                        Exit For
                                    End If
                                Case "D11"
                                    strPartDescription = Mid(strText, 4, 30).Trim

                                    If intLine + 1 <= intLineEnd Then
                                        Dim strText_NextLine As String = readText(intLine + 1 - 1) ' Karena mulai dari index ke 0
                                        If Mid(strText_NextLine, 1, 3).Trim <> "D20" Then
                                            ProcessInsertData(strAffiliateID, strSupplierID, strForwarderID, pFile, sqlConn, sqlTran, strMessage)
                                            If strMessage <> "" Then
                                                Exit For
                                            End If
                                        End If
                                    End If

                                Case "D20"
                                    strForecastQty1 = Val(Mid(strText, 4, 9).Trim)
                                    strForecastQty2 = Val(Mid(strText, 13, 9).Trim)
                                    strForecastQty3 = Val(Mid(strText, 22, 9).Trim)
                                    strForecastQty4 = Val(Mid(strText, 31, 9).Trim)
                                    strForecastQty5 = Val(Mid(strText, 40, 9).Trim)
                                    strForecastQty6 = Val(Mid(strText, 49, 9).Trim)
                                    strForecastQty7 = Val(Mid(strText, 58, 9).Trim)
                                    strForecastQty8 = Val(Mid(strText, 67, 9).Trim)
                                Case "T00"
                                    GoTo NextLine
                            End Select


                            If strHeaderID = "D20" Then
                                ProcessInsertData(strAffiliateID, strSupplierID, strForwarderID, pFile, sqlConn, sqlTran, strMessage)
                                If strMessage <> "" Then
                                    Exit For
                                End If
                            End If
NextLine:
                        Next

                        If strMessage = "" Then
                            sqlTran.Commit()
                        End If

                    End Using

                    sr.Close()
                    sr.Dispose()
                End Using

                If strMessage = "" Then
                    UpdateLog("Processing file : " & Dir(pFile) & " SUCCESS...")

                    If System.IO.File.Exists(pFile) = True Then
                        Dim fi As New FileInfo(pFile)
                        fi.CopyTo(gs_UploadPath & "\Backup\" & fi.Name, True)
                        fi.Delete()
                    End If
                Else
                    UpdateLog("Processing file : " & Dir(pFile) & " FAIL..." & vbCrLf & strMessage)
                    
                    Dim strSendMailMessage As String = ""
                    strSendMailMessage = uf_SendEmail(strMessage, strAffiliateCode)

                    If strSendMailMessage <> "" Then
                        UpdateLog(strSendMailMessage)
                    End If

                    If System.IO.File.Exists(pFile) = True Then
                        Dim fi As New FileInfo(pFile)
                        fi.CopyTo(gs_UploadPath & "\Error\" & fi.Name, True)
                        fi.Delete()
                    End If
                End If
            End Using
        Catch ex As Exception
            UpdateLog("Processing file : " & Dir(pFile) & " FAIL..." & vbCrLf & ex.Message)

            If System.IO.File.Exists(pFile) = True Then
                Dim fi As New FileInfo(pFile)
                fi.CopyTo(gs_UploadPath & "\Error\" & fi.Name, True)
                fi.Delete()
            End If
        End Try
    End Sub

    Private Sub ProcessInsertData(ByVal strAffiliateID As String, ByVal strSupplierID As String, ByVal strForwarderID As String, ByVal pFile As String, ByVal sqlConn As SqlConnection, ByVal sqlTran As SqlTransaction, ByRef strMessage As String)
        Dim SqlComm As SqlCommand
        Dim strSQL As String
        Dim ds As DataSet

        'Week 1
        If strOrderNo1 <> "" And (strOrderQty1 <> "0" Or strForecastQty1 <> "0" Or strForecastQty2 <> "0" Or strForecastQty3 <> "0") Then

            strSQL = "SELECT * FROM MS_ETD_Export " & vbCrLf & _
                "WHERE AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ETDVendor = '" & strETDVendor1 & "' " & vbCrLf & _
                "AND ETDPort = '" & strETDPort1 & "' " & vbCrLf & _
                "AND ETAPort = '" & strETAPort1 & "' " & vbCrLf & _
                "AND ETAFactory = '" & strETAFactory1 & "' " & vbCrLf & _
                "AND Period = '" & strDeliveryMonth1 & "' " & vbCrLf & _
                "AND [Week] = 1 "

            ds = uf_GetDataSet(strSQL)
            If ds.Tables(0).Rows.Count = 0 Then
                strMessage = "Can't get ETD and ETA for Week 1. Please check data Time Chart Master" & vbCrLf & _
                            "File Name : " & Dir(pFile) & vbCrLf & _
                            "Affiliate ID : " & strAffiliateID & vbCrLf & _
                            "Supplier ID : " & strSupplierID & vbCrLf & _
                            "ETD Vendor : " & strETDVendor1 & vbCrLf & _
                            "ETD Port : " & strETDPort1 & vbCrLf & _
                            "ETA Port : " & strETAPort1 & vbCrLf & _
                            "ETA Factory : " & strETAFactory1 & vbCrLf & _
                            "Period : " & strDeliveryMonth1
                Exit Sub
            End If
            ds.Clear()

            'PO Detail
            strSQL = "SELECT * FROM PO_Detail_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo1 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' " & vbCrLf & _
                "AND PartNo = '" & strPartNo & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Detail_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "PartNo, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "Week1, " & vbCrLf & _
                    "TotalPOQty, " & vbCrLf & _
                    "Forecast1, " & vbCrLf & _
                    "Forecast2, " & vbCrLf & _
                    "Forecast3, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo1 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strPartNo & "', " & vbCrLf & _
                    "'" & strOrderNo1 & "', " & vbCrLf & _
                    "'" & Val(strOrderQty1) & "', " & vbCrLf & _
                    "'" & Val(strOrderQty1) & "', " & vbCrLf & _
                    "'" & Val(strForecastQty1) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty2) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty3) & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()

            'PO Master
            strSQL = "SELECT * FROM PO_Master_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo1 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Master_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "Period, " & vbCrLf & _
                    "CommercialCls, " & vbCrLf & _
                    "EmergencyCls, " & vbCrLf & _
                    "ShipCls, " & vbCrLf & _
                    "ErrorStatus, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "ETDVendor1, " & vbCrLf & _
                    "ETDPort1, " & vbCrLf & _
                    "ETAPort1, " & vbCrLf & _
                    "ETAFactory1, " & vbCrLf & _
                    "UploadDate, " & vbCrLf & _
                    "UploadUser, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf

                strSQL = strSQL & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo1 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strDeliveryMonth1 & "', " & vbCrLf & _
                    "'" & IIf(strCommercialCls = "C", "1", "0") & "', " & vbCrLf & _
                    "'" & strOrderRequest & "', " & vbCrLf & _
                    "'" & strFreight & "', " & vbCrLf & _
                    "'OK', " & vbCrLf & _
                    "'" & strOrderNo1 & "', " & vbCrLf & _
                    "'" & strETDVendor1 & "', " & vbCrLf & _
                    "'" & strETDPort1 & "', " & vbCrLf & _
                    "'" & strETAPort1 & "', " & vbCrLf & _
                    "'" & strETAFactory1 & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()
        End If

        'Week 2
        If strOrderNo2 <> "" And (strOrderQty2 <> "0" Or strForecastQty1 <> "0" Or strForecastQty2 <> "0" Or strForecastQty3 <> "0") Then

            strSQL = "SELECT * FROM MS_ETD_Export " & vbCrLf & _
                "WHERE AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ETDVendor = '" & strETDVendor2 & "' " & vbCrLf & _
                "AND ETDPort = '" & strETDPort2 & "' " & vbCrLf & _
                "AND ETAPort = '" & strETAPort2 & "' " & vbCrLf & _
                "AND ETAFactory = '" & strETAFactory2 & "' " & vbCrLf & _
                "AND Period = '" & strDeliveryMonth2 & "' " & vbCrLf & _
                "AND [Week] = 2 "

            ds = uf_GetDataSet(strSQL)
            If ds.Tables(0).Rows.Count = 0 Then
                strMessage = "Can't get ETD and ETA for Week 2. Please check data Time Chart Master" & vbCrLf & _
                            "File Name : " & Dir(pFile) & vbCrLf & _
                            "Affiliate ID : " & strAffiliateID & vbCrLf & _
                            "Supplier ID : " & strSupplierID & vbCrLf & _
                            "ETD Vendor : " & strETDVendor2 & vbCrLf & _
                            "ETD Port : " & strETDPort2 & vbCrLf & _
                            "ETA Port : " & strETAPort2 & vbCrLf & _
                            "ETA Factory : " & strETAFactory2 & vbCrLf & _
                            "Period : " & strDeliveryMonth2
                Exit Sub
            End If
            ds.Clear()

            'PO Detail
            strSQL = "SELECT * FROM PO_Detail_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo2 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' " & vbCrLf & _
                "AND PartNo = '" & strPartNo & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Detail_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "PartNo, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "Week1, " & vbCrLf & _
                    "TotalPOQty, " & vbCrLf & _
                    "Forecast1, " & vbCrLf & _
                    "Forecast2, " & vbCrLf & _
                    "Forecast3, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo2 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strPartNo & "', " & vbCrLf & _
                    "'" & strOrderNo2 & "', " & vbCrLf & _
                    "'" & Val(strOrderQty2) & "', " & vbCrLf & _
                    "'" & Val(strOrderQty2) & "', " & vbCrLf & _
                    "'" & Val(strForecastQty1) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty2) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty3) & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()

            'PO Master
            strSQL = "SELECT * FROM PO_Master_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo2 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Master_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "Period, " & vbCrLf & _
                    "CommercialCls, " & vbCrLf & _
                    "EmergencyCls, " & vbCrLf & _
                    "ShipCls, " & vbCrLf & _
                    "ErrorStatus, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "ETDVendor1, " & vbCrLf & _
                    "ETDPort1, " & vbCrLf & _
                    "ETAPort1, " & vbCrLf & _
                    "ETAFactory1, " & vbCrLf & _
                    "UploadDate, " & vbCrLf & _
                    "UploadUser, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf

                strSQL = strSQL & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo2 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strDeliveryMonth2 & "', " & vbCrLf & _
                    "'" & IIf(strCommercialCls = "C", "1", "0") & "', " & vbCrLf & _
                    "'" & strOrderRequest & "', " & vbCrLf & _
                    "'" & strFreight & "', " & vbCrLf & _
                    "'OK', " & vbCrLf & _
                    "'" & strOrderNo2 & "', " & vbCrLf & _
                    "'" & strETDVendor2 & "', " & vbCrLf & _
                    "'" & strETDPort2 & "', " & vbCrLf & _
                    "'" & strETAPort2 & "', " & vbCrLf & _
                    "'" & strETAFactory2 & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()
        End If

        'Week 3
        If strOrderNo3 <> "" And (strOrderQty3 <> "0" Or strForecastQty1 <> "0" Or strForecastQty2 <> "0" Or strForecastQty3 <> "0") Then

            strSQL = "SELECT * FROM MS_ETD_Export " & vbCrLf & _
                "WHERE AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ETDVendor = '" & strETDVendor3 & "' " & vbCrLf & _
                "AND ETDPort = '" & strETDPort3 & "' " & vbCrLf & _
                "AND ETAPort = '" & strETAPort3 & "' " & vbCrLf & _
                "AND ETAFactory = '" & strETAFactory3 & "' " & vbCrLf & _
                "AND Period = '" & strDeliveryMonth3 & "' " & vbCrLf & _
                "AND [Week] = 3 "

            ds = uf_GetDataSet(strSQL)
            If ds.Tables(0).Rows.Count = 0 Then
                strMessage = "Can't get ETD and ETA for Week 3. Please check data Time Chart Master" & vbCrLf & _
                            "File Name : " & Dir(pFile) & vbCrLf & _
                            "Affiliate ID : " & strAffiliateID & vbCrLf & _
                            "Supplier ID : " & strSupplierID & vbCrLf & _
                            "ETD Vendor : " & strETDVendor3 & vbCrLf & _
                            "ETD Port : " & strETDPort3 & vbCrLf & _
                            "ETA Port : " & strETAPort3 & vbCrLf & _
                            "ETA Factory : " & strETAFactory3 & vbCrLf & _
                            "Period : " & strDeliveryMonth3
                Exit Sub
            End If
            ds.Clear()

            'PO Detail
            strSQL = "SELECT * FROM PO_Detail_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo3 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' " & vbCrLf & _
                "AND PartNo = '" & strPartNo & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Detail_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "PartNo, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "Week1, " & vbCrLf & _
                    "TotalPOQty, " & vbCrLf & _
                    "Forecast1, " & vbCrLf & _
                    "Forecast2, " & vbCrLf & _
                    "Forecast3, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo3 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strPartNo & "', " & vbCrLf & _
                    "'" & strOrderNo3 & "', " & vbCrLf & _
                    "'" & Val(strOrderQty3) & "', " & vbCrLf & _
                    "'" & Val(strOrderQty3) & "', " & vbCrLf & _
                    "'" & Val(strForecastQty1) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty2) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty3) & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()

            'PO Master
            strSQL = "SELECT * FROM PO_Master_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo3 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Master_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "Period, " & vbCrLf & _
                    "CommercialCls, " & vbCrLf & _
                    "EmergencyCls, " & vbCrLf & _
                    "ShipCls, " & vbCrLf & _
                    "ErrorStatus, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "ETDVendor1, " & vbCrLf & _
                    "ETDPort1, " & vbCrLf & _
                    "ETAPort1, " & vbCrLf & _
                    "ETAFactory1, " & vbCrLf & _
                    "UploadDate, " & vbCrLf & _
                    "UploadUser, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf

                strSQL = strSQL & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo3 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strDeliveryMonth3 & "', " & vbCrLf & _
                    "'" & IIf(strCommercialCls = "C", "1", "0") & "', " & vbCrLf & _
                    "'" & strOrderRequest & "', " & vbCrLf & _
                    "'" & strFreight & "', " & vbCrLf & _
                    "'OK', " & vbCrLf & _
                    "'" & strOrderNo3 & "', " & vbCrLf & _
                    "'" & strETDVendor3 & "', " & vbCrLf & _
                    "'" & strETDPort3 & "', " & vbCrLf & _
                    "'" & strETAPort3 & "', " & vbCrLf & _
                    "'" & strETAFactory3 & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()
        End If

        'Week 4
        If strOrderNo4 <> "" And (strOrderQty4 <> "0" Or strForecastQty1 <> "0" Or strForecastQty2 <> "0" Or strForecastQty3 <> "0") Then

            strSQL = "SELECT * FROM MS_ETD_Export " & vbCrLf & _
                "WHERE AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ETDVendor = '" & strETDVendor4 & "' " & vbCrLf & _
                "AND ETDPort = '" & strETDPort4 & "' " & vbCrLf & _
                "AND ETAPort = '" & strETAPort4 & "' " & vbCrLf & _
                "AND ETAFactory = '" & strETAFactory4 & "' " & vbCrLf & _
                "AND Period = '" & strDeliveryMonth4 & "' " & vbCrLf & _
                "AND [Week] = 4 "

            ds = uf_GetDataSet(strSQL)
            If ds.Tables(0).Rows.Count = 0 Then
                strMessage = "Can't get ETD and ETA for Week 4. Please check data Time Chart Master" & vbCrLf & _
                            "File Name : " & Dir(pFile) & vbCrLf & _
                            "Affiliate ID : " & strAffiliateID & vbCrLf & _
                            "Supplier ID : " & strSupplierID & vbCrLf & _
                            "ETD Vendor : " & strETDVendor4 & vbCrLf & _
                            "ETD Port : " & strETDPort4 & vbCrLf & _
                            "ETA Port : " & strETAPort4 & vbCrLf & _
                            "ETA Factory : " & strETAFactory4 & vbCrLf & _
                            "Period : " & strDeliveryMonth4
                Exit Sub
            End If
            ds.Clear()

            'PO Detail
            strSQL = "SELECT * FROM PO_Detail_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo4 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' " & vbCrLf & _
                "AND PartNo = '" & strPartNo & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Detail_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "PartNo, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "Week1, " & vbCrLf & _
                    "TotalPOQty, " & vbCrLf & _
                    "Forecast1, " & vbCrLf & _
                    "Forecast2, " & vbCrLf & _
                    "Forecast3, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo4 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strPartNo & "', " & vbCrLf & _
                    "'" & strOrderNo4 & "', " & vbCrLf & _
                    "'" & Val(strOrderQty4) & "', " & vbCrLf & _
                    "'" & Val(strOrderQty4) & "', " & vbCrLf & _
                    "'" & Val(strForecastQty1) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty2) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty3) & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()

            'PO Master
            strSQL = "SELECT * FROM PO_Master_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo4 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Master_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "Period, " & vbCrLf & _
                    "CommercialCls, " & vbCrLf & _
                    "EmergencyCls, " & vbCrLf & _
                    "ShipCls, " & vbCrLf & _
                    "ErrorStatus, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "ETDVendor1, " & vbCrLf & _
                    "ETDPort1, " & vbCrLf & _
                    "ETAPort1, " & vbCrLf & _
                    "ETAFactory1, " & vbCrLf & _
                    "UploadDate, " & vbCrLf & _
                    "UploadUser, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf

                strSQL = strSQL & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo4 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strDeliveryMonth4 & "', " & vbCrLf & _
                    "'" & IIf(strCommercialCls = "C", "1", "0") & "', " & vbCrLf & _
                    "'" & strOrderRequest & "', " & vbCrLf & _
                    "'" & strFreight & "', " & vbCrLf & _
                    "'OK', " & vbCrLf & _
                    "'" & strOrderNo4 & "', " & vbCrLf & _
                    "'" & strETDVendor4 & "', " & vbCrLf & _
                    "'" & strETDPort4 & "', " & vbCrLf & _
                    "'" & strETAPort4 & "', " & vbCrLf & _
                    "'" & strETAFactory4 & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()
        End If

        'Week 5
        If strOrderNo5 <> "" And (strOrderQty5 <> "0" Or strForecastQty1 <> "0" Or strForecastQty2 <> "0" Or strForecastQty3 <> "0") Then

            strSQL = "SELECT * FROM MS_ETD_Export " & vbCrLf & _
                "WHERE AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ETDVendor = '" & strETDVendor5 & "' " & vbCrLf & _
                "AND ETDPort = '" & strETDPort5 & "' " & vbCrLf & _
                "AND ETAPort = '" & strETAPort5 & "' " & vbCrLf & _
                "AND ETAFactory = '" & strETAFactory5 & "' " & vbCrLf & _
                "AND Period = '" & strDeliveryMonth5 & "' " & vbCrLf & _
                "AND [Week] = 5 "

            ds = uf_GetDataSet(strSQL)
            If ds.Tables(0).Rows.Count = 0 Then
                strMessage = "Can't get ETD and ETA for Week 5. Please check data Time Chart Master" & vbCrLf & _
                            "File Name : " & Dir(pFile) & vbCrLf & _
                            "Affiliate ID : " & strAffiliateID & vbCrLf & _
                            "Supplier ID : " & strSupplierID & vbCrLf & _
                            "ETD Vendor : " & strETDVendor5 & vbCrLf & _
                            "ETD Port : " & strETDPort5 & vbCrLf & _
                            "ETA Port : " & strETAPort5 & vbCrLf & _
                            "ETA Factory : " & strETAFactory5 & vbCrLf & _
                            "Period : " & strDeliveryMonth5
                Exit Sub
            End If
            ds.Clear()

            'PO Detail
            strSQL = "SELECT * FROM PO_Detail_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo5 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' " & vbCrLf & _
                "AND PartNo = '" & strPartNo & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Detail_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "PartNo, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "Week1, " & vbCrLf & _
                    "TotalPOQty, " & vbCrLf & _
                    "Forecast1, " & vbCrLf & _
                    "Forecast2, " & vbCrLf & _
                    "Forecast3, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo5 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strPartNo & "', " & vbCrLf & _
                    "'" & strOrderNo5 & "', " & vbCrLf & _
                    "'" & Val(strOrderQty5) & "', " & vbCrLf & _
                    "'" & Val(strOrderQty5) & "', " & vbCrLf & _
                    "'" & Val(strForecastQty1) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty2) & "',  " & vbCrLf & _
                    "'" & Val(strForecastQty3) & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()

            'PO Master
            strSQL = "SELECT * FROM PO_Master_Export WITH (NOLOCK) " & vbCrLf & _
                "WHERE PONo = '" & strOrderNo5 & "' " & vbCrLf & _
                "AND AffiliateID = '" & strAffiliateID & "' " & vbCrLf & _
                "AND SupplierID = '" & strSupplierID & "' " & vbCrLf & _
                "AND ForwarderID = '" & strForwarderID & "' "
            ds = uf_GetDataSet(strSQL)

            If ds.Tables(0).Rows.Count = 0 Then
                strSQL = "INSERT INTO PO_Master_Export (" & vbCrLf & _
                    "PONo, " & vbCrLf & _
                    "AffiliateID, " & vbCrLf & _
                    "SupplierID, " & vbCrLf & _
                    "ForwarderID, " & vbCrLf & _
                    "Period, " & vbCrLf & _
                    "CommercialCls, " & vbCrLf & _
                    "EmergencyCls, " & vbCrLf & _
                    "ShipCls, " & vbCrLf & _
                    "ErrorStatus, " & vbCrLf & _
                    "OrderNo1, " & vbCrLf & _
                    "ETDVendor1, " & vbCrLf & _
                    "ETDPort1, " & vbCrLf & _
                    "ETAPort1, " & vbCrLf & _
                    "ETAFactory1, " & vbCrLf & _
                    "UploadDate, " & vbCrLf & _
                    "UploadUser, " & vbCrLf & _
                    "EntryDate, " & vbCrLf & _
                    "EntryUser) " & vbCrLf

                strSQL = strSQL & _
                    "VALUES (" & vbCrLf & _
                    "'" & strOrderNo5 & "', " & vbCrLf & _
                    "'" & strAffiliateID & "', " & vbCrLf & _
                    "'" & strSupplierID & "', " & vbCrLf & _
                    "'" & strForwarderID & "', " & vbCrLf & _
                    "'" & strDeliveryMonth5 & "', " & vbCrLf & _
                    "'" & IIf(strCommercialCls = "C", "1", "0") & "', " & vbCrLf & _
                    "'" & strOrderRequest & "', " & vbCrLf & _
                    "'" & strFreight & "', " & vbCrLf & _
                    "'OK', " & vbCrLf & _
                    "'" & strOrderNo5 & "', " & vbCrLf & _
                    "'" & strETDVendor5 & "', " & vbCrLf & _
                    "'" & strETDPort5 & "', " & vbCrLf & _
                    "'" & strETAPort5 & "', " & vbCrLf & _
                    "'" & strETAFactory5 & "', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES', " & vbCrLf & _
                    "GETDATE(), " & vbCrLf & _
                    "'OES') "

                SqlComm = New SqlCommand(strSQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
            End If
            ds.Clear()
        End If
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub frmUpload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        up_GetSetting()

        lblVersion.Text = "Version " & Application.ProductVersion
        txtNext.Text = Format(DateAdd(DateInterval.Second, CDbl(txtSechedule.Text), Date.Now), "dd-MMM-yyyy HH:mm:ss")
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Format(Date.Now, "dd-MMM-yyyy HH:mm:ss") = txtNext.Text Then
            Timer1.Enabled = False

            UploadProcess()

            txtLast.Text = Format(Date.Now, "dd-MMM-yyyy HH:mm:ss")
            txtNext.Text = Format(DateAdd(DateInterval.Second, CDbl(txtSechedule.Text), Date.Now), "dd-MMM-yyyy HH:mm:ss")
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub btnManual_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnManual.Click
        Timer1.Enabled = False

        UploadProcess()

        txtLast.Text = Format(Date.Now, "dd-MMM-yyyy HH:mm:ss")
        txtNext.Text = Format(DateAdd(DateInterval.Second, CDbl(txtSechedule.Text), Date.Now), "dd-MMM-yyyy HH:mm:ss")
        Timer1.Enabled = True
    End Sub
End Class