'######################################################################################################
'System         : PGM  Production Instruction System
'Program        : Logging
'
'Overview       : This program about
'                 1. procedure write error log
'                 2. procedure write process log
'                 3. procedure write event log

'parameter : User ID, VIN, Process Name, Menu Name, Error Message

'Created By     : Joko
'Created Date   : October 09, 2013

'Modify History
'                           
'######################################################################################################

Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class clsLog

#Region "Declaration"
    Private UserLogin As String
    Private ConStr As String
    Private li_StartTime As Date
    Private li_EndTime As Date
    Private li_Duration As TimeSpan
    Private dtMsg As DataTable
    Const ConnectionErrorMsg As String = "A network-related or instance-specific error occurred while establishing a connection to SQL Server"

    Public Enum ErrSeverity
        ALERT = 1
        ERR = 2
        INFO = 3
    End Enum

#End Region

#Region "Initialization"

    Public Sub New(ByVal pConStr As String, ByVal pUserlogin As String)
        ' Add any initialization after the InitializeComponent() call.
        ConStr = pConStr
        UserLogin = pUserlogin
    End Sub

#End Region

#Region "Procedures and Functions"

    ''' <summary>
    ''' Gets message description from database based on message ID.
    ''' </summary>
    ''' <param name="pLogID">
    ''' Log ID.
    ''' </param>
    ''' <returns>
    ''' Returns message description.
    ''' </returns>
    ''' <remarks></remarks>
    Public Function uf_ConvertMsg(ByVal pLogID As String) As String
        Try
            Dim Msg As String = ""
            Dim column(0) As DataColumn
            column(0) = dtMsg.Columns(0)
            dtMsg.PrimaryKey = column

            Dim row As DataRow = dtMsg.Rows.Find(pLogID)
            If row IsNot Nothing AndAlso row(1) IsNot Nothing Then
                Msg = row(1).ToString.TrimEnd
            End If
            Return Msg
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Gets event ID from database based on Log ID.
    ''' </summary>
    ''' <param name="pLogID">
    ''' Log ID.
    ''' </param>
    ''' <returns>
    ''' Returns Event ID
    ''' </returns>
    ''' <remarks></remarks>
    Private Function uf_ConvertEventID(ByVal pLogID As String) As Integer
        Try
            Dim pEventID As Integer
            pEventID = 0
            Dim column(0) As DataColumn
            column(0) = dtMsg.Columns(0)
            dtMsg.PrimaryKey = column

            Dim row As DataRow = dtMsg.Rows.Find(pLogID)
            If row IsNot Nothing AndAlso row(3) IsNot Nothing Then
                pEventID = Val(row(3).ToString)
            End If
            Return pEventID
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    '*************************************************************
    'NAME:          WriteToErrorLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    pScreenName     - Name of Screen Process
    '               pErrSummary     - Err Summary Process
    '               pErrID          - Error ID
    '               pErrSeverity    - Enum ALERT, INFO and Err of the error file entry
    'RETURNS:       Nothing
    '*************************************************************
    ''' <summary>
    ''' this procedure for process error log
    ''' </summary>
    ''' <param name="pScreenName">Screen Name</param>
    ''' <param name="pErrSummary">Err Summary</param>
    ''' <param name="pErrID">Err ID</param>
    ''' <param name="pErrSeverity">Err Severity</param>
    ''' <remarks>
    ''' 1. create \Log\ directory if not create before
    ''' 2. create \Log\Process\ directory if not create before
    ''' 3. checking the process log file with the same date if found it will be add the process information with the exist file
    ''' 4. create the process detail information
    ''' 5. keep the process log file from last 90 days until now otherwise will removed
    ''' </remarks>
    Public Sub WriteToErrorLog(ByVal pScreenName As String, Optional ByVal pErrSummary As String = "", Optional ByVal pErrID As Integer = 9999, Optional ByVal pErrSeverity As ErrSeverity = ErrSeverity.ERR, Optional ByVal pHarigami As Boolean = False)

        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_dateFolder As String
        Dim ls_CompName As String = uf_CompName()
        Dim ls_LogFolder As String = "D:\PASI EBWEB\"
        ls_dateFolder = Format(Now, "yyyyMMdd")
        ls_Date = Format(Now, "yyyyMMdd")
        ls_ErrType = uf_ErrSeverity(pErrSeverity)
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")

        If Not System.IO.Directory.Exists("D:\") Then
            ls_LogFolder = Application.StartupPath
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder) Then
            System.IO.Directory.CreateDirectory(ls_LogFolder)
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder & "\Log" &
        "\Error") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log" &
            "\Error")
        End If

        'check the file
        Dim fs As FileStream = New FileStream(ls_LogFolder & "\Log" &
        "\Error\err" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)

        s.Close()
        fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream(ls_LogFolder & "\Log" &
        "\Error\err" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)

        s1.Write("" & Format(Now, "dd/MM/yyyy HH:mm:ss") & " ")
        s1.Write("[" & UserLogin & "]" & " ")
        s1.Write("[" & ls_CompName & "] ")
        s1.Write("" & pScreenName & "" & " ")
        s1.Write("[" & ls_ErrType & "]" & " ")
        s1.Write("" & pErrSummary & "" & "")
        s1.Write("" & vbCrLf)
        s1.Close()
        fs1.Close()

    End Sub

    '*************************************************************
    'NAME:          WriteToProcessLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    pStartTime      - Start of time process of the error file entry
    '               pScreenName     - Name of Screen Process
    '               pCustomMsg      - Name of Process of the error file entry
    '               pErrSummary     - Err summary process
    '               pErrID          - Error ID
    '               pErrSeverity    - Enum ALERT, INFO and Err of the error file entry
    '               pWriteEventLog  - True false for Write to Event Log
    '               pStartEndStatus - Different Format for start and end
    '               pUseLogTime     - True or False use log time
    'RETURNS:       Nothing
    '*************************************************************
    ''' <summary>
    ''' this procedure for process error log
    ''' </summary>
    ''' <param name="pStartTime">Start Time</param>
    ''' <param name="pScreenName">Screen Name</param>
    ''' <param name="pCustomMsg">Message</param>
    ''' <param name="pErrSummary">Err Summary</param>
    ''' <param name="pErrID">Err ID</param>
    ''' <param name="pErrSeverity">Err Severity</param>
    ''' <param name="pWriteToEventLog">Write to Event Log</param>
    ''' <param name="pStartEndStatus">Start or End Status</param>
    ''' <param name="pUseLogTime">Use Log Time</param>
    ''' <remarks>
    ''' 1. create \Log\ directory if not create before
    ''' 2. create \Log\Process\ directory if not create before
    ''' 3. checking the process log file with the same date if found it will be add the process information with the exist file
    ''' 4. create the process detail information
    ''' 5. keep the process log file from last 90 days until now otherwise will removed
    ''' </remarks>
    Public Sub WriteToProcessLog(ByVal pStartTime As Date, _
                                 ByVal pScreenName As String, _                                 
                                 Optional ByVal pCustomMsg As String = "", _
                                 Optional ByVal pErrSummary As String = "", _
                                 Optional ByVal pErrID As Integer = 9999, _
                                 Optional ByVal pErrSeverity As ErrSeverity = ErrSeverity.INFO, _
                                 Optional ByVal pWriteToEventLog As Boolean = False, _
                                 Optional ByVal pStartEndStatus As String = "", _
                                 Optional ByVal pUseLogTime As Boolean = False)

        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_Duration As String
        Dim ls_CompName As String = uf_CompName()
        Dim ls_LogFolder As String = "D:\PASI EBWEB\" 'pDirectory '"D:\LogFile"

        If Not System.IO.Directory.Exists("D:\") Then
            ls_LogFolder = Application.StartupPath
        End If

        li_StartTime = pStartTime
        li_EndTime = Now
        li_Duration = li_EndTime - li_StartTime
        ls_Duration = uf_AddSpace(Format(li_Duration.TotalMilliseconds, "###.#0") & " (ms)", 15)

        ls_Date = Format(Now, "yyyyMMdd")
        ls_ErrType = uf_ErrSeverity(pErrSeverity)
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")

        If Not System.IO.Directory.Exists(ls_LogFolder) Then
            System.IO.Directory.CreateDirectory(ls_LogFolder)
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder & "\Log") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log")
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder & "\Log" &
        "\Process") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log" &
            "\Process")
        End If

        'check the file
        Dim fs As FileStream = New FileStream(ls_LogFolder &
        "\Log\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)

        s.Close()
        fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream(ls_LogFolder &
        "\Log\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)

        If pStartEndStatus.ToUpper.Trim = "START" Then
            If pUseLogTime = True Then
                s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
            Else
                s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            End If
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("start.")
            s1.Write("" & vbCrLf)
            s1.Write("* * * * *")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        ElseIf pStartEndStatus.ToUpper.Trim = "END" Then
            s1.Write("* * * * *")
            s1.Write("" & vbCrLf)
            If pUseLogTime = True Then
                s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
            Else
                s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            End If
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("end.")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        ElseIf pStartEndStatus = "" Then
            If ls_ErrType = "I" Then
                If pUseLogTime = True Then
                    s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
                Else
                    s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
                End If
                s1.Write("[" & UserLogin & "] ")
                s1.Write("[" & ls_CompName & "] ")
                s1.Write("" & pScreenName & " ")
                s1.Write("[" & ls_ErrType & "] ")
                s1.Write("" & pCustomMsg & " ")
                s1.Write("" & pErrSummary & " ")
                s1.Write("" & ls_Duration & "")
                s1.Write("" & vbCrLf)
                s1.Close()
                fs1.Close()
            Else
                If pUseLogTime = True Then
                    s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
                Else
                    s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
                End If
                s1.Write("[" & UserLogin & "] ")
                s1.Write("[" & ls_CompName & "] ")
                s1.Write("" & pScreenName & " ")
                s1.Write("[" & ls_ErrType & "] ")
                s1.Write("" & pCustomMsg & " ")
                s1.Write("" & pErrSummary & "")
                s1.Write("" & vbCrLf)
                s1.Close()
                fs1.Close()
            End If
        End If

        'Dim ls_Dir As New IO.DirectoryInfo(ls_LogFolder & "\Log\Process\")
        'Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
        'Dim ls_File As IO.FileInfo
        'Dim li_index As Long
        'Dim ls_log As String
        'Dim li_CountDate As Long
        'Dim ls_Temp As List(Of String) = New List(Of String)

        'li_index = 0
        'For Each ls_File In ls_GetFile

        '    ls_Temp.Add(ls_File.ToString)

        '    ls_log = Right(ls_Temp.Item(li_index), 12)
        '    ls_log = Mid(ls_log, 1, 8)
        '    ls_log = Left(ls_log, 4) & "/" & Mid(ls_log, 5, 2) & "/" & Right(ls_log, 2)
        '    li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_log), "yyyy/MM/dd")), CDate(Format(Now, "yyyy/MM/dd")))
        '    If li_CountDate > 90 Then
        '        File.Delete(ls_LogFolder & "\Log\Process\prc" & pScreenName & "_" & Format(CDate(ls_log), "yyyyMMdd") & ".log")
        '    End If

        '    li_index = li_index + 1
        'Next
    End Sub

    '*************************************************************
    'NAME:          WriteToProcessLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    pStartTime      - Start of time process of the error file entry
    '               pEndTime        - End of time process of the error file entry
    '               pScreenName     - Name of Screen Process
    '               pCustomMsg      - Name of Process of the error file entry
    '               pErrSummary     - Err summary process
    '               pErrID          - Error ID
    '               pErrSeverity    - Enum ALERT, INFO and Err of the error file entry
    '               pWriteEventLog  - True false for Write to Event Log
    '               pStartEndStatus - Different Format for start and end
    '               pUseLogTime     - True or False use log time
    'RETURNS:       Nothing
    '*************************************************************
    ''' <summary>
    ''' this procedure for process error log
    ''' </summary>
    ''' <param name="pStartTime">Start Time</param>
    ''' <param name="pEndTime">End Time</param>
    ''' <param name="pScreenName">Screen Name</param>
    ''' <param name="pCustomMsg">Message</param>
    ''' <param name="pErrSummary">Err Summary</param>
    ''' <param name="pErrID">Err ID</param>
    ''' <param name="pErrSeverity">Err Severity</param>
    ''' <param name="pWriteToEventLog">Write to Event Log</param>
    ''' <param name="pStartEndStatus">Start or End Status</param>
    ''' <param name="pUseLogTime">Use Log Time</param>
    ''' <remarks>
    ''' 1. create \Log\ directory if not create before
    ''' 2. create \Log\Process\ directory if not create before
    ''' 3. checking the process log file with the same date if found it will be add the process information with the exist file
    ''' 4. create the process detail information
    ''' 5. keep the process log file from last 90 days until now otherwise will removed
    ''' </remarks>
    Public Sub WriteToProcessLog(ByVal pStartTime As Date, ByVal pEndTime As Date, ByVal pScreenName As String, ByVal pDirectory As String, _
                                 Optional ByVal pCustomMsg As String = "", _
                                 Optional ByVal pErrSummary As String = "", _
                                 Optional ByVal pErrID As Integer = 9999, _
                                 Optional ByVal pErrSeverity As ErrSeverity = ErrSeverity.INFO, _
                                 Optional ByVal pWriteToEventLog As Boolean = False, _
                                 Optional ByVal pStartEndStatus As String = "", _
                                 Optional ByVal pUseLogTime As Boolean = False, _
                                 Optional ByVal pLogID As String = "", _
                                 Optional ByVal pParameters As String = "")

        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_Duration As String
        Dim ls_DurationError As String
        Dim ls_datefolder As String
        Dim ls_CompName As String = uf_CompName()
        Dim ls_LogFolder As String = pDirectory '"D:\LogFile"

        'Dim cfg As New pisGlobal.clsConfig
        'If cfg.PISLog <> "" Then
        '    ls_LogFolder = cfg.PISLog
        'End If

        If Not System.IO.Directory.Exists("D:\") Then
            ls_LogFolder = Application.StartupPath
        End If


        li_StartTime = pStartTime
        li_EndTime = pEndTime
        li_Duration = li_EndTime - li_StartTime
        ls_Duration = uf_AddSpace(Format(li_Duration.TotalMilliseconds, "###.#0") & " (ms)", 15)
        ls_DurationError = uf_AddSpace("", 15)
        ls_datefolder = Format(Now, "yyyyMMdd")
        ls_Date = Format(Now, "yyyyMMdd")
        ls_ErrType = uf_ErrSeverity(pErrSeverity)
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")

        If Not System.IO.Directory.Exists(ls_LogFolder) Then
            System.IO.Directory.CreateDirectory(ls_LogFolder)
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder & "\Log") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log")
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder &
        "\Log\" & ls_datefolder) Then
            System.IO.Directory.CreateDirectory(ls_LogFolder &
            "\Log\" & ls_datefolder)
        End If


        If Not System.IO.Directory.Exists(ls_LogFolder &
        "\Log\" & ls_datefolder & "\Process") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder &
            "\Log\" & ls_datefolder & "\Process")
        End If

        'check the file
        Dim fs As FileStream = New FileStream(ls_LogFolder &
        "\Log\" & ls_datefolder & "\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)


        Dim s As StreamWriter = New StreamWriter(fs)

        s.Close()
        fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream(ls_LogFolder &
        "\Log\" & ls_datefolder & "\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)


        If pStartEndStatus.ToUpper.Trim = "START" Then
            If pUseLogTime = True Then
                s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
            Else
                s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            End If
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("start.")
            s1.Write("" & vbCrLf)
            s1.Write("* * * * *")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        ElseIf pStartEndStatus.ToUpper.Trim = "END" Then
            s1.Write("* * * * *")
            s1.Write("" & vbCrLf)
            If pUseLogTime = True Then
                s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
            Else
                s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            End If
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("end.")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        ElseIf pStartEndStatus = "" Then
            If ls_ErrType = "I" Then
                If pUseLogTime = True Then
                    s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
                Else
                    s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
                End If
                s1.Write("[" & UserLogin & "] ")
                s1.Write("[" & ls_CompName & "] ")
                s1.Write("" & pScreenName & " ")
                s1.Write("[" & ls_ErrType & "] ")
                s1.Write("" & ls_Duration & " ")
                s1.Write("" & pCustomMsg & " ")
                s1.Write("" & pErrSummary & "")
                s1.Write("" & vbCrLf)
                s1.Close()
                fs1.Close()
            Else
                If pUseLogTime = True Then
                    s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
                Else
                    s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
                End If
                s1.Write("[" & UserLogin & "] ")
                s1.Write("[" & ls_CompName & "] ")
                s1.Write("" & pScreenName & " ")
                s1.Write("[" & ls_ErrType & "] ")
                s1.Write("" & ls_DurationError & " ")
                s1.Write("" & pCustomMsg & " ")
                s1.Write("" & pErrSummary & "")
                s1.Write("" & vbCrLf)
                s1.Close()
                fs1.Close()
            End If
        End If
    End Sub

    '*************************************************************
    'NAME:          WriteToProcessLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    pStartTime - Start of time process of the error file entry
    '               pScreenName - Name of Screen Process
    '               pCustomMsg - Name of Process of the error file entry
    '               pErrSummary - Err summary process
    '               pErrSeverity - Enum ALERT, INFO and Err of the error file entry
    '               pErrMsg - msg of the error file entry
    'RETURNS:       Nothing
    '*************************************************************
    ''' <summary>
    ''' this procedure for process error log
    ''' </summary>
    ''' <param name="pStartTime">Start Time</param>
    ''' <param name="pScreenName">Screen Name</param>
    ''' <param name="pLogID">LogID</param>
    ''' <param name="pParameters">Parameters Message</param>
    ''' <remarks>
    ''' 1. create \Log\ directory if not create before
    ''' 2. create \Log\Process\ directory if not create before
    ''' 3. checking the process log file with the same date if found it will be add the process information with the exist file
    ''' 4. create the process detail information
    ''' 5. keep the process log file from last 90 days until now otherwise will removed
    ''' </remarks>
    Public Sub WriteToProcessLog(ByVal pLogID As String, ByVal pStartTime As Date, ByVal pScreenName As String, ByVal pDirectory As String, _
                                 Optional ByVal pParameters As String = "", _
                                 Optional ByVal pHarigami As Boolean = False)

        'check and make the directory if necessary; this is set to look in 
        'the application folder, you may wish to place the error log in 
        'another location depending upon the user's role and write access to 
        'different areas of the file system
        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_Duration As String
        Dim ls_CustomMsg As String
        Dim li_EventID As Integer
        Dim ls_LogFolder As String = "D:\Log"
        Dim Message As String
        Dim MsgFound As Boolean = False
        Dim ls_DateFolder As String = ""
        Dim ls_CompName As String = uf_CompName()
        '========================================================================
        'Message
        '========================================================================

        'Dim cfg As New pisGlobal.clsConfig
        'If cfg.PISLog <> "" Then
        '    ls_LogFolder = cfg.PISLog
        'End If
        If Not System.IO.Directory.Exists("D:\") Then
            ls_LogFolder = Application.StartupPath
        End If

        Message = uf_ConvertMsg(pLogID)

        If Message = "" Then
            Message = pLogID
        Else
            MsgFound = True
        End If
        Dim i As Integer, Position As Long
        Dim Parameters() As String
        Parameters = Split(pParameters, "|")
        If UBound(Parameters) <> -1 Then
            Position = InStr(1, Message, "%%")
            Do While Position > 0
                Message = Left(Message, Position - 1) & Parameters(i) & Mid(Message, Position + 2, Len(Message) - Position)
                Position = InStr(1, Message, "%%")
                i = i + 1
            Loop
        Else
            Message = Replace(Message, "%", "")
        End If
        If MsgFound Then
            ls_CustomMsg = Message
        Else
            ls_CustomMsg = Message
        End If

        li_EventID = uf_ConvertEventID(pLogID)

        li_StartTime = pStartTime
        li_EndTime = Now
        li_Duration = li_EndTime - li_StartTime
        ls_Duration = uf_AddSpace(Format(li_Duration.TotalMilliseconds, "###.#0") & " (ms)", 15)

        ls_Date = Format(Now, "yyyyMMdd")
        ls_DateFolder = ls_Date
        ls_ErrType = uf_ErrSeverity(ErrSeverity.ERR)
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")


        If pHarigami Then
            If Not System.IO.Directory.Exists(ls_LogFolder) Then
                System.IO.Directory.CreateDirectory(ls_LogFolder)
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder & "\Log") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log")
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder &
                    "\Log\" & ls_DateFolder) Then
                System.IO.Directory.CreateDirectory(ls_LogFolder &
                "\Log\" & ls_DateFolder)
            End If


            If Not System.IO.Directory.Exists(ls_LogFolder &
                    "\Log\" & ls_DateFolder & "\Process") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder &
                "\Log\" & ls_DateFolder & "\Process")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(ls_LogFolder &
            "\Log\" & ls_DateFolder & "\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)


            Dim s As StreamWriter = New StreamWriter(fs)

            s.Close()
            fs.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(ls_LogFolder &
            "\Log\" & ls_DateFolder & "\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)

            s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("" & ls_CustomMsg.Trim & "")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        Else
            If Not System.IO.Directory.Exists(ls_LogFolder) Then
                System.IO.Directory.CreateDirectory(ls_LogFolder)
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder &
            "\Log\Process") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder &
                "\Log\Process")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(ls_LogFolder &
            "\Log\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)
            Dim s As StreamWriter = New StreamWriter(fs)

            s.Close()
            fs.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(ls_LogFolder &
            "\Log\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)

            s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("" & ls_CustomMsg.Trim & "")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        End If

        If Not pHarigami Then
            Dim ls_Dir As New IO.DirectoryInfo(ls_LogFolder & "\Log\Process\")
            Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
            Dim ls_File As IO.FileInfo
            Dim li_index As Long
            Dim ls_log As String
            Dim li_CountDate As Long
            Dim ls_Temp As List(Of String) = New List(Of String)

            li_index = 0
            For Each ls_File In ls_GetFile

                ls_Temp.Add(ls_File.ToString)

                ls_log = Right(ls_Temp.Item(li_index), 12)
                ls_log = Mid(ls_log, 1, 8)
                ls_log = Left(ls_log, 4) & "/" & Mid(ls_log, 5, 2) & "/" & Right(ls_log, 2)
                li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_log), "yyyy/MM/dd")), CDate(Format(Now, "yyyy/MM/dd")))
                If li_CountDate > 90 Then
                    File.Delete(ls_LogFolder & "\Log\Process\prc" & pScreenName & "_" & Format(CDate(ls_log), "yyyyMMdd") & ".log")
                End If

                li_index = li_index + 1
            Next
        End If

    End Sub

    Private Function uf_ErrSeverity(ByVal pErrSeverity As ErrSeverity) As String
        If pErrSeverity = ErrSeverity.ALERT Then
            Return "A"
        ElseIf pErrSeverity = ErrSeverity.ERR Then
            Return "E"
        Else
            Return "I"
        End If
    End Function

    Private Function uf_AddSpace(ByVal pDuration As String, ByVal pSpace As Integer) As String
        Return Space(pSpace - pDuration.Length) & pDuration
    End Function

    Private Function uf_CompName() As String
        Return My.Computer.Name.ToString
    End Function

#End Region

End Class
