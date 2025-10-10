Imports System.Data

Imports System.Net
Imports System.IO

Imports System.Windows.Forms
Imports System.Reflection
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Transactions

Module MdlConn

    Public IsConnection As Boolean = False
    Public CheckConfigStatus As String

    Public gs_DBserver As String
    Public gs_DBdatabase As String
    Public gs_DBuser As String
    Public gs_DBpass As String
    Public gs_DBwinmode As String
    Public gs_Path As String
    Public gs_AppName As String
    Public gi_commandTimeOut As Integer
    Public gi_connectionTimeOut As Integer
    Public lts_Transaction As TimeSpan

    Public sql As String
    Public dt As DataTable
    Public ds As DataSet
    Public cmd As SqlCommand
    Public da As SqlClient.SqlDataAdapter
    Public i As Integer
    Public gs_errMessage As String
    Public Result_Update As Long

    Public gs_InquiryPeriod As String = ""
    Public Event Tick As EventHandler
    Public statusSave As Boolean = True
    Public MsgstatusSave As String = ""
    Public statusDNSupp As Boolean = True
    '*************************************************************
    Public Declare Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, ByVal lpKeyName As String, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Integer, ByVal lpFileName As String) As _
    Integer

    Public Declare Function WritePrivateProfileString Lib _
    "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, ByVal lpKeyName As String, _
    ByVal lpString As String, ByVal lpFileName As String) As _
    Long

    'Public Assem As Reflection.Assembly
    Public gb_CopyDataFromInquiry As Boolean = False
    Public gb_ReverseDataFromInquiry As Boolean = False
    Public gb_TransDataFromInquiry As Boolean = False

    '*****************************************************************************************************

    Public Function ReadConnection() As String
        'Assem = Me.GetType.Assembly
        '*********************************************************************
        'GetConString
        Dim settingReader As New AppSettingsReader

        Dim ret As String
        Dim ret1 As String
        Dim ret2 As String
        Dim ret3 As String
        Dim ret4 As String
        Dim ret5 As String
        Dim ls_path As String

        Dim lng As Long

        Try
            IsConnection = True
            ret = Space(1500)
            ret1 = Space(1500)
            ret2 = Space(1500)
            ret3 = Space(1500)
            ret4 = Space(1500)
            ret5 = Space(3000)
            ls_path = My.Application.Info.DirectoryPath & "\UPLOAD.ini"

            '******************************************************************
            ' SETTING SERVER
            '******************************************************************
            'server
            lng = GetPrivateProfileString("Setting", "server", "", _
                    ret1, 1500, ls_path)
            If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret1, CInt(lng))
            gs_DBserver = ret

            'database
            lng = GetPrivateProfileString("Setting", "DatabaseName", "", _
            ret2, 1500, ls_path)
            If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret2, CInt(lng))
            gs_DBdatabase = ret

            'User
            lng = GetPrivateProfileString("Setting", "User", "", _
            ret3, 1500, ls_path)
            If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret3, CInt(lng))
            gs_DBuser = ret

            'Password
            lng = GetPrivateProfileString("Setting", "Password", "", _
            ret4, 1500, ls_path)
            If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret4, CInt(lng))
            gs_DBpass = ret

            'PATH
            lng = GetPrivateProfileString("Setting", "Path", "", _
            ret5, 3000, ls_path)
            If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret5, CInt(lng))
            gs_Path = ret

            'ApplicationName
            lng = GetPrivateProfileString("Setting", "AppName", "", _
            ret5, 3000, ls_path)
            If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret5, CInt(lng))
            gs_AppName = ret

            gi_commandTimeOut = 60

            'Check Connection
            '1. FORMINGENDDATA
            Dim conn As New SqlConnection("Data Source=" & gs_DBserver & ";Initial Catalog=" & gs_DBdatabase & ";User ID=" & gs_DBuser & "" & ";pwd=" & gs_DBpass & "" & ";Application Name=" & gs_AppName & ";Connection Timeout=" & gi_connectionTimeOut & "" & "")
            conn.Open()
            conn.Close()
            'Check Connection

        Catch ex As Exception
            IsConnection = False
            'CREATE ERROR LOG
            Dim FilePath As String
            Dim FileName As String
            Dim LastModified As Date
            Dim st As New System.Diagnostics.StackTrace(True)

            FilePath = My.Application.Info.DirectoryPath & "\Log\Error\"
            FileName = "Upload" & Format(Now, "dd") & ".Log"

            If (Not System.IO.Directory.Exists(FilePath)) Then
                System.IO.Directory.CreateDirectory(FilePath)
            End If

            If Not System.IO.File.Exists(FilePath & FileName) Then
                System.IO.File.Create(FilePath & FileName).Dispose()
                Dim sw As New System.IO.StreamWriter(FilePath & FileName, True)
                sw.WriteLine(Format(Now, "yyyy/MM/dd HH:mm:ss") + " [PC NAME] " + My.Computer.Name + " [Content] " + st.GetFrame(0).GetMethod.DeclaringType.FullName.ToString + ":" + st.GetFrame(0).GetMethod.Name.ToString + " [Line] " + st.GetFrame(0).GetFileLineNumber.ToString)
                sw.WriteLine("Err Number : " + CStr(Err.Number) + ", Err Description : " + Err.Description)
                sw.WriteLine()
                sw.Close()
            Else
                LastModified = System.IO.File.GetLastAccessTime(FilePath & FileName)
                If Format(LastModified, "MM") = Format(Now, "MM") Then
                    Dim sw As New System.IO.StreamWriter(FilePath & FileName, True)
                    sw.WriteLine(Format(Now, "yyyy/MM/dd HH:mm:ss") + " [PC NAME] " + My.Computer.Name + " [Content] " + st.GetFrame(0).GetMethod.DeclaringType.FullName.ToString + ":" + st.GetFrame(0).GetMethod.Name.ToString + " [Line] " + st.GetFrame(0).GetFileLineNumber.ToString)
                    sw.WriteLine("Err Number : " + CStr(Err.Number) + ", Err Description : " + Err.Description)
                    sw.WriteLine()
                    sw.Close()
                Else
                    System.IO.File.Delete(FilePath & FileName)
                    System.IO.File.Create(FilePath & FileName).Dispose()
                    Dim sw As New System.IO.StreamWriter(FilePath & FileName, True)
                    sw.WriteLine(Format(Now, "yyyy/MM/dd HH:mm:ss") + " [PC NAME] " + My.Computer.Name + " [Content] " + st.GetFrame(0).GetMethod.DeclaringType.FullName.ToString + ":" + st.GetFrame(0).GetMethod.Name.ToString + " [Line] " + st.GetFrame(0).GetFileLineNumber.ToString)
                    sw.WriteLine("Err Number : " + CStr(Err.Number) + ", Err Description : " + Err.Description)
                    sw.WriteLine()
                    sw.Close()
                End If

            End If
            'CREATE ERROR LOG
        End Try
    End Function

    Public Function uf_GetDataSet(ByVal ls_query As String, Optional ByVal DbLock As SqlClient.SqlConnection = Nothing) As DataSet
        Dim lcon As New SqlConnection, lds As New DataSet
        gs_errMessage = ""

        Try
            If DbLock Is Nothing Then
                lcon.ConnectionString = uf_GetConString()

                lcon.Open()
                Dim lda As New SqlDataAdapter(ls_query, lcon)
                lda.Fill(lds)
                lcon.Close()
            Else
                Dim lda As New SqlDataAdapter(ls_query, DbLock)
                lda.Fill(lds)
                lcon.Close()

            End If

        Catch ex As SqlException
            ex.Message.ToString()
        End Try

        Return lds
    End Function

    Public Function uf_GetDatatabel(ByVal ls_query As String, Optional ByVal DbLock As SqlClient.SqlConnection = Nothing) As DataTable
        Dim lcon As New SqlConnection, ldt As New DataTable
        gs_errMessage = ""

        Try
            If DbLock Is Nothing Then
                lcon.ConnectionString = uf_GetConString()

                lcon.Open()
                Dim lda As New SqlDataAdapter(ls_query, lcon)
                lda.Fill(ldt)
                lcon.Close()
            Else
                Dim lda As New SqlDataAdapter(ls_query, DbLock)
                lda.Fill(ldt)
                lcon.Close()

            End If


        Catch ex As SqlException
            ex.Message.ToString()
        End Try

        Return ldt
    End Function


    Public Function uf_ExecuteSql(ByVal ls_query As String, Optional ByVal DbLock As SqlClient.SqlConnection = Nothing) As String
        Dim lResult As Long
        Err.Description = ""
        Dim cmd As New SqlCommand
        Try
            If DbLock Is Nothing Then
                Dim con As New SqlConnection
                con.ConnectionString = uf_GetConString()
                con.Open()
                cmd = New SqlCommand(ls_query, con)
                lResult = cmd.ExecuteNonQuery()
                Result_Update = lResult
                con.Close()
            Else
                cmd = New SqlCommand(ls_query, DbLock)
                lResult = cmd.ExecuteNonQuery()
                Result_Update = lResult
            End If

        Catch ex As Exception
            Return ex.Message.ToString
            Exit Function
        End Try


        Return ""
        Exit Function
fail:
        Return Err.Description

    End Function

    Public Function uf_GetConString() As String
        'If gs_DBWinmode = "mixed" Then
        Return "Data Source=" & gs_DBserver & ";Initial Catalog=" & gs_DBdatabase & ";User ID=" & gs_DBuser & ";pwd=" & gs_DBpass & ";Application Name=" & gs_AppName & ""
        'Else
        'Return "Data Source=" & gs_DBserver & ";Initial Catalog=" & gs_DBdatabase & ";Integrated Security=True"
        'End If
    End Function
End Module
