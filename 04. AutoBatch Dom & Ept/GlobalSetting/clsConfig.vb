Imports System.Data.SqlClient
Imports System.Xml

Public Class clsConfig

#Region "Declaration"

    Private builder As SqlConnectionStringBuilder
    Private ls_path As String

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

#End Region

#Region "Properties"
    Public Property Server As String
    Public Property Database As String
    Public Property User As String
    Public Property Password As String
    Public Property CommandTimeout As Integer
    Public Property ConnectionString As String
#End Region

#Region "Function"
    Public Function AddSlash(ByVal Path As String) As String
        Dim Result As String = Path
        If Path.EndsWith("\") = False Then
            Result = Result + "\"
        End If
        Return Result
    End Function
#End Region

#Region "Initialization"
    ''' <summary>
    ''' Open config file and store the value in local variables.
    ''' </summary>
    ''' <param name="pConfigFile"></param>
    ''' <remarks></remarks>
    Public Sub New(Optional ByVal pConfigFile As String = "UPLOAD.ini")
        Dim ret As String
        Dim ret1 As String
        Dim ret2 As String
        Dim ret3 As String
        Dim ret4 As String
        Dim ret5 As String
        Dim lng As Long

        ls_path = AddSlash(My.Application.Info.DirectoryPath) & pConfigFile

        If Not My.Computer.FileSystem.FileExists(ls_path) Then
            Throw New Exception("Config file is not found")
        End If

        ret = Space(1500)
        ret1 = Space(1500)
        ret2 = Space(1500)
        ret3 = Space(1500)
        ret4 = Space(1500)
        ret5 = Space(3000)

        '******************************************************************
        ' SETTING SERVER
        '******************************************************************
        'server
        lng = GetPrivateProfileString("Setting", "server", "", _
                ret1, 1500, ls_path)
        If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret1, CInt(lng))
        Server = ret

        'database
        lng = GetPrivateProfileString("Setting", "DatabaseName", "", _
        ret2, 1500, ls_path)
        If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret2, CInt(lng))
        Database = ret

        'User
        lng = GetPrivateProfileString("Setting", "User", "", _
        ret3, 1500, ls_path)
        If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret3, CInt(lng))
        User = ret

        'Password
        lng = GetPrivateProfileString("Setting", "Password", "", _
        ret4, 1500, ls_path)
        If lng <> 0 Then ret = Microsoft.VisualBasic.Left(ret4, CInt(lng))
        Password = ret

        CommandTimeout = 60

        builder = New SqlConnectionStringBuilder
        builder.DataSource = Server
        builder.InitialCatalog = Database
        builder.UserID = User
        builder.Password = Password
        builder.ConnectTimeout = CommandTimeout
        builder.ApplicationName = "Autobatch Scheduler Send To Supplier"

        ConnectionString = builder.ConnectionString

    End Sub
#End Region

End Class
