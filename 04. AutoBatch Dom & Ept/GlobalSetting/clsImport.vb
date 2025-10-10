'########################################################################################################################################
'System         : ADM Karawang Assyplant Production Instruction System
'Program        : History File Import and Export
'Overview       : This program about
'                 for Harigami Item Master        
'                 1. Copy File Import
'                 2. Copy File Export  
'Parameter Input: Name File, Screen Nam
'Created By     : Defrizal
'Created Date   : 29 Feb 2012
'Last Update By :
'Last Update    :
'Modify Update  ([Date],[Editor],[Summary],[Version])
'########################################################################################################################################
Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics
Public Class ClsImport
    ''' <summary>
    ''' this procedure for Copy File Import
    ''' </summary>
    ''' <param name="namefile">NameFile</param>
    ''' <param name="pScreenName">ScreenName</param>
    ''' <remarks>
    ''' 1. create \History\ directory if not create before
    ''' 2. create \History\Import\ directory if not create before
    ''' 3. create the ScreenName detail information
    ''' 4. Copy NameFile and change nameFile to ScreenName
    ''' 5. keep the Copy file Import from last 30 days until now otherwise will removed
    ''' </remarks>
    Public Sub Copy_fileImport(ByVal namefile As String, ByVal pScreenName As String, ByVal pDirectory As String)
        Dim filecopy As String
        Dim exs As String
        Dim ls_date As String
        Dim rename As String
        ls_date = Format(Now, "yyyyMMdd hhmmss")

        'cek and create Directory if not exists
        If Not System.IO.Directory.Exists(pDirectory & "\Archived\") Then
            System.IO.Directory.CreateDirectory(pDirectory & "\Archived\")
        End If

        'cek and create Directory if not exists
        'If Not System.IO.Directory.Exists(pDirectory & "\Archived\Import\") Then
        '    System.IO.Directory.CreateDirectory(pDirectory & "\Archived\Import\")
        'End If

        filecopy = namefile
        exs = Right(namefile, 4)
        If exs = "xlsx" Then
            rename = pScreenName & "_" & ls_date & "." & exs
        Else
            rename = pScreenName & "_" & ls_date '& exs
        End If


        'Create new Name file 
        Dim NewLocation As String = pDirectory & "\Archived\" & rename & ""
        Dim folderLocation As String = pDirectory & "\Archived\"

        'Copy FileCopy  to New Location
        If (System.IO.Directory.Exists(folderLocation)) Then
            If (System.IO.File.Exists(filecopy)) Then
                System.IO.File.Move(filecopy, NewLocation)
                'System.IO.File.Copy(filecopy, NewLocation, True)
            End If
        End If


        'Dim ls_Dir As New IO.DirectoryInfo(Application.StartupPath & "\Archived\Import\")
        'Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
        'Dim ls_File As IO.FileInfo
        'Dim li_index As Long
        'Dim ls_import As String
        'Dim ls_import2 As Date
        'Dim li_CountDate As Long
        'Dim ls_Temp As List(Of String) = New List(Of String)
        'Dim ls_namefile As String
        'Dim exsdel As String


        'li_index = 0
        'For Each ls_File In ls_GetFile
        '    'keep file Import for 30 days
        '    ls_Temp.Add(ls_File.ToString)
        '    exsdel = Right(ls_File.Name, 4)
        '    If exsdel = "xlsx" Then
        '        ls_import = Right(ls_Temp.Item(li_index), 20)
        '        ls_import = Mid(ls_import, 1, 8)
        '        ls_import = Mid(ls_import, 5, 2) & "/" & Right(ls_import, 2) & "/" & Left(ls_import, 4)
        '        ls_import2 = New Date(Right(ls_import, 4), Left(ls_import, 2), Mid(ls_import, 4, 2))
        '        'li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_import), "MM/dd/yyyy")), CDate(Format(Now, "MM/dd/yyyy")))
        '        li_CountDate = DateDiff(DateInterval.Day, ls_import2, Now)
        '        ls_namefile = Left(ls_File.Name, Len(ls_File.Name) - 21)
        '    Else
        '        ls_import = Right(ls_Temp.Item(li_index), 19)
        '        ls_import = Mid(ls_import, 1, 8)
        '        ls_import = Mid(ls_import, 5, 2) & "/" & Right(ls_import, 2) & "/" & Left(ls_import, 4)
        '        ls_import2 = New Date(Right(ls_import, 4), Left(ls_import, 2), Mid(ls_import, 4, 2))
        '        'li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_import), "MM/dd/yyyy")), CDate(Format(Now, "MM/dd/yyyy")))
        '        li_CountDate = DateDiff(DateInterval.Day, ls_import2, Now)
        '        ls_namefile = Left(ls_File.Name, Len(ls_File.Name) - 20)
        '    End If

        '    If li_CountDate > 30 Then
        '        If pScreenName = ls_namefile Then
        '            File.Delete(ls_File.FullName)
        '        End If
        '    End If

        '    li_index = li_index + 1
        'Next
    End Sub

    Public Sub Copy_fileImportAutomatic(ByVal namefile As String, ByVal pScreenName As String)
        Dim filecopy As String
        Dim exs As String
        Dim ls_date As String
        Dim rename As String
        ls_date = Format(Now, "yyyyMMdd hhmmss")

        'cek and create Directory if not exists
        If Not System.IO.Directory.Exists(Application.StartupPath & "\Archived\") Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Archived\")
        End If

        'cek and create Directory if not exists
        If Not System.IO.Directory.Exists(Application.StartupPath & "\Archived\Import\") Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Archived\Import\")
        End If

        filecopy = namefile
        exs = Right(namefile, 4)
        If exs = "xlsx" Then
            rename = pScreenName & "_" & ls_date & "." & exs
        Else
            rename = pScreenName & "_" & ls_date & exs
        End If


        'Create new Name file 
        Dim NewLocation As String = Application.StartupPath & "\Archived\Import\" & rename & ""
        Dim folderLocation As String = Application.StartupPath & "\Archived\Import\"

        'Copy FileCopy  to New Location
        If (System.IO.Directory.Exists(folderLocation)) Then
            If (System.IO.File.Exists(filecopy)) Then
                'System.IO.File.Move(filecopy, NewLocation)
                System.IO.File.Copy(filecopy, NewLocation, True)
            End If
        End If


        'Dim ls_Dir As New IO.DirectoryInfo(Application.StartupPath & "\Archived\Import\")
        'Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
        'Dim ls_File As IO.FileInfo
        'Dim li_index As Long
        'Dim ls_import As String
        'Dim ls_import2 As Date
        'Dim li_CountDate As Long
        'Dim ls_Temp As List(Of String) = New List(Of String)
        'Dim ls_namefile As String
        'Dim exsdel As String


        'li_index = 0
        'For Each ls_File In ls_GetFile
        '    'keep file Import for 30 days
        '    ls_Temp.Add(ls_File.ToString)
        '    exsdel = Right(ls_File.Name, 4)
        '    If exsdel = "xlsx" Then
        '        ls_import = Right(ls_Temp.Item(li_index), 20)
        '        ls_import = Mid(ls_import, 1, 8)
        '        ls_import = Mid(ls_import, 5, 2) & "/" & Right(ls_import, 2) & "/" & Left(ls_import, 4)
        '        ls_import2 = New Date(Right(ls_import, 4), Left(ls_import, 2), Mid(ls_import, 4, 2))
        '        'li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_import), "MM/dd/yyyy")), CDate(Format(Now, "MM/dd/yyyy")))
        '        li_CountDate = DateDiff(DateInterval.Day, ls_import2, Now)
        '        ls_namefile = Left(ls_File.Name, Len(ls_File.Name) - 21)
        '    Else
        '        ls_import = Right(ls_Temp.Item(li_index), 19)
        '        ls_import = Mid(ls_import, 1, 8)
        '        ls_import = Mid(ls_import, 5, 2) & "/" & Right(ls_import, 2) & "/" & Left(ls_import, 4)
        '        ls_import2 = New Date(Right(ls_import, 4), Left(ls_import, 2), Mid(ls_import, 4, 2))
        '        'li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_import), "MM/dd/yyyy")), CDate(Format(Now, "MM/dd/yyyy")))
        '        li_CountDate = DateDiff(DateInterval.Day, ls_import2, Now)
        '        ls_namefile = Left(ls_File.Name, Len(ls_File.Name) - 20)
        '    End If

        '    If li_CountDate > 30 Then
        '        If pScreenName = ls_namefile Then
        '            File.Delete(ls_File.FullName)
        '        End If
        '    End If

        '    li_index = li_index + 1
        'Next
    End Sub

    ''' <summary>
    ''' this procedure for Copy File Export
    ''' </summary>
    ''' <param name="namefile">NameFile</param>
    ''' <param name="pScreenName">ScreenName</param>
    ''' <remarks>
    ''' 1. create \History\ directory if not create before
    ''' 2. create \History\Export\ directory if not create before
    ''' 3. create the ScreenName detail information
    ''' 4. Copy NameFile and change nameFile to ScreenName
    ''' 5. keep the Copy file Export from last 30 days until now otherwise will removed
    ''' </remarks>
    Public Sub Copy_fileExport(ByVal namefile As String, ByVal pScreenName As String)
        Dim filecopy As String
        Dim ls_date As String
        ls_date = Format(Now, "yyyyMMdd_hhmmss")

        'cek and create Directory if not exists
        If Not System.IO.Directory.Exists(Application.StartupPath & "\Archived\") Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Archived\")
        End If

        'cek and create Directory if not exists
        If Not System.IO.Directory.Exists(Application.StartupPath & "\Archived\Export\") Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Archived\Export\")
        End If

        'Create new Name file 
        Dim NewLocation As String = Application.StartupPath & "\Archived\Export\" & pScreenName & ""
        Dim folderLocation As String = Application.StartupPath & "\Archived\Export\"

        filecopy = namefile
        'Copy FileCopy  to New Location
        If (System.IO.Directory.Exists(folderLocation)) Then
            If (System.IO.File.Exists(filecopy)) Then
                System.IO.File.Copy(filecopy, NewLocation, True)
            End If
        End If

        Dim ls_Dir As New IO.DirectoryInfo(Application.StartupPath & "\Archived\Export\")
        Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
        Dim ls_File As IO.FileInfo
        Dim li_index As Long
        Dim ls_import As String
        Dim li_CountDate As Long
        Dim ls_Temp As List(Of String) = New List(Of String)


        li_index = 0
        For Each ls_File In ls_GetFile
            'keep file Export for 30 days
            ls_Temp.Add(ls_File.ToString)

            ls_import = Left(ls_Temp.Item(li_index), 17)
            ls_import = Mid(ls_import, 10, 8)
            ls_import = Mid(ls_import, 1, 4) & "/" & Mid(ls_import, 5, 2) & "/" & Right(ls_import, 2)
            li_CountDate = DateDiff(DateInterval.Day, CDate(ls_import), CDate(Format(Now, "yyyy/MM/dd")))
            If li_CountDate > 30 Then
                File.Delete(ls_File.FullName)
            End If

            li_index = li_index + 1
        Next
    End Sub

End Class
