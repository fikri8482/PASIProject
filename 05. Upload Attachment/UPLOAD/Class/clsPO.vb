Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsPO
    Shared Sub up_SendPODomestic(ByVal EF As Excel.Worksheet, ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")
        Dim xxxx As String

        If EF.Range("I16").Value Is Nothing Then
            xxxx = ""
        Else
            xxxx = EF.Range("I16").Value
        End If
    End Sub
End Class
