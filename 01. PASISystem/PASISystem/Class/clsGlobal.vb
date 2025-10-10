Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.IO
Imports System.Web
Imports System.Drawing
Imports DevExpress.Web


Public Class clsGlobal
    Inherits System.Web.UI.Page

#Region "Declaration"
    Public Declare Function uf_GetPrivateProfileString Lib _
        "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As String, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Integer, ByVal lpFileName As String) As _
        Integer

    Public Declare Function uf_WritePrivateProfileString Lib _
        "kernel32" Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As String, _
        ByVal lpString As String, ByVal lpFileName As String) As _
        Long

    Dim builder As SqlConnectionStringBuilder
    Public ConnectionString As String = ""

    Public Enum pSet
        SetDefault = True
        None = False
    End Enum

    'Color
    Public Const ColorGridHeader As Single = 49407
    Public Const gs_All As String = "== ALL =="
#End Region

#Region "Function"
    Public Function GetCompanyName() As String
        Dim retValue As String = ""

        Try
            Using sqlConn As New SqlConnection(ConnectionString)
                sqlConn.Open()

                Dim sqlDA As New SqlDataAdapter("SELECT TOP 1 Company_Name FROM dbo.Company_Profile", sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)
                If ds.Tables(0).Rows.Count > 0 Then
                    retValue = Trim(ds.Tables(0).Rows(0).Item("Company_Name").ToString())
                End If
            End Using
        Catch ex As Exception
            retValue = ""
        End Try

        Return retValue
    End Function

    Public Function FormatQtyBOM(ByVal pQtyBOM As String) As String
        Dim ls_Result As String = "", lb_Dot As Boolean = False, li_Ctr As Byte = 1
        Dim ls_QtyBOM As String = pQtyBOM

        If InStr(ls_QtyBOM, ".") = 0 Then
            ls_QtyBOM = ls_QtyBOM & ".00000"
        Else
            ls_QtyBOM = FormatNumber(pQtyBOM, 5, TriState.True, TriState.False, TriState.False)
        End If

        For iRow As Integer = Len(ls_QtyBOM) To 1 Step -1
            If lb_Dot = True Then 'Already pass decimal -> 999,999[.]99999
                If li_Ctr < 4 Then
                    ls_Result = Mid(ls_QtyBOM, iRow, 1) & ls_Result
                    li_Ctr = li_Ctr + 1
                Else
                    ls_Result = Mid(ls_QtyBOM, iRow, 1) & "," & ls_Result
                    li_Ctr = 2
                End If
            Else
                If Mid(ls_QtyBOM, iRow, 1) = "." Then
                    lb_Dot = True
                End If
                ls_Result = Mid(ls_QtyBOM, iRow, 1) & ls_Result
            End If
        Next iRow

        Return ls_Result
        'Return FormatNumber(pQtyBOM, 5, TriState.True)
    End Function

    Public Function FormatQty(ByVal pQty As String) As String
        Dim ls_Result As String = "", lb_Dot As Boolean = False, li_Ctr As Byte = 1
        Dim ls_Qty As String = pQty

        Return FormatNumber(pQty, 2, TriState.True)
    End Function

    Public Function FormatTime(ByVal pQty As String) As String
        Dim ls_Result As String = "", lb_Dot As Boolean = False, li_Ctr As Byte = 1
        Dim ls_Qty As String = pQty

        If InStr(ls_Qty, ".") = 0 Then
            ls_Qty = ls_Qty & ".00"
        Else
            ls_Qty = FormatNumber(pQty, 2, TriState.True, TriState.False, TriState.False)
        End If

        For iRow As Integer = Len(ls_Qty) To 1 Step -1
            If lb_Dot = True Then 'Already pass decimal -> 999,999[.]99
                If li_Ctr < 4 Then
                    ls_Result = Mid(ls_Qty, iRow, 1) & ls_Result
                    li_Ctr = li_Ctr + 1
                Else
                    ls_Result = Mid(ls_Qty, iRow, 1) & "," & ls_Result
                    li_Ctr = 2
                End If
            Else
                If Mid(ls_Qty, iRow, 1) = "." Then
                    lb_Dot = True
                End If
                ls_Result = Mid(ls_Qty, iRow, 1) & ls_Result
            End If
        Next iRow

        Return ls_Result
    End Function

    Public Function FormatPrice(ByVal pPrice As String) As String
        Dim ls_Result As String = "", lb_Dot As Boolean = False, li_Ctr As Byte = 1
        Dim ls_Price As String = pPrice

        Return FormatNumber(ls_Price, 4, TriState.True)
    End Function

    Public Function FormatAmount(ByVal pAmount As String) As String
        Dim ls_Result As String = "", lb_Dot As Boolean = False, li_Ctr As Byte = 1
        Dim ls_Amount As String = pAmount

        If pAmount = "" Then pAmount = 0
        Return FormatNumber(pAmount, 2, TriState.True)
    End Function

    Public Function FormatRupiah(ByVal pRupiah As String) As String
        Dim ls_Result As String = "", lb_Dot As Boolean = False, li_Ctr As Byte = 1
        Dim ls_Rupiah As String = pRupiah

        If InStr(ls_Rupiah, ".") = 0 Then
            ls_Rupiah = ls_Rupiah & ".00"
        Else
            If CDbl(pRupiah) < 9999999999999999 Then
                ls_Rupiah = FormatNumber(pRupiah, 2, TriState.True, TriState.False, TriState.False)
            End If
        End If

        For iRow As Integer = Len(ls_Rupiah) To 1 Step -1
            If lb_Dot = True Then 'Already pass decimal -> 999,999[.]99
                If li_Ctr < 4 Then
                    ls_Result = Mid(ls_Rupiah, iRow, 1) & ls_Result
                    li_Ctr = li_Ctr + 1
                Else
                    ls_Result = Mid(ls_Rupiah, iRow, 1) & "," & ls_Result
                    li_Ctr = 2
                End If
            Else
                If Mid(ls_Rupiah, iRow, 1) = "." Then
                    lb_Dot = True
                End If
                ls_Result = Mid(ls_Rupiah, iRow, 1) & ls_Result
            End If
        Next iRow

        Return ls_Result
    End Function

    Public Function AddSlash(ByVal Path As String) As String
        Dim Result As String = Path
        If Path.EndsWith("\") = False Then
            Result = Result + "\"
        End If
        Return Result
    End Function

    Public Function GetUrl(ByVal pMenuName As String) As String
        Dim ls_SQL As String = ""

        Using SqlConn As New SqlConnection(ConnectionString)
            SqlConn.Open()

            ls_SQL = "SELECT MenuName FROM dbo.SC_UserMenu WHERE MenuDesc = '" & Trim(pMenuName) & "' and PASIMenu in ('1','2')"

            Dim sqlCommand As New SqlCommand(ls_SQL, SqlConn)
            Dim sqlRdr As SqlDataReader
            sqlRdr = sqlCommand.ExecuteReader()
            If sqlRdr.Read() Then
                GetUrl = Replace(sqlRdr("MenuName").ToString.Trim, " ", "")
            Else
                GetUrl = "~/MainMenu.aspx"
            End If
        End Using
    End Function

    Public Function GetMenuID(ByVal pMenuDesc As String) As String
        Dim ls_SQL As String = "SELECT MenuID FROM dbo.SC_UserMenu WHERE MenuDesc = '" & Trim(pMenuDesc) & "'"

        Using sqlConn As New SqlConnection(ConnectionString)
            sqlConn.Open()

            Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
            Dim sqlRdr As SqlDataReader = sqlCmd.ExecuteReader()

            If sqlRdr.Read() Then
                Return sqlRdr("MenuID").ToString.Trim
            Else
                Return ""
            End If

            sqlRdr.Close()
            sqlCmd.Dispose()
            sqlConn.Close()
        End Using
    End Function

    Public Function uf_GetShortMonth(ByVal pMonthName As String) As String
        Dim ls_RetValue As String = ""

        Select Case pMonthName
            Case "Jan"
                ls_RetValue = "01"
            Case "Feb"
                ls_RetValue = "02"
            Case "Mar"
                ls_RetValue = "03"
            Case "Apr"
                ls_RetValue = "04"
            Case "May"
                ls_RetValue = "05"
            Case "Jun"
                ls_RetValue = "06"
            Case "Jul"
                ls_RetValue = "07"
            Case "Aug"
                ls_RetValue = "08"
            Case "Sep"
                ls_RetValue = "09"
            Case "Oct"
                ls_RetValue = "10"
            Case "Nov"
                ls_RetValue = "11"
            Case "Dec"
                ls_RetValue = "12"
        End Select

        Return ls_RetValue
    End Function

    Public Function uf_GetMediumMonth(ByVal pMonthNumber As String) As String
        Dim ls_RetValue As String = ""

        Select Case pMonthNumber
            Case "01"
                ls_RetValue = "Jan"
            Case "02"
                ls_RetValue = "Feb"
            Case "03"
                ls_RetValue = "Mar"
            Case "04"
                ls_RetValue = "Apr"
            Case "05"
                ls_RetValue = "May"
            Case "06"
                ls_RetValue = "Jun"
            Case "07"
                ls_RetValue = "Jul"
            Case "08"
                ls_RetValue = "Aug"
            Case "09"
                ls_RetValue = "Sep"
            Case "10"
                ls_RetValue = "Oct"
            Case "11"
                ls_RetValue = "Nov"
            Case "12"
                ls_RetValue = "Dec"
        End Select

        Return ls_RetValue
    End Function

    Public Function GetServerDate() As Date
        Dim ls_SQL As String = "SELECT ServerDate = GETDATE()", retValue As Date = Now

        Dim sqlDA As New SqlDataAdapter(ls_SQL, ConnectionString)
        Dim ds As New DataSet
        sqlDA.Fill(ds)

        If ds.Tables(0).Rows.Count > 0 Then
            retValue = ds.Tables(0).Rows(0).Item("ServerDate")
        End If

        Return retValue
    End Function

    Public Function Auth_UserUpdate(ByVal pUserID As String, ByVal pMenuID As String) As Boolean
        Dim retVal As Boolean = False

        Using sqlConn As New SqlConnection(ConnectionString)
            sqlConn.Open()

            Dim ls_SQL As String = "SELECT AllowUpdate FROM dbo.SC_UserPrivilege WHERE AppID = 'P01' AND UserID = '" & Trim(pUserID) & "' AND MenuID = '" & Trim(pMenuID) & "' AND UserCls = '1'"
            Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
            Dim sqlRdr As SqlDataReader = sqlCmd.ExecuteReader()

            If sqlRdr.Read() Then
                If sqlRdr("AllowUpdate") = "1" Then
                    retVal = True
                ElseIf sqlRdr("AllowUpdate") = "0" Then
                    retVal = False
                End If
            Else
                retVal = False
            End If

            sqlConn.Close()
        End Using

        Return retVal
    End Function

    Public Function Auth_UserDelete(ByVal pUserID As String, ByVal pMenuID As String) As Boolean
        Dim retVal As Boolean = False

        Using sqlConn As New SqlConnection(ConnectionString)
            sqlConn.Open()

            Dim ls_SQL As String = "SELECT AllowDelete FROM dbo.SC_UserPrivilege WHERE AppID = 'P01' AND UserID = '" & Trim(pUserID) & "' AND MenuID = '" & Trim(pMenuID) & "' AND UserCls = '1'"
            Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
            Dim sqlRdr As SqlDataReader = sqlCmd.ExecuteReader()

            If sqlRdr.Read() Then
                If sqlRdr("AllowDelete") = "1" Then
                    retVal = True
                ElseIf sqlRdr("AllowDelete") = "0" Then
                    retVal = False
                End If
            Else
                retVal = False
            End If

            sqlConn.Close()
        End Using

        Return retVal
    End Function

    Public Function Auth_UserConfirm(ByVal pUserID As String, ByVal pMenuID As String) As Boolean
        Dim retVal As Boolean = False

        Using sqlConn As New SqlConnection(ConnectionString)
            sqlConn.Open()

            Dim ls_SQL As String = "SELECT AllowConfirm FROM dbo.SC_UserPrivilege WHERE AppID = 'P01' AND UserID = '" & Trim(pUserID) & "' AND MenuID = '" & Trim(pMenuID) & "' AND UserCls = '1'"
            Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
            Dim sqlRdr As SqlDataReader = sqlCmd.ExecuteReader()

            If sqlRdr.Read() Then
                If sqlRdr("AllowConfirm") = "1" Then
                    retVal = True
                ElseIf sqlRdr("AllowConfirm") = "0" Then
                    retVal = False
                End If
            Else
                retVal = False
            End If

            sqlConn.Close()
        End Using

        Return retVal
    End Function
#End Region

#Region "Procedure"
    Public Sub InitializeComponentAndDesign(ByVal pSetASPxTextBox As pSet, ByVal pSetASPxButton As pSet, ByVal pSetASPxGridView As pSet, ByRef pRefPage As Page)
        For Each ctl As Control In pRefPage.Controls
            If ctl.Controls.Count > 0 Then
                For Each ctlSet As Control In ctl.Controls
                    'ASPxTextBox
                    If pSetASPxTextBox = True Then
                        If TypeOf (ctlSet) Is ASPxEditors.ASPxTextBox Then
                            CType(ctlSet, ASPxEditors.ASPxTextBox).Font.Name = "Verdana"
                            CType(ctlSet, ASPxEditors.ASPxTextBox).Font.Size = "8"
                            CType(ctlSet, ASPxEditors.ASPxTextBox).ForeColor = Color.Black
                        End If
                    End If

                    'ASPxButton
                    If pSetASPxButton = True Then
                        If TypeOf (ctlSet) Is ASPxEditors.ASPxButton Then
                            CType(ctlSet, ASPxEditors.ASPxButton).Font.Name = "Verdana"
                            CType(ctlSet, ASPxEditors.ASPxButton).Font.Size = "8"
                            CType(ctlSet, ASPxEditors.ASPxButton).ForeColor = Color.Black
                            CType(ctlSet, ASPxEditors.ASPxButton).Width = 80
                        End If
                    End If

                    'ASPxGridView
                    If pSetASPxGridView = True Then
                        If TypeOf (ctlSet) Is ASPxGridView.ASPxGridView Then
                            CType(ctlSet, ASPxGridView.ASPxGridView).Font.Name = "Verdana"
                            CType(ctlSet, ASPxGridView.ASPxGridView).Font.Size = "8"
                            CType(ctlSet, ASPxGridView.ASPxGridView).Styles.Header.BackColor = Color.Orange
                        End If
                    End If
                Next ctlSet
            End If
        Next ctl
    End Sub

    Public Sub FillMonth(ByRef pASPxComboBox As DevExpress.Web.ASPxEditors.ASPxComboBox, ByVal pUseShortMonth As Boolean)
        With pASPxComboBox
            .Items.Clear()

            If pUseShortMonth = True Then
                .Items.Add("Jan", "01")
                .Items.Add("Feb", "02")
                .Items.Add("Mar", "03")
                .Items.Add("Apr", "04")
                .Items.Add("May", "05")
                .Items.Add("Jun", "06")
                .Items.Add("Jul", "07")
                .Items.Add("Aug", "08")
                .Items.Add("Sep", "09")
                .Items.Add("Oct", "10")
                .Items.Add("Nov", "11")
                .Items.Add("Dec", "12")
            Else
                .Items.Add("January", "01")
                .Items.Add("February", "02")
                .Items.Add("March", "03")
                .Items.Add("April", "04")
                .Items.Add("May", "05")
                .Items.Add("June", "06")
                .Items.Add("July", "07")
                .Items.Add("August", "08")
                .Items.Add("September", "09")
                .Items.Add("October", "10")
                .Items.Add("November", "11")
                .Items.Add("December", "12")
            End If
        End With
    End Sub
#End Region

#Region "Initialization"
    ''' <summary>
    ''' Open config file and store the value in local variables.
    ''' </summary>
    ''' <param name="pConfigFile"></param>
    ''' <remarks></remarks>
    Public Sub New(Optional ByVal pConfigFile As String = "config.ini")
        ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("KonString").ToString
    End Sub
#End Region

End Class
