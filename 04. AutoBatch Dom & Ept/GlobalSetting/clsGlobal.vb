'******************************************************************************************************************
'System name    : ADM PIS Karawang
'Class          : clsGlobal
'Overview       : 
'   class containing global functions that used in the entire solution
'******************************************************************************************************************
'Created by     : Ari
'Create Date    : 20-Dec-2011
'Modify History :
'Remarks        : 
'******************************************************************************************************************

Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing

Public Class clsGlobal

#Region "Declaration"
    Dim nRetry As Integer
    Dim ConStr As String
    Dim UserID As String
    Const AppID As String = "P01"
    Dim dtMsg As DataTable
    Dim lg As clsLog
    Const ConnectionErrorMsg As String = "A network-related or instance-specific error occurred while establishing a connection to SQL Server"
    Const TransportErrorMsg As String = "A transport-level error has occurred"

    Public Enum MsgTypeEnum
        InformationMsg = 0
        ErrorMsg = 1
    End Enum

    Public Enum MsgButtonEnum
        OKOnly = 0
        OKCancel = 1
        YesNo = 3
        YesNoCancel = 4
    End Enum

    Public Enum MsgResultEnum
        Cancel = 0
        OK = 1
        No = 2
        Yes = 3
    End Enum

    Public Enum MsgIDEnum
        InvalidUserIDorPassword = 1
        PleaseInput_XX = 2
        PleaseInputValid_XX = 3
        VINisNotFoundInWOS = 4
        XX_IsNotFound = 5
        DataIsAlreadyUsedIn_XX = 6
        InvalidFileFormat = 7
        InvalidFile = 8
        NoDataToImport = 9
        InvalidDateFormat = 10
        VINDoesNotMatchWithChassisNo = 11
        VINAlreadySuspend = 12
        VINAlreadyCancelSuspend = 13
        VINAlreadyJigIn = 14
        DataAlreadyExists = 15
        YouDontHavePrivilegeToUpdate = 16
        YouDontHavePrivilegeToReprint = 17
        DataIsNotFound = 18
        PleaseSelect_XX = 19
        VINAlreadyScrap = 20
        NoDataToDelete = 21
        NoDataToSave = 22
        VINisAlreadyActiveForPlanJiginDate = 23
        NoDataToCopy = 24
        StartTimeMustbeMoreThanPreviousEndTime = 25
        EndTimeMustBeMoreThanStartTime = 26
        BreakTimeMustBeWithinStartAndEndShift = 27
        DayShiftMustNotOverlapWithNightShift = 29
        OvertimeMustNotOverlapWithStartAndEndOfShift = 30
        WorkingHourCodeHasNotBeenSet = 33
        ProductionDateHasAlreadyPassed = 34
        VINLengthMustBe17 = 35
        NoDataChanges = 37
        VINHasJustBeenScannedHere = 40

        InvalidTPRoute = 48
        XX_Failed = 98
        DatabaseConnectionFailed = 99

        LoadDataSuccessfull = 101
        InsertDataSuccesfull = 102
        UpdateDataSuccesfull = 103
        DeleteDataSuccesfull = 104
        PrintDataSuccessfull = 105
        ExportToExcelSuccessfull = 106
        VINSuspendSuccessfull = 107
        VINCancelSuspendSuccessfull = 108

        XX_Successful = 121

    End Enum
#End Region

    ''' <summary>
    ''' Get form backcolor
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Formbackcolor As Color
        Get
            Return Color.LightSteelBlue
        End Get
    End Property

    ''' <summary>
    ''' Get grid backcolor
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GridBackcolor As Color
        Get
            Return Color.LemonChiffon
        End Get
    End Property

    ''' <summary>
    ''' Message box prompt when user wants to delete data.
    ''' </summary>
    ''' <param name="pItemDeleted"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConfirmDelete(Optional ByVal pItemDeleted As String = "this data") As MsgBoxResult
        Return MsgBox("Are you sure you want to delete " & pItemDeleted & " ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + vbDefaultButton2, "Delete")
    End Function

    ''' <summary>
    ''' Message box prompt when user wants to cancel operation.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConfirmCancel() As MsgBoxResult
        Return MsgBox("Are you sure you want to cancel this data without any changes ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + vbDefaultButton2, "Cancel")
    End Function

    Public Function uf_GetDataSet(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataSet
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.SelectCommand.CommandTimeout = 300
            da.Fill(ds)
            da = Nothing
            Return ds
        Else
            Using Cn As New SqlConnection(ConStr)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                da.SelectCommand.CommandTimeout = 300
                da.Fill(ds)
                da = Nothing
                Return ds
            End Using
        End If
    End Function

    Public Function uf_GetDataTable(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataTable
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            Dim dt As New DataTable
            da.Fill(ds)
            Return ds.Tables(0)
        Else
            Using Cn As New SqlConnection(ConStr)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                Dim dt As New DataTable
                da.Fill(ds)
                Return ds.Tables(0)
            End Using
        End If
    End Function

    Public Function uf_ExecuteSql(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As Integer
        Dim cmd As New SqlCommand
        Dim RowAff As Long

        If pCon Is Nothing Then
            Dim con As New SqlConnection
            con.ConnectionString = ConStr
            con.Open()
            cmd = New SqlCommand(Query, con)
            RowAff = cmd.ExecuteNonQuery
            con.Close()
        Else
            cmd = New SqlCommand(Query, pCon)
            If pTrans IsNot Nothing Then
                cmd.Transaction = pTrans
            End If
            RowAff = cmd.ExecuteNonQuery
        End If

        Return RowAff
    End Function

    Public Function uf_ExecuteScalar(ByVal Query As String, ByVal pCon As SqlConnection, Optional ByVal pTrans As SqlTransaction = Nothing) As Object
        Dim cmd As New SqlCommand(Query, pCon)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If

        Dim Result As Object
        Result = cmd.ExecuteScalar
        Return Result
    End Function

    ''' <summary>
    ''' Gets message description from database based on message ID.
    ''' </summary>
    ''' <param name="pMsgID">
    ''' Message ID.
    ''' </param>
    ''' <returns>
    ''' Returns message description.
    ''' </returns>
    ''' <remarks></remarks>
    Public Function uf_ConvertMsg(ByVal pMsgID As String) As String
        Try
            If Not IsNumeric(pMsgID) Then
                Return pMsgID
            End If
            Dim Msg As String = ""
            Dim column(0) As DataColumn
            column(0) = dtMsg.Columns(0)
            dtMsg.PrimaryKey = column

            Dim row As DataRow = dtMsg.Rows.Find(pMsgID)
            If row IsNot Nothing AndAlso row(1) IsNot Nothing Then
                Msg = row(1).ToString.TrimEnd
            End If
            Return Msg
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Sub WriteLog(ByVal pLogID As String)
        Dim q As String = "Select "

    End Sub

    Public Sub up_ShowMsg(ByVal pMessage As String, ByVal ptxtMsg As Control, Optional ByVal pMsgType As MsgTypeEnum = MsgTypeEnum.InformationMsg)
        If pMsgType = MsgTypeEnum.InformationMsg Then
            ptxtMsg.ForeColor = Color.Blue
        Else
            ptxtMsg.ForeColor = Color.Red
        End If
        ptxtMsg.Text = uf_ConError(pMessage)
    End Sub

    Private Function uf_ConError(ByVal strMsg As String) As String
        If strMsg.Contains(ConnectionErrorMsg) Or strMsg.Contains(TransportErrorMsg) Then
            Return "Database connection error"
        Else
            Return strMsg
        End If
    End Function

    Public Sub New(ByVal pConStr As String, ByVal pUserID As String)
        ConStr = pConStr
        UserID = pUserID
        lg = New clsLog(ConStr, UserID)
    End Sub

End Class


