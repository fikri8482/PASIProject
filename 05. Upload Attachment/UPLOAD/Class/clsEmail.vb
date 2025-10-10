Public Class clsEmail
    Private m_SMTPServer As String
    Private m_SMTPDesc As String
    Private m_Port As String
    Private m_SMTPTimeout As Integer
    Private m_EnableSSL As String
    Private m_DefaultCredentials As Boolean
    Private m_UserName As String
    Private m_Password As String
    Private m_ReceipientName As String
    Private m_EmailAffiliateTo As String
    Private m_EmailAffiliateCC As String
    Private m_EmailFWDTo As String
    Private m_EmailFWDCC As String
    Private m_EmailPASITo As String
    Private m_EmailPASICC As String
    Private m_EmailSupplierTo As String
    Private m_EmailSupplierCC As String
    Private m_Subject As String
    Private m_Body As String
    Private m_MessageHeader As String
    Private m_MessageFooter As String
    Private m_LastUpdate As Date
    Private m_UserUpdate As String

    Public Property SMTPServer As String
        Get
            Return m_SMTPServer
        End Get
        Set(ByVal value As String)
            m_SMTPServer = value
        End Set
    End Property

    Public Property SMTPDesc As String
        Get
            Return m_SMTPDesc
        End Get
        Set(ByVal value As String)
            m_SMTPDesc = value
        End Set
    End Property

    Public Property Port As String
        Get
            Return m_Port
        End Get
        Set(ByVal value As String)
            m_Port = value
        End Set
    End Property

    Public Property EnableSSL As String
        Get
            Return m_EnableSSL
        End Get
        Set(ByVal value As String)
            m_EnableSSL = value
        End Set
    End Property

    Public Property DefaultCredentials As Boolean
        Get
            Return m_DefaultCredentials
        End Get
        Set(ByVal value As Boolean)
            m_DefaultCredentials = value
        End Set
    End Property

    Public Property SMTPTimeout As Integer
        Get
            Return m_SMTPTimeout
        End Get
        Set(ByVal value As Integer)
            m_SMTPTimeout = value
        End Set
    End Property

    Public Property SenderName As String
        Get
            Return m_UserName
        End Get
        Set(ByVal value As String)
            m_UserName = value
        End Set
    End Property

    Public Property EmailPASITo As String
        Get
            Return m_EmailPASITo
        End Get
        Set(ByVal value As String)
            m_EmailPASITo = value
        End Set
    End Property

    Public Property EmailPASICC As String
        Get
            Return m_EmailPASICC
        End Get
        Set(ByVal value As String)
            m_EmailPASICC = value
        End Set
    End Property

    Public Property EmailSupplierTo As String
        Get
            Return m_EmailSupplierTo
        End Get
        Set(ByVal value As String)
            m_EmailSupplierTo = value
        End Set
    End Property

    Public Property EmailSupplierCC As String
        Get
            Return m_EmailSupplierCC
        End Get
        Set(ByVal value As String)
            m_EmailSupplierCC = value
        End Set
    End Property

    Public Property EmailAffiliateTo As String
        Get
            Return m_EmailAffiliateTo
        End Get
        Set(ByVal value As String)
            m_EmailAffiliateTo = value
        End Set
    End Property

    Public Property EmailAffiliateCC As String
        Get
            Return m_EmailAffiliateCC
        End Get
        Set(ByVal value As String)
            m_EmailAffiliateCC = value
        End Set
    End Property

    Public Property EmailFWDTo As String
        Get
            Return m_EmailFWDTo
        End Get
        Set(ByVal value As String)
            m_EmailFWDTo = value
        End Set
    End Property

    Public Property EmailFWDCC As String
        Get
            Return m_EmailFWDCC
        End Get
        Set(ByVal value As String)
            m_EmailFWDCC = value
        End Set
    End Property

    Public Property Subject As String
        Get
            Return m_Subject
        End Get
        Set(ByVal value As String)
            m_Subject = value
        End Set
    End Property

    Public Property Body As String
        Get
            Return m_Body
        End Get
        Set(ByVal value As String)
            m_Body = value
        End Set
    End Property

    Public Property MessageHeader As String
        Get
            Return m_MessageHeader
        End Get
        Set(ByVal value As String)
            m_MessageHeader = value
        End Set
    End Property

    Public Property MessageFooter As String
        Get
            Return m_MessageFooter
        End Get
        Set(ByVal value As String)
            m_MessageFooter = value
        End Set
    End Property

    Public Property Password As String
        Get
            Return m_Password
        End Get
        Set(ByVal value As String)
            m_Password = value
        End Set
    End Property

    Public Property LastUpdate As Date
        Get
            Return m_lastUpdate
        End Get
        Set(ByVal value As Date)
            m_lastUpdate = value
        End Set
    End Property

    Public Property UserUpdate As String
        Get
            Return m_UserUpdate
        End Get
        Set(ByVal value As String)
            m_UserUpdate = value
        End Set
    End Property
End Class
