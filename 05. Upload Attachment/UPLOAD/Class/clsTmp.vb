Public Class clsTmp
    Private m_InvoiceNo As String
    Private m_SuratJalanNo As String
    Private m_KanbanNo As String
    Private m_AffiliateID As String
    Private m_AffiliateName As String
    Private m_SupplierID As String
    Private m_KanbanCycle As String
    Private m_InvoiceDate As Date
    Private m_DueDate As Date
    Private m_PaymentItem As String
    Private m_KanbanDate As Date
    Private m_DeliveryDate As Date
    Private m_Period As Date
    Private m_KanbanTime As String
    Private m_KanbanStatus As String
    Private m_PIC As String
    Private m_PartNo As String
    Private m_PONo As String
    Private m_PORevNo As String
    Private m_POSeqNo As Integer
    Private m_POKanbanCls As String
    Private m_CommercialCls As String
    Private m_ShipCls As String
    Private m_CurrCls As String
    Private m_ReceiveCurrCls As String
    Private m_InvCurrCls As String
    Private m_UnitCls As String
    Private m_DifferentCls As String
    Private m_JenisArmada As String
    Private m_DriverName As String
    Private m_DriverCont As String
    Private m_NoPol As String
    Private m_TotalBox As Integer
    Private m_AffiliateApproveUser As String
    Private m_AffiliateApproveDate As Date
    Private m_SupplierApproveUser As String
    Private m_SupplierApproveDate As Date
    Private m_EntryDate As Date
    Private m_EntryUser As String
    Private m_UpdateDate As Date
    Private m_UpdateUser As String
    Private m_DeliveryLocation As String
    Private m_Remarks As String

    Private m_KanbanQty As Integer
    Private m_DOQty As Integer
    Private m_POQty As Integer
    Private m_POQtyOld As Integer
    Private m_ReceiveQty As Integer
    Private m_InvQty As Integer
    Private m_Price As Double
    Private m_ReceivePrice As Double
    Private m_InvPrice As Double
    Private m_Amount As Double
    Private m_ReceiveAmount As Double
    Private m_InvAmount As Double
    Private m_TotalAmount As Double

    Private m_EmailAddress As String
    Private m_UserName As String
    Private m_Password As String
    Private m_Port As String
    Private m_POP As String
    Private m_AttachmentFolder As String
    Private m_AttachmentBackupFolder As String
    Private m_Interval As Integer

    Public Property ForwarderID As String
    Public Property ETDVendor1 As Date
    Public Property Forecast1 As Integer
    Public Property Forecast2 As Integer
    Public Property Forecast3 As Integer
    Public Property QtyBox As String

    Public Property Vassel As String
    Public Property NamaKapal As String
    Public Property ContainerNo As String
    Public Property SealNo As String
    Public Property Tare As String
    Public Property Gross As String
    Public Property TotalCarton As String

    Public Property DONo As String
    Public Property SizeContainer As String
    Public Property ETDJakarta As String
    Public Property ShippingLine As String
    Public Property DestinationPort As String

    Public Property PalletNo As String
    Public Property BoxNo As String
    Public Property BoxNo2 As String
    Public Property TotBoxEx As Integer
    Public Property Length As String
    Public Property Width As String
    Public Property Height As String
    Public Property M3 As String
    Public Property WeightPallet As String
    Public Property StuffingDate As String

    Private m_OrderNo As String
    Private m_GoodRecQty As Integer
    Private m_DefectRecQty As Integer
    Private m_ForwarderID As String

    Public m_JmlBox As Integer
    Public m_BoxNo As String

    Public Property InvoiceNo As String
        Get
            Return m_InvoiceNo
        End Get
        Set(ByVal value As String)
            m_InvoiceNo = value
        End Set
    End Property

    Public Property SuratJalanNo As String
        Get
            Return m_SuratJalanNo
        End Get
        Set(ByVal value As String)
            m_SuratJalanNo = value
        End Set
    End Property

    Public Property KanbanNo As String
        Get
            Return m_KanbanNo
        End Get
        Set(ByVal value As String)
            m_KanbanNo = value
        End Set
    End Property

    Public Property AffiliateID As String
        Get
            Return m_AffiliateID
        End Get
        Set(ByVal value As String)
            m_AffiliateID = value
        End Set
    End Property

    Public Property ETDSplit As Date
    Public Property QtySplit As String

    Public Property ETDSplit1 As String
    Public Property QtySplit1 As String

    Public Property ETDSplit2 As String
    Public Property QtySplit2 As String

    Public Property ETDSplit3 As String
    Public Property QtySplit3 As String

    Public Property ETDSplit4 As String
    Public Property QtySplit4 As String

    Public Property ETDSplit5 As String
    Public Property QtySplit5 As String

    'Public Property OrderNO1 As String
    'Public Property OrderNO2 As String
    'Public Property OrderNO3 As String
    'Public Property OrderNO4 As String
    'Public Property OrderNO5 As String

    Public Property AffiliateName As String
        Get
            Return m_AffiliateName
        End Get
        Set(ByVal value As String)
            m_AffiliateName = value
        End Set
    End Property

    Public Property SupplierID As String
        Get
            Return m_SupplierID
        End Get
        Set(ByVal value As String)
            m_SupplierID = value
        End Set
    End Property

    Public Property KanbanCycle As String
        Get
            Return m_KanbanCycle
        End Get
        Set(ByVal value As String)
            m_KanbanCycle = value
        End Set
    End Property

    Public Property KanbanDate As Date
        Get
            Return m_KanbanDate
        End Get
        Set(ByVal value As Date)
            m_KanbanDate = value
        End Set
    End Property

    Public Property DeliveryDate As Date
        Get
            Return m_DeliveryDate
        End Get
        Set(ByVal value As Date)
            m_DeliveryDate = value
        End Set
    End Property

    Public Property Period As Date
        Get
            Return m_Period
        End Get
        Set(ByVal value As Date)
            m_Period = value
        End Set
    End Property

    Public Property InvoiceDate As Date
        Get
            Return m_InvoiceDate
        End Get
        Set(ByVal value As Date)
            m_InvoiceDate = value
        End Set
    End Property

    Public Property DueDate As Date
        Get
            Return m_DueDate
        End Get
        Set(ByVal value As Date)
            m_DueDate = value
        End Set
    End Property

    Public Property PaymentItem As String
        Get
            Return m_PaymentItem
        End Get
        Set(ByVal value As String)
            m_PaymentItem = value
        End Set
    End Property

    Public Property KanbanTime As String
        Get
            Return m_KanbanTime
        End Get
        Set(ByVal value As String)
            m_KanbanTime = value
        End Set
    End Property

    Public Property KanbanStatus As String
        Get
            Return m_KanbanStatus
        End Get
        Set(ByVal value As String)
            m_KanbanStatus = value
        End Set
    End Property

    Public Property PIC As String
        Get
            Return m_PIC
        End Get
        Set(ByVal value As String)
            m_PIC = value
        End Set
    End Property

    Public Property PartNo As String
        Get
            Return m_PartNo
        End Get
        Set(ByVal value As String)
            m_PartNo = value
        End Set
    End Property

    Public Property PONo As String
        Get
            Return m_PONo
        End Get
        Set(ByVal value As String)
            m_PONo = value
        End Set
    End Property

    Public Property PORevNo As String
        Get
            Return m_PORevNo
        End Get
        Set(ByVal value As String)
            m_PORevNo = value
        End Set
    End Property

    Public Property POSeqNo() As Integer
        Get
            Return m_POSeqNo
        End Get
        Set(ByVal value As Integer)
            m_POSeqNo = value
        End Set
    End Property

    Public Property POKanbanCls As String
        Get
            Return m_POKanbanCls
        End Get
        Set(ByVal value As String)
            m_POKanbanCls = value
        End Set
    End Property

    Public Property CommercialCls As String
        Get
            Return m_CommercialCls
        End Get
        Set(ByVal value As String)
            m_CommercialCls = value
        End Set
    End Property

    Public Property ShipCls As String
        Get
            Return m_ShipCls
        End Get
        Set(ByVal value As String)
            m_ShipCls = value
        End Set
    End Property

    Public Property CurrCls As String
        Get
            Return m_CurrCls
        End Get
        Set(ByVal value As String)
            m_CurrCls = value
        End Set
    End Property

    Public Property ReceiveCurrCls As String
        Get
            Return m_ReceiveCurrCls
        End Get
        Set(ByVal value As String)
            m_ReceiveCurrCls = value
        End Set
    End Property

    Public Property InvCurrCls As String
        Get
            Return m_InvCurrCls
        End Get
        Set(ByVal value As String)
            m_InvCurrCls = value
        End Set
    End Property

    Public Property UnitCls As String
        Get
            Return m_UnitCls
        End Get
        Set(ByVal value As String)
            m_UnitCls = value
        End Set
    End Property

    Public Property DifferentCls As String
        Get
            Return m_DifferentCls
        End Get
        Set(ByVal value As String)
            m_DifferentCls = value
        End Set
    End Property

    Public Property JenisArmada() As String
        Get
            Return m_JenisArmada
        End Get
        Set(ByVal value As String)
            m_JenisArmada = value
        End Set
    End Property
    Public Property DriverName() As String
        Get
            Return m_DriverName
        End Get
        Set(ByVal value As String)
            m_DriverName = value
        End Set
    End Property
    Public Property DriverCont() As String
        Get
            Return m_DriverCont
        End Get
        Set(ByVal value As String)
            m_DriverCont = value
        End Set
    End Property
    Public Property NoPol() As String
        Get
            Return m_NoPol
        End Get
        Set(ByVal value As String)
            m_NoPol = value
        End Set
    End Property
    Public Property TotalBox() As Integer
        Get
            Return m_TotalBox
        End Get
        Set(ByVal value As Integer)
            m_TotalBox = value
        End Set
    End Property

    Public Property AffiliateApproveUser As String
        Get
            Return m_AffiliateApproveUser
        End Get
        Set(ByVal value As String)
            m_AffiliateApproveUser = value
        End Set
    End Property

    Public Property AffiliateApproveDate As Date
        Get
            Return m_AffiliateApproveDate
        End Get
        Set(ByVal value As Date)
            m_AffiliateApproveDate = value
        End Set
    End Property

    Public Property SupplierApproveUser As String
        Get
            Return m_SupplierApproveUser
        End Get
        Set(ByVal value As String)
            m_SupplierApproveUser = value
        End Set
    End Property

    Public Property SupplierApproveDate As Date
        Get
            Return m_SupplierApproveDate
        End Get
        Set(ByVal value As Date)
            m_SupplierApproveDate = value
        End Set
    End Property

    Public Property EntryDate As Date
        Get
            Return m_EntryDate
        End Get
        Set(ByVal value As Date)
            m_EntryDate = value
        End Set
    End Property

    Public Property EntryUser As String
        Get
            Return m_EntryUser
        End Get
        Set(ByVal value As String)
            m_EntryUser = value
        End Set
    End Property

    Public Property UpdateDate As Date
        Get
            Return m_UpdateDate
        End Get
        Set(ByVal value As Date)
            m_UpdateDate = value
        End Set
    End Property

    Public Property UpdateUser As String
        Get
            Return m_UpdateUser
        End Get
        Set(ByVal value As String)
            m_UpdateUser = value
        End Set
    End Property

    Public Property DeliveryLocation As String
        Get
            Return m_DeliveryLocation
        End Get
        Set(ByVal value As String)
            m_DeliveryLocation = value
        End Set
    End Property

    Public Property Remarks As String
        Get
            Return m_Remarks
        End Get
        Set(ByVal value As String)
            m_Remarks = value
        End Set
    End Property

    Public Property InvoiceQty As Integer
        Get
            Return m_InvQty
        End Get
        Set(ByVal value As Integer)
            m_InvQty = value
        End Set
    End Property

    Public Property ReceiveQty As Integer
        Get
            Return m_ReceiveQty
        End Get
        Set(ByVal value As Integer)
            m_ReceiveQty = value
        End Set
    End Property

    Public Property KanbanQty As Integer
        Get
            Return m_KanbanQty
        End Get
        Set(ByVal value As Integer)
            m_KanbanQty = value
        End Set
    End Property

    Public Property DOQty As Integer
        Get
            Return m_DOQty
        End Get
        Set(ByVal value As Integer)
            m_DOQty = value
        End Set
    End Property

    Public Property POQty As Integer
        Get
            Return m_POQty
        End Get
        Set(ByVal value As Integer)
            m_POQty = value
        End Set
    End Property

    Public Property POQtyOld As Integer
        Get
            Return m_POQtyOld
        End Get
        Set(ByVal value As Integer)
            m_POQtyOld = value
        End Set
    End Property

    Public Property Price As Double
        Get
            Return m_Price
        End Get
        Set(ByVal value As Double)
            m_Price = value
        End Set
    End Property

    Public Property Amount As Double
        Get
            Return m_Amount
        End Get
        Set(ByVal value As Double)
            m_Amount = value
        End Set
    End Property

    Public Property ReceivePrice As Double
        Get
            Return m_ReceivePrice
        End Get
        Set(ByVal value As Double)
            m_ReceivePrice = value
        End Set
    End Property

    Public Property ReceiveAmount As Double
        Get
            Return m_ReceiveAmount
        End Get
        Set(ByVal value As Double)
            m_ReceiveAmount = value
        End Set
    End Property

    Public Property InvPrice As Double
        Get
            Return m_InvPrice
        End Get
        Set(ByVal value As Double)
            m_InvPrice = value
        End Set
    End Property

    Public Property InvAmount As Double
        Get
            Return m_InvAmount
        End Get
        Set(ByVal value As Double)
            m_InvAmount = value
        End Set
    End Property

    Public Property TotalAmount As Double
        Get
            Return m_TotalAmount
        End Get
        Set(ByVal value As Double)
            m_TotalAmount = value
        End Set
    End Property

    Public Property EmailAddress As String
        Get
            Return m_EmailAddress
        End Get
        Set(ByVal value As String)
            m_EmailAddress = value
        End Set
    End Property

    Public Property UserName As String
        Get
            Return m_UserName
        End Get
        Set(ByVal value As String)
            m_UserName = value
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

    Public Property Port As String
        Get
            Return m_Port
        End Get
        Set(ByVal value As String)
            m_Port = value
        End Set
    End Property

    Public Property POP As String
        Get
            Return m_POP
        End Get
        Set(ByVal value As String)
            m_POP = value
        End Set
    End Property

    Public Property Attachment As String
        Get
            Return m_AttachmentFolder
        End Get
        Set(ByVal value As String)
            m_AttachmentFolder = value
        End Set
    End Property

    Public Property AttachmentBackup As String
        Get
            Return m_AttachmentBackupFolder
        End Get
        Set(ByVal value As String)
            m_AttachmentBackupFolder = value
        End Set
    End Property

    Public Property Interval As Integer
        Get
            Return m_Interval
        End Get
        Set(ByVal value As Integer)
            m_Interval = value
        End Set
    End Property

    Public Property OrderNo As String
        Get
            Return m_OrderNo
        End Get
        Set(ByVal value As String)
            m_OrderNo = value
        End Set
    End Property

    Public Property GoodRecQty() As Integer
        Get
            Return m_GoodRecQty
        End Get
        Set(ByVal value As Integer)
            m_GoodRecQty = value
        End Set
    End Property

    Public Property DefectRecQty() As Integer
        Get
            Return m_DefectRecQty
        End Get
        Set(ByVal value As Integer)
            m_DefectRecQty = value
        End Set
    End Property

End Class
