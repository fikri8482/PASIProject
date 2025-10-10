Public Class clsTallyProperty
    Inherits clsGeneralProperty

    'Master
    Public Property ShippingInstructionNo As String
    Public Property ContainerNo As String
    Public Property SealNo As String

    Public Property Tare As Double
    Public Property Gross As Double
    Public Property TotalCarton As Double

    Public Property Vessel As String
    Public Property SizeContainer As String
    Public Property ShippingLine As String
    Public Property DestinationPort As String
    Public Property VesselName As String
    Public Property StuffingDate As String

    'Detail
    Public Property PalletNo As String
    Public Property BoxNoFrom As String
    Public Property BoxNoTo As String
    Public Property Length As Double
    Public Property Width As Double
    Public Property Height As Double
    Public Property M3 As Double
    Public Property WeightPallet As Double
    Public Property TotalBox As Double
End Class
