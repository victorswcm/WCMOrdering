Public Class COrderDetail

    Public Property Parent As COrderHeader
    Public Property ID As Integer
    Public Property TS As Byte()

    Public Property ProductId As Integer
    Public Property ProductCode As String
    Public Property ProductCode_Supplier As String = String.Empty
    Public Property ProductCode_Branded As String = String.Empty
    Public Property Product_Supplier As String = String.Empty
    Public Property Product_Branded As String = String.Empty
    Public Property Product As String
    Public Property UnitPrice As Decimal = 0

    Public Property IsProduce As Boolean
    Public Property SupplyingDepotID As Long
    Public Property TypeID As Long
    ''--------------------------------------------
    Public Property DeliveryNoteNum As String
    Public Property Qty As Decimal
    Public Property SOQty As Decimal 'Quantity on standing order - used for the amendments
    Public Property QtyDelivered As Decimal?
    'used for standing order
    Public Property Sun As Decimal
    Public Property Mon As Decimal
    Public Property Tue As Decimal
    Public Property Wed As Decimal
    Public Property Thu As Decimal
    Public Property Fri As Decimal
    Public Property Sat As Decimal

    Public Sub New(ByVal Parent As COrderHeader, ByVal nId As Integer, ByVal bytTS As Byte())
        Parent = Parent
        ID = nId
        TS = bytTS
        ProductId = 0
        Qty = 0
        SOQty = 0
    End Sub

    Public Sub New(ByVal Parent As COrderHeader, ByVal nId As Integer, ByVal bytTS As Byte(), ByVal nProductId As Integer, ByVal dQty As Decimal, ByVal dSOQty As Decimal)
        Parent = Parent
        ID = nId
        TS = bytTS
        ProductId = nProductId
        Qty = dQty
        SOQty = dSOQty
    End Sub

    Public Sub New(ByVal Parent As COrderHeader, ByVal nId As Integer, ByVal bytTS As Byte(), ByVal nProductId As Integer, ByVal dQty As Decimal, ByVal dSOQty As Decimal,
                   ByVal strCode As String, ByVal strProduct As String)
        Parent = Parent
        ID = nId
        TS = bytTS
        ProductId = nProductId
        Qty = dQty
        SOQty = dSOQty
        ProductCode = strCode
        Product = strProduct
    End Sub

    Public Sub New(ByVal Parent As COrderHeader, ByVal nId As Integer, ByVal bytTS As Byte(), ByVal nProductId As Integer, ByVal strCode As String, ByVal strProduct As String,
                   ByVal dSun As Decimal, ByVal dMon As Decimal, ByVal dTue As Decimal, ByVal dWed As Decimal, ByVal dThu As Decimal, ByVal dFri As Decimal, ByVal dSat As Decimal)
        Parent = Parent
        ID = nId
        TS = bytTS
        ProductId = nProductId
        ProductCode = strCode
        Product = strProduct
        Sun = dSun
        Mon = dMon
        Tue = dTue
        Wed = dWed
        Thu = dThu
        Fri = dFri
        Sat = dSat
    End Sub
End Class