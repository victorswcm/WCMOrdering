Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.IO

Public Class COrderHeader
    Inherits CollectionBase

    Public Shared BuyerSequence As Integer = 0

    Public Enum BuyerSeq
        EFoods_3 = 3    ' Not in use
        Mitie_4 = 4
        Compass_5 = 5  'Not in use
        CaffeNero_6 = 6 'CrunchTime!
        Elior_7 = 7 'Lost contract from Oct 2020
        FoodBuy_Online_8 = 8 ' Compass
        Cypad_9 = 9 ' Not in use
        Medina_10 = 10
        BourneLeisure_11 = 11
        Interserve_12 = 12
        Poundland_14 = 14
        Zupa_16 = 16
        Johal_17 = 17
        McColls_18 = 18
        Weezy_19 = 19
        DN_Grahams_98 = 98
        EmailOrders_99 = 99
        PushEmailWebApp_500 = 500
        PushEmailP2P_999 = 999
        Order_Upload_1000 = 1000 ' use to generate orders
    End Enum
    Private Const _MeName As String = "COrderHeader"

    Private mCn As SqlClient.SqlConnection

    Private _OrderID As Integer = 0
    Private mbytTS As Byte()
    Private _p2pOrderId As Integer = 0 'Read Only
    Private _OrderNum As String = Nothing
    Private _CustomerId As Long = 0
    Private _HOId As Long = 0
    Private _StoreId As String = Nothing ' customer store num ( acc num)
    Private _Customer As String = Nothing
    Private _CustAcc As String = Nothing
    Private _Address1 As String = Nothing
    Private _Postcode As String = Nothing
    Private _CustEmail As String = Nothing
    Private _CustAltEmail As String = Nothing
    Private _CustAltEmail2 As String = Nothing
    Private _POFOrmat As String = Nothing ' Read Only

    Private _CustomerLocationCode As String = Nothing ' Read Only
    Private _CustomerUnitCode As String = Nothing ' Read Only
    Private _CustomerASNCode As String = Nothing ' Read Only
    Private _CustomerAgreementCode As String = Nothing ' Read Only
    Private _ASNSent As Boolean = False
    Private _RestrictedSalesGroup As String = Nothing

    Private _DepotId As Integer = 0
    Private _Depot As String = Nothing
    
    Private _DepotEmail As String = Nothing
    Private _DepotAltEmail As String = Nothing
    Private _DepotAltEmail2 As String = Nothing
    Private _Supplier As String = Nothing
    Private _SupplierID As Long = 0
    Private _Supplier_Produce As String = Nothing
    Private _SupplierID_Produce As Long = 0
    Private _ServingAccNum As String = Nothing
    Private _ServingAccNum_Produce As String = Nothing
    Private _DeliveryDate As Date = Nothing
    Private _DateEffectiveSun As Date = Nothing ' SUSPENDED used for standing order (week_ending effective - 6 days)
    Private _DateEffective As Date = Nothing ' 
    Private _Notes As String = ""
    Private _UserId As Integer? = Nothing
    Private _UserName As String = Nothing
    Private _LastUpdated As Date
    Private _EmailedDateTime As Date? = Nothing

    Private _CutOffDayMilk As Integer = 0
    Private _CutOffDayGoods As Integer = 0
    Private _CutOffDayDairy As Integer = 0

    Private msCutOffTimeMilk As String = "12:00"
    Private msCutOffTimeGoods As String = "12:00"
    Private msCutOffTimeDairy As String = "12:00"

    Private mbEDIOrdersFlag As Boolean? = Nothing

    Public CutOffDateTimeMilk As String = String.Empty
    Public CutOffDateTimeGoods As String = String.Empty
    Public CutOffDateTimeDairyGoods As String = String.Empty

    Private _DT_CutOff As DataTable

    Public DepotCode As String
    Public DeliveryNoteID As Integer = 0

    Public IsSoAmendment As Boolean = False ' standing order amendment flag, true when created from standing order
    Public DeliveryDate_Old As Date = Nothing
    Public RestrictedSalesGroupID As Integer = 0
    Public ErrorMsg As String = String.Empty
    Public OrderLoadedFromFile As Boolean = False
    Public SOLastDeliveryDate As Date
    Public DT_SLA As DataTable
    Public SLAInfo As String = String.Empty
    Public Property IssueID As Integer

    Public Property MinValueDairyAndBread As Decimal = 0
    Public Property MinValueProduce As Decimal = 0

    Private _DepotId_Produce As Integer = 0
    Private _Depot_Produce As String = Nothing
    Private _DepotEmail_Produce As String = Nothing
    Private _DepotAltEmail_Produce As String = Nothing
    Private _DepotAltEmail2_Produce As String = Nothing
    Public SLAInfo_Produce As String = String.Empty
    Private _EmailedDateTime_Produce As Date? = Nothing
    Public DepotCode_Produce As String

    Private mintBuyingGroupID As Integer? = Nothing

    Public Event Report_Error(pErrMsg As String)
    Public Sub New()

    End Sub

    Public Sub New(ByVal intOrderId As Integer)
        GetOrder(intOrderId)
    End Sub

    Public Sub New(ByVal intOrderId As Integer, pCheckIfAmendment As Boolean)
        _CheckIfAmendment = pCheckIfAmendment
        GetOrder(intOrderId)
    End Sub

    Public Sub Add(ByVal intId As Integer, ByVal bytTS As Byte(), ByVal intProductId As Integer, ByVal strCode As String, ByVal strProduct As String, ByVal decQty As Decimal, ByVal decSOQty As Decimal)
        Dim oOrderLine As New COrderDetail(Me, intId, bytTS, intProductId, decQty, decSOQty, strCode, strProduct)
        List.Add(oOrderLine)
    End Sub

    'items to be deleted
    Public Sub Add(ByVal intId As Integer, ByVal bytTS As Byte())
        Dim oOrderLine As New COrderDetail(Me, intId, bytTS)
        List.Add(oOrderLine)
    End Sub

    Public Sub Add(ByVal value As COrderDetail)
        List.Add(value)
    End Sub

    'used for standing orders:
    Public Sub Add(ByVal intId As Integer, ByVal bytTS As Byte(), ByVal intProductId As Integer, ByVal strCode As String, ByVal strProduct As String, _
                   ByVal decSun As Decimal, ByVal decMon As Decimal, ByVal decTue As Decimal, ByVal decWed As Decimal, ByVal decThu As Decimal, ByVal decFri As Decimal, ByVal decSat As Decimal)

        Dim oOrderLine As New COrderDetail(Me, intId, bytTS, intProductId, strCode, strProduct, decSun, decMon, decTue, decWed, decThu, decFri, decSat)
        List.Add(oOrderLine)
    End Sub

    'remove items with id = 0 (those not in the database)
    Public Sub Clean()
        Dim idx As Integer

        For idx = Me.Count - 1 To 0 Step -1
            If DirectCast(List.Item(idx), COrderDetail).ID = 0 Then
                List.RemoveAt(idx)
            End If
        Next
    End Sub

    Public Function Contains(ByVal value As COrderDetail) As Boolean
        Return List.Contains(value)
    End Function

    Public Function Contains(ByVal sProductCode As String) As Boolean
        Dim idx As Integer

        For idx = 0 To List.Count - 1
            If StrComp(DirectCast(List.Item(idx), COrderDetail).ProductCode, sProductCode, vbTextCompare) = 0 Then
                Return True
            End If
        Next idx
        Return False
    End Function

    Public Function Exists(ByVal intProductId As Integer) As Boolean
        Dim idx As Integer

        For idx = 0 To List.Count - 1
            If DirectCast(List.Item(idx), COrderDetail).ProductId = intProductId Then
                Return True
            End If
        Next idx
        Return False
    End Function

    Public Function IndexOf(ByVal value As COrderDetail) As Integer
        Return List.IndexOf(value)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal value As COrderDetail)
        List.Insert(index, value)
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As COrderDetail
        Get
            Return DirectCast(List.Item(index), COrderDetail)
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal sProductCode As String) As COrderDetail
        Get
            Dim oItem As COrderDetail
            Dim idx As Integer

            For idx = 0 To List.Count - 1
                oItem = DirectCast(List.Item(idx), COrderDetail)
                If StrComp(oItem.ProductCode, sProductCode, vbTextCompare) = 0 Then
                    Return oItem
                End If
            Next idx
            Return Nothing
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal pID As Integer, ByVal SearchBy As String) As COrderDetail
        Get
            Dim oItem As COrderDetail
            Dim idx As Integer

            If SearchBy = "ProductId" Then
                For idx = 0 To List.Count - 1
                    oItem = DirectCast(List.Item(idx), COrderDetail)
                    If oItem.ProductId = pID Then
                        Return oItem
                    End If
                Next idx
            ElseIf SearchBy = "RecordId" Then
                For idx = 0 To List.Count - 1
                    oItem = DirectCast(List.Item(idx), COrderDetail)
                    If oItem.ID = pID Then
                        Return oItem
                    End If
                Next idx
            End If
            Return Nothing
        End Get
    End Property

    ''' <summary>
    ''' if any of the products BD1000 BD1001 BD1002 BD1003 have been ordered need to create separate email for specific supplier (TBA)
    ''' </summary>
    ''' <value></value>
    ''' <returns>true or false</returns>
    ''' <remarks></remarks>

    Public ReadOnly Property Is_BDProduct_Ordered As Boolean
        Get
            Dim oItem As COrderDetail
            Dim idx As Integer

            For idx = 0 To List.Count - 1
                oItem = DirectCast(List.Item(idx), COrderDetail)

                If oItem.ProductId >= 5385 AndAlso oItem.ProductId <= 5388 Then
                    Return True
                End If
            Next idx
            Return False
        End Get
    End Property
    ''' <summary>
    ''' if any of the products besides BD1000 BD1001 BD1002 BD1003 have been ordered return true
    ''' </summary>
    ''' <value></value>
    ''' <returns>true or false</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Is_Non_BDProduct_Ordered As Boolean
        Get
            Dim oItem As COrderDetail
            Dim idx As Integer

            For idx = 0 To List.Count - 1
                oItem = DirectCast(List.Item(idx), COrderDetail)

                If oItem.ProductId < 5385 OrElse oItem.ProductId > 5388 Then
                    Return True
                End If
            Next idx
            Return False
        End Get
    End Property
    Public ReadOnly Property ShowProductCodes_Supplier As Boolean
        Get
            'TODO - hard coded for HOV1 - Hovis Limited; need to create flag on supplier table 
            Return SupplierID = 117621290
        End Get
    End Property
    Public ReadOnly Property ShowProductCodes_Supplier_Produce As Boolean
        Get
            'TODO - hard coded for HOV1 - Hovis Limited; need to create flag on supplier table 
            Return SupplierID_Produce = 117621290
        End Get
    End Property
    Public Sub Remove(ByVal Value As COrderDetail)
        If (List.Contains(Value)) Then
            List.Remove(Value)
        End If
    End Sub

    Private Sub SetDateEffectiveSunday(ByVal dtDate As Date)
        'Get Sunday (day 1) of this week
        While dtDate.DayOfWeek <> DayOfWeek.Sunday
            dtDate = dtDate.AddDays(-1)
        End While

        _DateEffectiveSun = dtDate
    End Sub

    Public ReadOnly Property TotalQty As Integer
        Get
            Dim clsLine As COrderDetail
            Dim lTotal As Integer = 0
            For Each clsLine In Me.List
                lTotal += clsLine.Qty
            Next
            Return lTotal
        End Get
    End Property

    ' not nesesseraly milk, depends on what is set up in  depot_product lookup table for current customer
    Public ReadOnly Property MilkItems As List(Of COrderDetail)
        Get
            Return GetItems(DepotId)
        End Get
    End Property
    ' not nesesseraly milk, depends on what is set up in  depot_product lookup table for current customer
    Public ReadOnly Property ProduceItems As List(Of COrderDetail)
        Get
            Return GetItems(DepotId_Produce)
        End Get
    End Property

    Private Function GetItems(pDepotId As Long) As List(Of COrderDetail)
        Dim lList As New List(Of COrderDetail)

        For Each lItem As COrderDetail In Me.List
            If lItem.SupplyingDepotID = pDepotId Then
                lList.Add(lItem)
            End If
        Next
        Return lList
    End Function


    Public Property OrderId() As Integer
        Get
            Return _OrderID
        End Get
        Set(ByVal value As Integer)
            _OrderID = value
        End Set
    End Property

    Public Property CustomerId() As Long
        Get
            Return _CustomerId
        End Get
        Set(ByVal value As Long)
            _CustomerId = value
        End Set
    End Property

    Public Property HOId() As Long
        Get
            Return _HOId
        End Get
        Set(ByVal value As Long)
            _HOId = value
        End Set
    End Property

    Public Property Customer() As String
        Get
            If _Customer Is Nothing Then
                GetCustomerInfo()
            End If
            Return _Customer
        End Get
        Set(ByVal value As String)
            _Customer = value
        End Set
    End Property

    Public Property CustAcc() As String
        Get
            If _CustAcc Is Nothing Then
                GetCustomerInfo()
            End If
            Return _CustAcc.Trim
        End Get
        Set(ByVal value As String)
            _CustAcc = value
        End Set
    End Property
    
    Public Property DepotId_Produce() As Integer
        Get
            Return _DepotId_Produce
        End Get
        Set(ByVal value As Integer)
            _DepotId_Produce = value
        End Set
    End Property
    Public Property Depot_Produce() As String
        Get
            If _Depot_Produce Is Nothing Then
                GetDepotInfo(True)
            End If
            Return _Depot_Produce
        End Get
        Set(ByVal value As String)
            _Depot_Produce = value
        End Set
    End Property
    
    Public Property DepotId() As Integer
        Get
            Return _DepotId
        End Get
        Set(ByVal value As Integer)
            _DepotId = value
        End Set
    End Property

    Public Property Depot() As String
        Get
            If _Depot Is Nothing Then
                GetDepotInfo()
            End If
            Return _Depot
        End Get
        Set(ByVal value As String)
            _Depot = value
        End Set
    End Property
   
  Public Property DepotEmail() As String
        Get
            If _DepotEmail Is Nothing Then
                GetDepotInfo()
            End If
            Return _DepotEmail
        End Get
        Set(ByVal value As String)
            _DepotEmail = value
        End Set
    End Property

    Public Property DepotAltEmail() As String
        Get
            If _DepotAltEmail Is Nothing Then
                GetDepotInfo()
            End If
            Return _DepotAltEmail
        End Get
        Set(ByVal value As String)
            _DepotAltEmail = value
        End Set
    End Property

    Public Property DepotAltEmail2() As String
        Get
            If _DepotAltEmail2 Is Nothing Then
                GetDepotInfo()
            End If
            Return _DepotAltEmail2
        End Get
        Set(ByVal value As String)
            _DepotAltEmail2 = value
        End Set
    End Property

    Public Property DepotEmail_Produce() As String
        Get
            If _DepotEmail_Produce Is Nothing Then
                GetDepotInfo(True)
            End If
            Return _DepotEmail_Produce
        End Get
        Set(ByVal value As String)
            _DepotEmail_Produce = value
        End Set
    End Property

    Public Property DepotAltEmail_Produce() As String
        Get
            If _DepotAltEmail_Produce Is Nothing Then
                GetDepotInfo(True)
            End If
            Return _DepotAltEmail_Produce
        End Get
        Set(ByVal value As String)
            _DepotAltEmail_Produce = value
        End Set
    End Property

    Public Property DepotAltEmail2_Produce() As String
        Get
            If _DepotAltEmail2_Produce Is Nothing Then
                GetDepotInfo(True)
            End If
            Return _DepotAltEmail2_Produce
        End Get
        Set(ByVal value As String)
            _DepotAltEmail2_Produce = value
        End Set
    End Property
    Public Property EDIOrdersFlag() As Boolean
        Get
            If mbEDIOrdersFlag Is Nothing Then
                GetDepotInfo()
            End If
            Return Nz(Of Boolean)(mbEDIOrdersFlag, 0)
        End Get
        Set(ByVal value As Boolean)
            mbEDIOrdersFlag = value
        End Set
    End Property

    Public Property Supplier() As String
        Get
            If _Supplier Is Nothing Then
                GetCurrentSupplier()
            End If
            Return _Supplier
        End Get
        Set(ByVal value As String)
            _Supplier = value
        End Set
    End Property

    Public Property SupplierID() As Integer
        Get
            If _SupplierID = 0 Then
                GetCurrentSupplier()
            End If
            Return _SupplierID
        End Get
        Set(ByVal value As Integer)
            _SupplierID = value
        End Set
    End Property
    Public Property Supplier_Produce() As String
        Get
            If _Supplier_Produce Is Nothing Then
                GetCurrentSupplier(True)
            End If
            Return _Supplier_Produce
        End Get
        Set(ByVal value As String)
            _Supplier_Produce = value
        End Set
    End Property

    Public Property SupplierID_Produce() As Integer
        Get
            If _SupplierID_Produce = 0 Then
                GetCurrentSupplier(True)
            End If
            Return _SupplierID_Produce
        End Get
        Set(ByVal value As Integer)
            _SupplierID_Produce = value
        End Set
    End Property
    Public Property ServingAccNum() As String
        Get
            If _ServingAccNum Is Nothing Then
                GetServingAccNum_and_SLA()
            End If
            Return _ServingAccNum
        End Get
        Set(ByVal value As String)
            _ServingAccNum = value
        End Set
    End Property

    Public Property ServingAccNum_Produce() As String
        Get
            If _ServingAccNum_Produce Is Nothing Then
                GetServingAccNum_and_SLA(True)
            End If
            Return _ServingAccNum_Produce
        End Get
        Set(ByVal value As String)
            _ServingAccNum_Produce = value
        End Set
    End Property
    Public Property DateEffective() As Date
        Get
            Return _DateEffective
        End Get
        Set(ByVal value As Date)
            _DateEffective = value
        End Set
    End Property

    Public Property DateEffectiveSun() As Date
        Get
            Return _DateEffectiveSun
        End Get
        Set(ByVal value As Date)
            _DateEffectiveSun = value
        End Set
    End Property

    Public Property DeliveryDate() As Date
        Get
            Return _DeliveryDate
        End Get
        Set(ByVal value As Date)
            _DeliveryDate = value
        End Set
    End Property

    Public Property Notes() As String
        Get
            Return _Notes
        End Get
        Set(ByVal value As String)
            _Notes = value
        End Set
    End Property

    Public Property OrderNum() As String
        Get
            Dim lNum As Integer
            If String.IsNullOrEmpty(_OrderNum) Then
                If HOId = 160973922 OrElse HOId = 160974269 OrElse HOId = 160974511 OrElse HOId = 245739642 OrElse HOId = 245740000 Then
                    'do nothing
                Else
                    'If POFOrmat = "" Then
                    If Not String.IsNullOrEmpty(DepotCode) Then
                        lNum = DepotCode.Length
                        If lNum > 3 Then lNum = 3
                        If DepotCode.Length > lNum Then
                            _OrderNum = CustAcc & DepotCode.Substring(0, lNum) & Format(_DeliveryDate, "yyMMdd")
                        Else
                            _OrderNum = CustAcc & Depot.Substring(0, 1) & Format(_DeliveryDate, "yyMMdd")
                        End If
                    Else
                        _OrderNum = CustAcc & Format(_DeliveryDate, "yyMMdd")
                    End If
                    'Else
                    '_OrderNum = POFOrmat
                    'End If
                End If
            End If
            Return _OrderNum

        End Get
        Set(ByVal value As String)
            _OrderNum = value
        End Set
    End Property
    Public Property LastUpdated() As Date
        Get
            Return _LastUpdated
        End Get
        Set(ByVal value As Date)
            _LastUpdated = value
        End Set
    End Property

    Public Property UserId() As Integer
        Get
            If _UserId Is Nothing Then
                _UserId = 1
            End If
            Return _UserId
        End Get
        Set(ByVal value As Integer)
            _UserId = value
        End Set
    End Property

    Public Property EmailedDateTime() As Date?
        Get
            Return _EmailedDateTime
        End Get
        Set(ByVal value As Date?)
            _EmailedDateTime = value
        End Set
    End Property
    
    Public Property EmailedDateTime_Produce() As Date?
        Get
            Return _EmailedDateTime_Produce
        End Get
        Set(ByVal value As Date?)
            _EmailedDateTime_Produce = value
        End Set
    End Property
    
    Public Property BuyingGroupID() As Integer
        Get
            If _BuyingGroupID Is Nothing Then
                _BuyingGroupID = 0
            End If
            Return _BuyingGroupID
        End Get
        Set(ByVal value As Integer)
            _BuyingGroupID = value
        End Set
    End Property
    Public Property StoreId() As String
        Get
            If _StoreId Is Nothing Then
                GetCustomerInfo()
            End If
            Return _StoreId
        End Get
        Set(ByVal value As String)
            _StoreId = value.Trim()
        End Set
    End Property

    Public Property Address1() As String
        Get
            If _Address1 Is Nothing Then
                GetCustomerInfo()
            End If
            Return _Address1
        End Get
        Set(ByVal value As String)
            _Address1 = value.Trim()
        End Set
    End Property

    Public Property Postcode() As String
        Get
            If _Postcode Is Nothing Then
                GetCustomerInfo()
            End If
            Return _Postcode
        End Get
        Set(ByVal value As String)
            _Postcode = value.Trim
        End Set
    End Property

    Public Property CustEmail() As String
        Get
            If _CustEmail Is Nothing Then
                GetCustomerInfo()
            End If
            Return _CustEmail
        End Get
        Set(ByVal value As String)
            _CustEmail = value.Trim()
        End Set
    End Property

    Public Property CustAltEmail() As String
        Get
            If _CustAltEmail Is Nothing Then
                GetCustomerInfo()
            End If
            Return _CustAltEmail
        End Get
        Set(ByVal value As String)
            _CustAltEmail = value.Trim()
        End Set
    End Property

    Public Property CustAltEmail2() As String
        Get
            If _CustAltEmail2 Is Nothing Then
                GetCustomerInfo()
            End If
            Return _CustAltEmail2
        End Get
        Set(ByVal value As String)
            _CustAltEmail2 = value.Trim()
        End Set
    End Property

    Public ReadOnly Property OrderMinValue() As Decimal
        Get
            Return MinValueDairyAndBread + MinValueProduce
        End Get
    End Property

    Public ReadOnly Property POFOrmat() As String
        Get
            If _POFOrmat Is Nothing Then
                GetCustomerInfo()
            End If
            Return _POFOrmat
        End Get
    End Property

    Public ReadOnly Property CustomerLocationCode() As String
        Get
            If _CustomerLocationCode Is Nothing Then
                GetCustomerInfo()
            End If
            Return _CustomerLocationCode
        End Get
    End Property

    Public ReadOnly Property CustomerAgreementCode() As String
        Get
            If _CustomerAgreementCode Is Nothing Then
                GetCustomerInfo()
            End If
            Return _CustomerAgreementCode
        End Get
    End Property

    Public ReadOnly Property ASNSent() As Boolean
        Get

            Return _ASNSent
        End Get
    End Property

    Public ReadOnly Property CustomerASNCode() As String
        Get
            If _CustomerASNCode Is Nothing Then
                GetCustomerInfo()
            End If
            Return _CustomerASNCode
        End Get
    End Property

    Public ReadOnly Property CustomerUnitCode() As String
        Get
            If _CustomerUnitCode Is Nothing Then
                GetCustomerInfo()
            End If
            Return _CustomerUnitCode
        End Get
    End Property

    Public ReadOnly Property p2pOrderId() As Integer
        Get
            Return _p2pOrderId
        End Get
    End Property

    Public Function IsValidPOFormat(pOrderNum As String) As Boolean

        Select Case POFOrmat.ToLower
            Case "", "generic", "na", "n/a"
                Return True
            Case "por", "por-xxxxx"
                If (pOrderNum Like "POR-PHONE") OrElse (pOrderNum Like "POR-BC") OrElse (pOrderNum Like "POR????????") OrElse (pOrderNum Like "???-???????") Then
                    Return True
                End If
        End Select

        'Validate "Regex" like DHOxxxxxxx, 801953xxxxxxx, xxxxxxMCIL,, etc.

        If pOrderNum Like POFOrmat.Replace("x", "?") Then
            Return True
        End If

        Return False
    End Function
    Public Property RestrictedSalesGroup() As String
        Get
            If _RestrictedSalesGroup Is Nothing Then
                GetCustomerInfo()
            End If
            Return _RestrictedSalesGroup
        End Get
        Set(ByVal value As String)
            _RestrictedSalesGroup = value.Trim
        End Set
    End Property

    Public Property TS() As Byte()
        Get
            Return mbytTS
        End Get
        Set(ByVal value As Byte())
            mbytTS = value
        End Set
    End Property

    Public Property UserName() As String
        Get
            If _UserName Is Nothing Then
                _UserName = "Admin"
            End If
            Return _UserName
        End Get
        Set(ByVal value As String)
            _UserName = value
        End Set
    End Property

    'Get Order by Id
    Public Function GetOrder(ByVal nOrderId As Integer) As Boolean
        If RetrieveOrderHeader(nOrderId, Nothing, Nothing, Nothing) Then
            If RetrieveOrderLines(nOrderId) = False Then
                Return False
            End If
        End If
        Return True
    End Function

    'Get Order by Order Number
    Public Function GetOrder(ByVal strOrderNum As String) As Boolean
        If RetrieveOrderHeader(Nothing, strOrderNum, Nothing, Nothing) Then
            If RetrieveOrderLines(OrderId) = False Then
                Return False
            End If
        End If
        Return True
    End Function

    'Get Order by Order Number and delivery date 
    Public Function GetOrder(ByVal strOrderNum As String, ByVal dtDeliveryDate As Date) As Boolean
        If RetrieveOrderHeader(Nothing, strOrderNum, Nothing, dtDeliveryDate) Then
            If RetrieveOrderLines(OrderId) = False Then
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub GetCustomerInfo()
        Dim cmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim l_DB As DB
        Dim sbsql As New System.Text.StringBuilder
        Try
            If _CustomerId <> 0 Then
                l_DB = New DB
                l_DB.Open()

                With sbsql
                    .AppendLine("SELECT customer_account_number, isnull(SiteName,customer_account_name) customer, cu_store_number, address_line1, postcode, isnull(rsg_restricted_sales_group,'') restricted_sales_group")
                    .AppendLine(" FROM v_customer_links  ")
                    .AppendLine(" LEFT JOIN restricted_sales_groups ON rsg_restricted_sales_group_id = cu_restricted_sales_group_id ")
                    .Append(" WHERE cu_sage_customer_id =  ") : .AppendLine(_CustomerId.ToString)

                    cmd = New SqlClient.SqlCommand(.ToString, l_DB.Connection)
                End With
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    If dr.Read Then
                        Me.CustAcc = dr("customer_account_number").ToString
                        Me.Customer = dr("customer").ToString
                        Me.StoreId = dr("cu_store_number").ToString
                        Me.Address1 = dr("address_line1").ToString
                        Me.Postcode = dr("postcode").ToString
                        Me.RestrictedSalesGroup = dr("restricted_sales_group").ToString
                    End If
                End If
                dr.Close()
                l_DB.Close()
            Else
                mstrCustomer = ""
                mstrStoreId = ""
            End If

        Catch ex As Exception
            RaiseEvent Report_Error("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)

        Finally

        End Try
    End Sub

    Public Sub GetDepotInfo()
        GetDepotInfo(0)
    End Sub

    Public Sub GetDepotInfo(pIsProduce As Boolean)
        Dim cmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim l_DB As DB
        Dim sbsql As New System.Text.StringBuilder
        Dim lDepotID As Integer = _DepotId

        Try

            If pIsProduce Then lDepotID = _DepotId_Produce

            If _DepotId <> 0 Then
                l_DB = New DB
                l_DB.Open()

                With sbsql
                    .AppendLine("SELECT de_depot, de_code, de_email, de_alt_email, de_alt_email2, de_edi_orders")
                    .AppendLine("FROM depots ")
                    .Append("WHERE de_depot_id = ") : .Append(_DepotId.ToString)

                    cmd = New SqlClient.SqlCommand(.ToString, l_DB.Connection)
                End With
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    If dr.Read Then
                        If pIsProduce Then
                            Me.Depot_Produce = dr("de_depot").ToString
                            Me.DepotCode_Produce = Nz(Of String)(dr("de_code"), "")
                            Me.DepotEmail_Produce = Nz(Of String)(dr("de_email"), "")
                            Me.DepotAltEmail_Produce = Nz(Of String)(dr("de_alt_email"), "")
                            Me.DepotAltEmail2_Produce = Nz(Of String)(dr("de_alt_email2"), "")
                        Else
                            _Depot = dr("de_depot").ToString
                            DepotCode = Nz(Of String)(dr("de_code"), "")
                            _DepotEmail = If(IsDBNull(dr("de_email")), "", dr("de_email"))
                            _DepotAltEmail = If(IsDBNull(dr("de_alt_email")), "", dr("de_alt_email"))
                            _DepotAltEmail2 = If(IsDBNull(dr("de_alt_email2")), "", dr("de_alt_email2"))
                            mbEDIOrdersFlag = Nz(Of Boolean)(dr("de_edi_orders"), 0)
                        End If

                    End If
                End If
                dr.Close()
                l_DB.Close()
            Else
                _Depot = ""
                _DepotEmail = Nothing
                _DepotAltEmail = Nothing
                _DepotAltEmail2 = Nothing
                mbEDIOrdersFlag = Nothing
            End If
        Catch ex As Exception
            RaiseEvent Report_Error("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)

        Finally

        End Try
    End Sub

    Public Sub GetServingAccNum_and_SLA()
        GetServingAccNum_and_SLA(0)
    End Sub

    Public Sub GetServingAccNum_and_SLA(pIsProduce As Boolean)
        Dim dr As SqlClient.SqlDataReader
        Dim lDB As DB
        Dim lCmd As SqlClient.SqlCommand
        Dim lParam As SqlClient.SqlParameter
        Dim sbsql As New System.Text.StringBuilder

        Try
            If pIsProduce AndAlso _DepotId_Produce = 0 Then
                _ServingAccNum_Produce = ""
                Return
            End If

            ' If _CustomerId <> 0 Then
            lDB = New DB
            lDB.Open()
            lCmd = New SqlClient.SqlCommand("p_serving_acc_and_sla_get", lDB.Connection)
            With lCmd
                .CommandType = CommandType.StoredProcedure

                'Create Parameters
                lParam = .Parameters.Add("@cust_id", SqlDbType.BigInt)
                lParam.Value = _CustomerId

                lParam = .Parameters.Add("@acc_num", SqlDbType.VarChar, 60)
                lParam.Value = _CustAcc

                lParam = .Parameters.Add("@depot_id", SqlDbType.Int)
                'add to WCMOrdering
                If pIsProduce Then
                    lParam.Value = _DepotId_Produce
                Else
                    lParam.Value = _DepotId
                End If

                lParam = .Parameters.Add("@delivery_date", SqlDbType.Date)
                If IsNothing(_DeliveryDate) Then
                    lParam.Value = DBNull.Value
                Else
                    lParam.Value = _DeliveryDate
                End If

                dr = .ExecuteReader()

                If dr.HasRows Then
                    dr.Read()
                    If pIsProduce Then
                        _ServingAccNum_Produce = Nz(Of String)(dr("serving_code"), "")
                    Else
                        _ServingAccNum = Nz(Of String)(dr("serving_code"), "")

                    End If

                Else
                    If pIsProduce Then
                        _ServingAccNum_Produce = ""
                    Else
                        _ServingAccNum = ""
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseEvent Report_Error("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Finally

        End Try
    End Sub

    Public Sub GetCurrentSupplier()
        GetCurrentSupplier(0)
    End Sub

    Public Sub GetCurrentSupplier(pIsProduce As Boolean)
        Dim lCmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim lParam As SqlClient.SqlParameter
        Dim lResult As Object = Nothing
        Dim lDB As DB
        Dim sbsql As New System.Text.StringBuilder
        Dim lDepotID As Integer = _DepotId

        Try
            If pIsProduce Then
                lDepotID = _DepotId_Produce
            End If

            lDB = New DB
            lDB.Open()
            lCmd = New SqlClient.SqlCommand("p_customer_ordering_info_get", lDB.Connection)
            With lCmd
                .CommandType = CommandType.StoredProcedure

                'Create Parameters
                lParam = .Parameters.Add("@return", SqlDbType.Int)
                lParam.Direction = ParameterDirection.ReturnValue

                lParam = .Parameters.Add("@cust_id", SqlDbType.BigInt)
                lParam.Value = _CustomerId

                lParam = .Parameters.Add("@delivery_date", SqlDbType.Date)
                If IsNothing(_DeliveryDate) Then
                    lParam.Value = DBNull.Value
                Else
                    lParam.Value = _DeliveryDate
                End If

                lParam = .Parameters.Add("@depot_id", SqlDbType.Int)
                lParam.Direction = ParameterDirection.InputOutput
                lParam.Value = lDepotID

                lParam = .Parameters.Add("@depot", SqlDbType.VarChar, 50)
                lParam.Direction = ParameterDirection.InputOutput

                lParam = .Parameters.Add("@serv_code", SqlDbType.VarChar, 16)
                lParam.Direction = ParameterDirection.InputOutput

                lParam = .Parameters.Add("@supplier_id", SqlDbType.BigInt)
                lParam.Direction = ParameterDirection.InputOutput

                lParam = .Parameters.Add("@supplier_acc", SqlDbType.VarChar, 16)
                lParam.Direction = ParameterDirection.InputOutput

                lParam = .Parameters.Add("@supplier", SqlDbType.VarChar, 60)
                lParam.Direction = ParameterDirection.InputOutput

                If pIsProduce Then
                    lParam = .Parameters.Add("@supply_type", SqlDbType.Int)
                    lParam.Value = 2
                Else
                    'defaults to 1
                End If
                lParam = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                lParam.Direction = ParameterDirection.Output

                .ExecuteNonQuery()

                'note it returns customer_price_info_id
                Select Case .Parameters("@return").Value
                    Case 0 'return customer_price_info_id
                        '
                    Case Is > 0
                        If pIsProduce Then
                            _DepotId_Produce = Nz(Of Integer)(.Parameters("@depot_id").Value, 0)
                            _Depot_Produce = Nz(Of String)(.Parameters("@depot").Value, "")
                            _ServingAccNum_Produce = Nz(Of String)(.Parameters("@serv_code").Value, "")
                            _SupplierID_Produce = Nz(Of Long)(.Parameters("@supplier_id").Value, 0)
                            _Supplier_Produce = Nz(Of String)(.Parameters("@supplier").Value, "")
                        Else
                            _DepotId = Nz(Of Integer)(.Parameters("@depot_id").Value, 0)
                            _Depot = Nz(Of String)(.Parameters("@depot").Value, "")
                            _ServingAccNum = Nz(Of String)(.Parameters("@serv_code").Value, "")
                           _SupplierID = Nz(Of Long)(.Parameters("@supplier_id").Value, 0)
                             _Supplier = Nz(Of String)(.Parameters("@supplier").Value, "")
                        End If
                    Case Else ' other errors
                        RaiseEvent Report_Error("ERROR: " & .Parameters("@err_msg").Value & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)

                End Select
            End With

        Catch ex As Exception
            RaiseEvent Report_Error("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Finally

        End Try
    End Sub

    Private Function RetrieveOrderHeader(ByVal nOrderId? As Integer, ByVal strOrderNum As String, ByVal lngCustomerId? As Long, ByVal dtDeliveryDate As Date?) As Integer
        Return RetrieveOrderHeader(nOrderId, strOrderNum, lngCustomerId, dtDeliveryDate, _CheckIfAmendment)
    End Function

    'return order id found on the database - could be previous order for this customer which will be used as template to create new order 
    Private Function RetrieveOrderHeader(ByVal nOrderId? As Integer, ByVal strOrderNum As String, ByVal lngCustomerId? As Long, ByVal dtDeliveryDate As Date?, pIsAmendment As Boolean) As Integer
        Dim l_DB As New DB
        Dim objCmd As SqlClient.SqlCommand = Nothing
        Dim dr As SqlClient.SqlDataReader = Nothing
        Dim objParam As SqlClient.SqlParameter
        Dim intOrderId As Integer ' order id found on database for the customer on the same weekday as delivery date
        Dim bIsNewOrder As Boolean

        Try

            l_DB.Open()

            objCmd = New SqlClient.SqlCommand("p_customer_order_header_get", l_DB.Connection)
            objCmd.CommandType = CommandType.StoredProcedure

            'Create Parameters
            objParam = objCmd.Parameters.Add("@order_header_id", SqlDbType.Int)
            If nOrderId Is Nothing Then
                objParam.Value = DBNull.Value
            Else
                objParam.Value = nOrderId
            End If

            objParam = objCmd.Parameters.Add("@order_num", SqlDbType.VarChar, 30)
            If String.IsNullOrEmpty(strOrderNum) Then
                objParam.Value = DBNull.Value
            Else
                objParam.Value = strOrderNum
            End If

            objParam = objCmd.Parameters.Add("@customer_id", SqlDbType.BigInt)
            If lngCustomerId Is Nothing Then
                objParam.Value = DBNull.Value
            Else
                objParam.Value = lngCustomerId
            End If

            objParam = objCmd.Parameters.Add("@delivery_date", SqlDbType.Date)
            objParam.Direction = ParameterDirection.InputOutput
            If dtDeliveryDate Is Nothing Then
                objParam.Value = DBNull.Value
            Else
                objParam.Value = dtDeliveryDate
            End If

            objParam = objCmd.Parameters.Add("@IsAmendment", SqlDbType.Bit)
            objParam.Direction = ParameterDirection.Input
            objParam.Value = pIsAmendment

            dr = objCmd.ExecuteReader()

            If dr.HasRows Then
                dr.Read()
                bIsNewOrder = CType(dr("is_existing"), Boolean)
                intOrderId = CType(dr("coh_customer_order_header_id"), Integer)
                If bIsNewOrder Then
                    Me.OrderId = 0
                    TS = Nothing
                    OrderNum = Nothing
                    DepotId = 0
                    Depot = Nothing
                    DepotEmail = Nothing
                    DepotAltEmail = Nothing
                    DepotAltEmail2 = Nothing
                    Supplier = ""
                    ServingAccNum = Nothing
                    DeliveryDate = objCmd.Parameters("@delivery_date").Value
                    Notes = ""
                    LastUpdated = Nothing
                    UserId = 1
                    UserName = ""
                    EmailedDateTime = Nothing
                    BuyingGroupID = Nz(Of Integer)(dr("cu_buying_group_id"), 0)
                Else
                    Me.OrderId = intOrderId
                    TS = CType(dr("coh_ts"), Byte())
                    OrderNum = If(IsDBNull(dr("coh_order_num")), Nothing, dr("coh_order_num"))
                    DepotId = CType(dr("coh_depot_id"), Integer)
                    Depot = If(IsDBNull(dr("coh_depot")), Nothing, dr("coh_depot"))
                    DepotEmail = If(IsDBNull(dr("de_email")), Nothing, dr("de_email"))
                    DepotAltEmail = If(IsDBNull(dr("de_alt_email")), Nothing, dr("de_alt_email"))
                    DepotAltEmail2 = If(IsDBNull(dr("de_alt_email2")), Nothing, dr("de_alt_email2"))
                    Supplier = If(IsDBNull(dr("coh_supplier")), "", dr("coh_supplier"))
                    SupplierID = Nz(Of Integer)(dr("coh_supplier_id"), 0)
                    ServingAccNum = If(IsDBNull(dr("coh_serving_acc_num")), Nothing, dr("coh_serving_acc_num"))                    DepotId_Produce = Nz(Of Integer)(dr("coh_depot_id_produce"), 0)
                    Depot_Produce = Nz(Of String)(dr("coh_depot_produce"), Nothing)
                    DepotEmail_Produce = Nz(Of String)(dr("de_email_produce"), Nothing)
                    DepotAltEmail = Nz(Of String)(dr("de_alt_email_produce"), Nothing)
                    DepotAltEmail2 = Nz(Of String)(dr("de_alt_email2_produce"), Nothing)
                    Supplier_Produce = Nz(Of String)(dr("coh_supplier_produce"), "")
                    SupplierID_Produce = Nz(Of Integer)(dr("coh_supplier_id_produce"), 0)
                    ServingAccNum_Produce = Nz(Of String)(dr("coh_serving_acc_num_produce"), Nothing)
                    DeliveryDate = CType(dr("coh_delivery_date"), Date)
                    Notes = If(IsDBNull(dr("coh_notes")), "", dr("coh_notes"))
                    LastUpdated = CType(dr("coh_date_last_update"), Date)
                    UserId = CType(dr("userId"), Integer)
                    UserName = dr("person_updated")
                    EmailedDateTime = If(IsDBNull(dr("coh_date_emailed")), Nothing, dr("coh_date_emailed"))
                    EmailedDateTime_Produce = Nz(Of Date)(dr("coh_date_emailed_produce"), Nothing)
                    DeliveryNoteID = Nz(Of Integer)(dr("pdh_record_id"), 0)
                    If pIsAmendment Then
                        IsSoAmendment = CType(If(IsDBNull(dr("IsAmendment")), 0, dr("IsAmendment")), Boolean)
                    End If
                    BuyingGroupID = Nz(Of Integer)(dr("cu_buying_group_id"), 0)
                End If

                CustomerId = CType(dr("coh_sage_customer_id"), Long)
                HOId = CType(dr("cu_head_office_id"), Long)
                StoreId = If(IsDBNull(dr("cu_store_number")), Nothing, dr("cu_store_number").ToString.Trim)
                Customer = If(IsDBNull(dr("coh_customer")), Nothing, dr("coh_customer"))
				_ASNSent = Nz(Of Boolean)(dr("coh_ASN_sent"), False)

            Else
				GetCustomerDepots(lngCustomerId, dtDeliveryDate)
                Return 0
            End If

            dr.Close()
            objCmd.Dispose()
            l_DB.Close()

            Return intOrderId

        Catch ex As Exception
            RaiseEvent Report_Error("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)

            Return 0
        Finally
            If Not dr.IsClosed Then
                dr.Close()
            End If
        End Try

    End Function

    Public Function RetrieveOrderLines(ByVal nOrderId As Integer) As Boolean
        Dim l_DB As New DB
        Dim objCmd As SqlClient.SqlCommand = Nothing
        Dim objRdr As SqlClient.SqlDataReader = Nothing
        Dim objParam As SqlClient.SqlParameter

        Try
            l_DB.Open()

            objCmd = New SqlClient.SqlCommand("p_customer_order_lines_get", l_DB.Connection)
            objCmd.CommandType = CommandType.StoredProcedure

            'Create Parameters
            objParam = objCmd.Parameters.Add("@order_header_id", SqlDbType.Int)
            objParam.Value = nOrderId

            If IsSoAmendment Then
                objParam = objCmd.Parameters.Add("@amendment", SqlDbType.Bit)
                objParam.Value = 1
            End If

            objRdr = objCmd.ExecuteReader()

            If objRdr.HasRows Then
                While objRdr.Read()
                    'If OrderId = 0 Then
                    '    'New Order
                    '    Me.Add(0, _
                    '            Nothing, _
                    '            CType(objRdr("cod_product_id"), Integer), _
                    '            objRdr("product_code").ToString,
                    '            objRdr("product").ToString,
                    '            CType(objRdr("cod_qty"), Decimal),
                    '            CType(objRdr("cod_so_qty"), Decimal)
                    '        )
                    'Else
                    'Existing order

                    Me.Add(CType(objRdr("cod_record_id"), Integer), _
                            CType(objRdr("cod_ts"), Byte()), _
                            CType(objRdr("cod_product_id"), Integer), _
                            objRdr("product_code").ToString,
                            objRdr("product").ToString,
                            CType(objRdr("cod_qty"), Decimal),
                            CType(objRdr("cod_so_qty"), Decimal)
                        )
                    Me.Item(objRdr("product_code").ToString).SupplyingDepotID = Nz(Of Long)(objRdr("cod_supplying_depot_id"), DepotId)
                    'End If
                    Me.Item(objRdr("product_code").ToString).IsProduce = CType(objRdr("IsProduce"), Boolean)
                End While
            Else
                Return False
            End If
            objRdr.Close()
            objCmd.Dispose()
            l_DB.Close()

            Return True

        Catch ex As Exception
            RaiseEvent Report_Error("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
            Return False
        Finally

        End Try

    End Function

    Public Function RetrieveStandingOrder(ByVal pOrderId As Integer) As Boolean
        Dim l_DB As New DB
        Dim lCmd As SqlClient.SqlCommand
        Dim lPapam As SqlClient.SqlParameter
        Dim lPrevProductID As Integer = 0
        Dim lOrderLine As COrderDetail = Nothing
        Try

            l_DB.Open()


            lCmd = New SqlClient.SqlCommand("p_standing_order_get", l_DB.Connection)
            lCmd.CommandType = CommandType.StoredProcedure

            'Create Parameters
            lPapam = lCmd.Parameters.Add("@cust_id", SqlDbType.BigInt)
            lPapam.Value = 0
            lPapam = lCmd.Parameters.Add("@date_effective", SqlDbType.Date)
            lPapam.Value = DBNull.Value
            'the above two parameters are irrelevant here

            lPapam = lCmd.Parameters.Add("@order_id", SqlDbType.Int)
            lPapam.Value = pOrderId

            With lCmd.ExecuteReader()
                If .HasRows Then
                    If .Read Then
                        OrderId = .Item("soh_record_id")
                        OrderNum = Nz(Of String)(.Item("soh_order_num"), "")
                        Notes = Nz(Of String)(.Item("soh_notes"), "")
                        DateEffective = Nz(Of Date)(CType(.Item("soh_date_effective"), Date), Now.Date)
                        IsSOCancelled = Nz(Of Boolean)(.Item("soh_stopped"), False)
                        IsSuspended = Nz(Of Boolean)(.Item("is_suspended"), False)
                        If IsSOCancelled Then
                            SOLastDeliveryDate = DateAdd(DateInterval.Day, -1, Nz(Of Date)(CType(.Item("soh_date_stopped"), Date), Now.Date))
                        ElseIf IsSuspended Then
                            SOLastDeliveryDate = Nz(Of Date)(CType(.Item("suspended_until"), Date), Date.MinValue)
                        Else
                            SOLastDeliveryDate = Date.MinValue
                        End If
                        _CustomerId = .Item("cust_id")
                        _CustAcc = Nz(Of String)(.Item("acc_num"), "")
                        DepotId = Nz(Of Integer)(.Item("depot_id"), 0)
                        ServingAccNum = Nz(Of String)(.Item("serving_code"), "").Trim
                        DeliveryBreakReason = Nz(Of String)(.Item("delivery_break_reason"), "").Trim

                        .NextResult()
                        While .Read
                            If lPrevProductID <> .Item("sod_product_id") Then
                                If Not lOrderLine Is Nothing Then
                                    Me.Add(lOrderLine)
                                End If
                                lOrderLine = New COrderDetail(Me, 0, Nothing, .Item("sod_product_id"), .Item("pr_code"), .Item("pr_product"), 0, 0, 0, 0, 0, 0, 0)
                                lPrevProductID = .Item("sod_product_id")
                            End If
                            Select Case .Item("sod_wkday")
                                Case 1 : lOrderLine.Sun = .Item("sod_qty")
                                Case 2 : lOrderLine.Mon = .Item("sod_qty")
                                Case 3 : lOrderLine.Tue = .Item("sod_qty")
                                Case 4 : lOrderLine.Wed = .Item("sod_qty")
                                Case 5 : lOrderLine.Thu = .Item("sod_qty")
                                Case 6 : lOrderLine.Fri = .Item("sod_qty")
                                Case 7 : lOrderLine.Sat = .Item("sod_qty")
                            End Select
                        End While
                        If Not lOrderLine Is Nothing Then
                            Me.Add(lOrderLine)
                        End If
                    End If
                Else
                    Return False
                End If
                .Close()
            End With

            lCmd.Dispose()
            l_DB.Close()
            Return True

        Catch ex As Exception
            RaiseEvent Report_Error("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
            Return False
        Finally

        End Try
    End Function

    Public Sub UpdateDateEmailed(ByVal dtDateTime As Date)
        UpdateDateEmailed(dtDateTime, 0)
    End Sub

    Public Sub UpdateDateEmailed(ByVal dtDateTime As Date, pIsProduce As Boolean)
        Dim l_DB As New DB
        Dim lCmd As SqlClient.SqlCommand

        Try
            l_DB.Open()
            With New System.Text.StringBuilder
                If IsStandingOrder Then
                    .AppendLine("UPDATE standing_order_headers SET soh_date_last_emailed = GETDATE() ")
                    .AppendLine("WHERE soh_record_id = @order_id")
                Else
                    .AppendLine("UPDATE customer_order_headers SET ")
                    If pIsProduce Then
                        .AppendLine(" coh_date_emailed_produce = GETDATE() ")
                    Else
                        .AppendLine(" coh_date_emailed = GETDATE() ")
                    End If
                    .AppendLine("WHERE coh_customer_order_header_id = @order_id")
                End If

                lCmd = New SqlClient.SqlCommand(.ToString, l_DB.Connection)
            End With
            lCmd.Parameters.AddWithValue("@order_id", OrderId)
            lCmd.ExecuteNonQuery()

            lCmd.Dispose()
            l_DB.Close()
            If pIsProduce Then
                _EmailedDateTime_Produce = dtDateTime
            Else
                _EmailedDateTime = dtDateTime
            End If
            _EmailedDateTime = dtDateTime

        Catch ex As Exception
            RaiseEvent Report_Error("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Finally

        End Try

    End Sub

    Private Function Nz(Of objDataType)(ByVal objVal As Object, ByVal objRet As objDataType) As objDataType
        If IsDBNull(objVal) OrElse IsNothing(objVal) Then
            Return objRet
        Else
            Return CType(objVal, objDataType)
        End If
    End Function
End Class







