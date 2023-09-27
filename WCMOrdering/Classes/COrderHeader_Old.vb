Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.Specialized

Public Class COrderHeader
    Inherits CollectionBase

    Private mCn As SqlClient.SqlConnection

    Private mnOrderID As Integer = 0
    Private mbytTS As Byte()
    Private mstrOrderNum As String = Nothing
    Private mlngCustomerId As Long = 0
    Private mlngHOId As Long = 0
    Private mstrStoreId As String = Nothing ' customer store num ( acc num)
    Private mstrCustomer As String = Nothing
    Private mstrCustAcc As String = Nothing
    Private mstrAddress1 As String = Nothing
    Private mstrPostcode As String = Nothing
    Private mstrRestrictedSalesGroup As String = Nothing
    Private mintDepotId As Integer = 0
    Private mstrDepot As String = Nothing
    Private mstrDepotEmail As String = Nothing
    Private mstrDepotAltEmail As String = Nothing
    Private mstrDepotAltEmail2 As String = Nothing
    Private mstrSupplier As String = Nothing
    Private mstrServingAccNum As String = Nothing
    Private mdtDeliveryDate As Date = Nothing
    Private mdtDateEffectiveSun As Date = Nothing ' SUSPENDED used for standing order (week_ending effective - 6 days)
    Private mdtDateEffective As Date = Nothing ' 
    Private mstrNotes As String = ""
    Private mintUserId As Integer? = Nothing
    Private mstrUserName As String = Nothing
    Private mdtLastUpdated As Date
    Private mdtEmailedDateTime As Date? = Nothing

    Public IsSoAmendment As Boolean = False ' standing order amendment flag, true when created from standing order

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

        mdtDateEffectiveSun = dtDate
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

    Public Sub New(ByVal dtDate As Date)
        mdtDateEffective = dtDate
        SetDateEffectiveSunday(dtDate)
    End Sub

    Public Sub New(ByVal intOrderId As Integer)
        GetOrder(intOrderId)
    End Sub

    Public Sub New(ByVal pCreateFromStandingOrder As String, ByVal lngCustomerId As Long, ByVal dtDeliveryDate As Date)
        Me.IsSoAmendment = True
        GetOrder(lngCustomerId, dtDeliveryDate)
        mdtDateEffective = dtDeliveryDate
        SetDateEffectiveSunday(dtDeliveryDate)
    End Sub

    Public Sub New(ByVal lngCustomerId As Long, ByVal dtDeliveryDate As Date)
        GetOrder(lngCustomerId, dtDeliveryDate)
    End Sub

    Public Sub New(ByVal pOrderNum As String, ByVal pDeliveryDate As Date)
        GetOrder(pOrderNum, pDeliveryDate)
    End Sub

    Public Property ConnDB() As SqlClient.SqlConnection
        Get
            Return mCn
        End Get
        Set(ByVal value As SqlClient.SqlConnection)
            mCn = value
        End Set
    End Property

    Public Property OrderId() As Integer
        Get
            Return mnOrderID
        End Get
        Set(ByVal value As Integer)
            mnOrderID = value
        End Set
    End Property

    Public Property CustomerId() As Long
        Get
            Return mlngCustomerId
        End Get
        Set(ByVal value As Long)
            mlngCustomerId = value
        End Set
    End Property

    Public Property HOId() As Long
        Get
            Return mlngHOId
        End Get
        Set(ByVal value As Long)
            mlngHOId = value
        End Set
    End Property

    Public Property Customer() As String
        Get
            If mstrCustomer Is Nothing Then
                GetCustomerInfo()
            End If
            Return mstrCustomer
        End Get
        Set(ByVal value As String)
            mstrCustomer = value
        End Set
    End Property

    Public Property CustAcc() As String
        Get
            If mstrCustAcc Is Nothing Then
                GetCustomerInfo()
            End If
            Return mstrCustAcc
        End Get
        Set(ByVal value As String)
            mstrCustAcc = value
        End Set
    End Property
    Public Property DepotId() As Integer
        Get
            Return mintDepotId
        End Get
        Set(ByVal value As Integer)
            mintDepotId = value
        End Set
    End Property

    Public Property Depot() As String
        Get
            If mstrDepot Is Nothing Then
                GetDepotInfo()
            End If
            Return mstrDepot
        End Get
        Set(ByVal value As String)
            mstrDepot = value
        End Set
    End Property

    Public Property DepotEmail() As String
        Get
            If mstrDepotEmail Is Nothing Then
                GetDepotInfo()
            End If
            Return mstrDepotEmail
        End Get
        Set(ByVal value As String)
            mstrDepotEmail = value
        End Set
    End Property

    Public Property DepotAltEmail() As String
        Get
            If mstrDepotAltEmail Is Nothing Then
                GetDepotInfo()
            End If
            Return mstrDepotAltEmail
        End Get
        Set(ByVal value As String)
            mstrDepotAltEmail = value
        End Set
    End Property

    Public Property DepotAltEmail2() As String
        Get
            If mstrDepotAltEmail2 Is Nothing Then
                GetDepotInfo()
            End If
            Return mstrDepotAltEmail2
        End Get
        Set(ByVal value As String)
            mstrDepotAltEmail2 = value
        End Set
    End Property

    Public Property Supplier() As String
        Get
            If mstrSupplier Is Nothing Then
                GetCurrentSupplier()
            End If
            Return mstrSupplier
        End Get
        Set(ByVal value As String)
            mstrSupplier = value
        End Set
    End Property

    Public Property ServingAccNum() As String
        Get
            If mstrServingAccNum Is Nothing Then
                GetServingAccNum()
            End If
            Return mstrServingAccNum
        End Get
        Set(ByVal value As String)
            mstrServingAccNum = value
        End Set
    End Property

    Public Property DateEffective() As Date
        Get
            Return mdtDateEffective
        End Get
        Set(ByVal value As Date)
            mdtDateEffective = value
        End Set
    End Property

    Public Property DateEffectiveSun() As Date
        Get
            Return mdtDateEffectiveSun
        End Get
        Set(ByVal value As Date)
            mdtDateEffectiveSun = value
        End Set
    End Property

    Public Property DeliveryDate() As Date
        Get
            Return mdtDeliveryDate
        End Get
        Set(ByVal value As Date)
            mdtDeliveryDate = value
        End Set
    End Property

    Public Property Notes() As String
        Get
            Return mstrNotes
        End Get
        Set(ByVal value As String)
            mstrNotes = value
        End Set
    End Property

    Public Property OrderNum() As String
        Get
            If String.IsNullOrEmpty(mstrOrderNum) Then
                Return CustAcc & Format(mdtDeliveryDate, "yyMMdd")
            Else
                Return mstrOrderNum
            End If

        End Get
        Set(ByVal value As String)
            mstrOrderNum = value
        End Set
    End Property
    Public Property LastUpdated() As Date
        Get
            Return mdtLastUpdated
        End Get
        Set(ByVal value As Date)
            mdtLastUpdated = value
        End Set
    End Property

    Public Property UserId() As Integer
        Get
            If mintUserId Is Nothing Then
                mintUserId = modCurrentUser.UserId
            End If
            Return mintUserId
        End Get
        Set(ByVal value As Integer)
            mintUserId = value
        End Set
    End Property

    Public Property EmailedDateTime() As Date?
        Get
            Return mdtEmailedDateTime
        End Get
        Set(ByVal value As Date?)
            mdtEmailedDateTime = value
        End Set
    End Property

    Public Property StoreId() As String
        Get
            If mstrStoreId Is Nothing Then
                GetCustomerInfo()
            End If
            Return mstrStoreId
        End Get
        Set(ByVal value As String)
            mstrStoreId = value.Trim()
        End Set
    End Property

    Public Property Address1() As String
        Get
            If mstrAddress1 Is Nothing Then
                GetCustomerInfo()
            End If
            Return mstrAddress1
        End Get
        Set(ByVal value As String)
            mstrAddress1 = value.Trim()
        End Set
    End Property

    Public Property Postcode() As String
        Get
            If mstrPostcode Is Nothing Then
                GetCustomerInfo()
            End If
            Return mstrPostcode
        End Get
        Set(ByVal value As String)
            mstrPostcode = value.Trim()
        End Set
    End Property

    Public Property RestrictedSalesGroup() As String
        Get
            If mstrRestrictedSalesGroup Is Nothing Then
                GetCustomerInfo()
            End If
            Return mstrRestrictedSalesGroup
        End Get
        Set(ByVal value As String)
            mstrRestrictedSalesGroup = value.Trim()
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
            If mstrUserName Is Nothing Then
                mstrUserName = modCurrentUser.Fullname
            End If
            Return mstrUserName
        End Get
        Set(ByVal value As String)
            mstrUserName = value
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
    'Get Latest Order by Customer ID - on or before delivery date, if delivery date not specified get the latest customer order for current weekday
    Public Function GetOrder(ByVal lngCustomerId As Long, ByVal dtDeliveryDate As Date?) As Boolean
        Dim intOrderId As Integer
        Dim bGotLines As Boolean = False

        intOrderId = RetrieveOrderHeader(Nothing, Nothing, lngCustomerId, dtDeliveryDate)
        If Me.IsSoAmendment Then
            'for the existing amendment Me.OrderId will not be zero
            intOrderId = Me.OrderId
        End If
        If intOrderId <> 0 Then
            bGotLines = RetrieveOrderLines(intOrderId)

        End If
        If bGotLines Then
            Return True
        Else
            If Year(Me.DeliveryDate) < 2000 Then
                Me.DeliveryDate = dtDeliveryDate
                Me.CustomerId = lngCustomerId
            End If
            If Year(Me.DeliveryDate) > 2000 Then
                Return CreateOrderLinesFromInvoice(lngCustomerId, Me.DeliveryDate) ' NOTE: use Me.Delivery date because dtDeliveryDate parameter could be nothing so delivery date will be set in RetrieveOrderHeader
            Else
                Return False
            End If
        End If
    End Function

    Private Sub GetCustomerInfo()
        Dim cmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim clsDb As DB
        Dim sbsql As New System.Text.StringBuilder
        Try
            If mlngCustomerId <> 0 Then
                clsDb = New DB
                clsDb.Open()

                With sbsql
                    .AppendLine("SELECT customer_account_number, isnull(SiteName,customer_account_name) customer, cu_store_number, address_line1, postcode, isnull(rsg_restricted_sales_group,'') restricted_sales_group")
                    .AppendLine(" FROM v_customer_links  ")
                    .AppendLine(" LEFT JOIN restricted_sales_groups ON rsg_restricted_sales_group_id = cu_restricted_sales_group_id ")
                    .Append(" WHERE cu_sage_customer_id =  ") : .AppendLine(mlngCustomerId.ToString)

                    cmd = New SqlClient.SqlCommand(.ToString, clsDb.Connection)
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
            Else
                mstrCustomer = ""
                mstrStoreId = ""
            End If
        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "GetCustomerInfo")

        Finally
            If Not dr Is Nothing AndAlso Not dr.IsClosed Then
                dr.Close()
            End If
        End Try
    End Sub

    Public Sub GetDepotInfo()
        Dim cmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim clsDb As DB
        Dim sbsql As New System.Text.StringBuilder

        Try
            If mintDepotId <> 0 Then
                clsDb = New DB
                clsDb.Open()

                With sbsql
                    .AppendLine("SELECT de_depot, de_email, de_alt_email, de_alt_email2")
                    .AppendLine("FROM depots ")
                    .Append("WHERE de_depot_id = ") : .Append(mintDepotId.ToString)

                    cmd = New SqlClient.SqlCommand(.ToString, clsDb.Connection)
                End With
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    If dr.Read Then
                        Me.Depot = dr("de_depot").ToString
                        Me.DepotEmail = Nz(Of String)(dr("de_email"), "")
                        Me.DepotAltEmail = Nz(Of String)(dr("de_alt_email"), "")
                        Me.DepotAltEmail2 = Nz(Of String)(dr("de_alt_email2"), "")
                    End If
                End If
            Else
                mstrDepot = ""
                mstrDepotEmail = Nothing
                mstrDepotAltEmail = Nothing
                mstrDepotAltEmail2 = Nothing
            End If
        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "GetDepotInfo")

        Finally
            If Not dr Is Nothing AndAlso Not dr.IsClosed Then
                dr.Close()
            End If
        End Try
    End Sub

    Public Sub GetServingAccNum()
        Dim cmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim clsDb As DB
        Dim sbsql As New System.Text.StringBuilder
        Try
            If mlngCustomerId <> 0 AndAlso mintDepotId <> 0 Then
                clsDb = New DB
                clsDb.Open()

                With sbsql
                    .AppendLine("SELECT cud_serving_acc_num")
                    .AppendLine("FROM customer_depots ")
                    .Append("WHERE cud_sage_customer_id = ") : .AppendLine(mlngCustomerId.ToString)
                    .Append("AND cud_depot_id = ") : .Append(mintDepotId.ToString)

                    cmd = New SqlClient.SqlCommand(.ToString, clsDb.Connection)
                End With
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    If dr.Read Then
                        Me.ServingAccNum = dr("cud_serving_acc_num").ToString
                    End If
                End If
            Else
                mstrServingAccNum = ""
            End If
        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "GetServingAccNum")

        Finally
            If Not dr Is Nothing AndAlso Not dr.IsClosed Then
                dr.Close()
            End If
        End Try
    End Sub

    Public Sub GetCurrentSupplier()
        Dim cmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim clsDb As DB
        Dim sbsql As New System.Text.StringBuilder
        Try
            If mlngCustomerId <> 0 Then
                clsDb = New DB
                clsDb.Open()

                With sbsql


                    .AppendLine("SELECT SupplierAccountName")
                    .AppendLine("FROM v_suppliers ")
                    .AppendLine("INNER JOIN bandings ON bn_parent_id = PLSupplierAccountID")
                    .AppendLine("INNER JOIN customer_price_info ON cpi_cost_price_banding_id = bn_banding_id")
                    .Append("WHERE cpi_customer_id =  ") : .AppendLine(mlngCustomerId.ToString)
                    cmd = New SqlClient.SqlCommand(.ToString, clsDb.Connection)
                End With
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    If dr.Read Then
                        Me.Supplier = dr("SupplierAccountName").ToString
                    End If
                End If
            Else
                mstrSupplier = ""
            End If
        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "GetCurrentSupplier")

        Finally
            If Not dr Is Nothing AndAlso Not dr.IsClosed Then
                dr.Close()
            End If
        End Try
    End Sub

    'return order id found on the database - could be previous order for this customer which will be used as template to create new order 
    Private Function RetrieveOrderHeader(ByVal nOrderId? As Integer, ByVal strOrderNum As String, ByVal lngCustomerId? As Long, ByVal dtDeliveryDate As Date?) As Integer
        Dim objCmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader = Nothing
        Dim objParam As SqlClient.SqlParameter
        Dim intOrderId As Integer ' order id found on database for the customer on the same weekday as delivery date
        Dim bIsNewOrder As Boolean

        Try
            If ConnDB Is Nothing Then
                Dim clsDB As New DB
                clsDB.Open()
                ConnDB = clsDB.Connection
            End If

            objCmd = New SqlClient.SqlCommand("p_customer_order_header_get", ConnDB)
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
                    UserId = modCurrentUser.UserId
                    UserName = ""
                    EmailedDateTime = Nothing
                Else
                    Me.OrderId = intOrderId
                    TS = CType(dr("coh_ts"), Byte())
                    OrderNum = Nz(Of String)(dr("coh_order_num"), Nothing)
                    DepotId = CType(dr("coh_depot_id"), Integer)
                    Depot = Nz(Of String)(dr("coh_depot"), Nothing)
                    DepotEmail = Nz(Of String)(dr("de_email"), Nothing)
                    DepotAltEmail = Nz(Of String)(dr("de_alt_email"), Nothing)
                    DepotAltEmail2 = Nz(Of String)(dr("de_alt_email2"), Nothing)
                    Supplier = Nz(Of String)(dr("coh_supplier"), "")
                    ServingAccNum = Nz(Of String)(dr("coh_serving_acc_num"), Nothing)
                    DeliveryDate = CType(dr("coh_delivery_date"), Date)
                    Notes = Nz(Of String)(dr("coh_notes"), "")
                    LastUpdated = CType(dr("coh_date_last_update"), Date)
                    UserId = CType(dr("userId"), Integer)
                    UserName = dr("person_updated")
                    EmailedDateTime = Nz(Of Date)(dr("coh_date_emailed"), Nothing)
                End If

                CustomerId = CType(dr("coh_sage_customer_id"), Long)
                HOId = CType(dr("cu_head_office_id"), Long)
                StoreId = Nz(Of String)(dr("cu_store_number").ToString.Trim, Nothing)
                Customer = Nz(Of String)(dr("coh_customer"), Nothing)

            Else
                Return 0
            End If

            Return intOrderId

        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "RetrieveOrderHeader")
            Return 0
        Finally
            If Not dr.IsClosed Then
                dr.Close()
            End If
        End Try

    End Function

    Public Function RetrieveOrderLines(ByVal nOrderId As Integer) As Boolean
        Dim objCmd As SqlClient.SqlCommand
        Dim objRdr As SqlClient.SqlDataReader = Nothing
        Dim objParam As SqlClient.SqlParameter
        Try
            If ConnDB Is Nothing Then
                Dim clsDB As New DB
                clsDB.Open()
                ConnDB = clsDB.Connection
            End If

            objCmd = New SqlClient.SqlCommand("p_customer_order_lines_get", ConnDB)
            objCmd.CommandType = CommandType.StoredProcedure

            'Create Parameters
            objParam = objCmd.Parameters.Add("@order_header_id", SqlDbType.Int)
            objParam.Value = nOrderId

            objRdr = objCmd.ExecuteReader()

            If objRdr.HasRows Then
                While objRdr.Read()
                    If OrderId = 0 Then
                        'New Order
                        Me.Add(0, _
                                Nothing, _
                                CType(objRdr("cod_product_id"), Integer), _
                                objRdr("product_code").ToString,
                                objRdr("product").ToString,
                                CType(objRdr("cod_qty"), Decimal),
                                CType(objRdr("cod_so_qty"), Decimal)
                            )
                    Else
                        'Existing order

                        Me.Add(CType(objRdr("cod_record_id"), Integer), _
                                CType(objRdr("cod_ts"), Byte()), _
                                CType(objRdr("cod_product_id"), Integer), _
                                objRdr("product_code").ToString,
                                objRdr("product").ToString,
                                CType(objRdr("cod_qty"), Decimal),
                                CType(objRdr("cod_so_qty"), Decimal)
                            )
                    End If
                End While
            Else
                Return False
            End If

            Return True

        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "RetrieveOrderLines")
            Return False
        Finally
            If Not objRdr.IsClosed Then
                objRdr.Close()
            End If
        End Try

    End Function

    Public Function CreateOrderLinesFromInvoice(ByVal lngCustomerId As Long, ByVal dtDeliveryDate As Date) As Boolean
        Dim objCmd As SqlClient.SqlCommand
        Dim objRdr As SqlClient.SqlDataReader = Nothing
        Dim objParam As SqlClient.SqlParameter

        Try
            If ConnDB Is Nothing Then
                Dim clsDB As New DB
                clsDB.Open()
                ConnDB = clsDB.Connection
            End If

            ' NOTE - this will first attempt to get product list from the latest standing order, if not exists then from latest invoice for this weekday
            objCmd = New SqlClient.SqlCommand("p_customer_order_lines_from_latest_invoice", ConnDB)
            objCmd.CommandType = CommandType.StoredProcedure

            'Create Parameters
            objParam = objCmd.Parameters.Add("@customer_id", SqlDbType.BigInt)
            objParam.Value = lngCustomerId

            objParam = objCmd.Parameters.Add("@delivery_date", SqlDbType.Date)
            objParam.Value = dtDeliveryDate

            objRdr = objCmd.ExecuteReader()

            If objRdr.HasRows Then
                While objRdr.Read()
                    Me.Add(0, Nothing, _
                            CType(objRdr("id_product_id"), Integer), _
                            objRdr("product_code").ToString,
                            objRdr("product").ToString,
                            CType(objRdr("qty"), Decimal),
                            CType(objRdr("so_qty"), Decimal)
                        )
                End While
            Else
                Return False
            End If

            Return True

        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "CreateOrderLinesFromInvoice")
            Return False
        Finally
            If Not objRdr.IsClosed Then
                objRdr.Close()
            End If
        End Try

    End Function

    Public Function SaveOrder() As Boolean
        If ConnDB Is Nothing Then
            Dim clsDB As New DB
            clsDB.Open()
            ConnDB = clsDB.Connection
        End If
        If SaveOrderHeader(ConnDB) Then
            Return SaveOrderLines(ConnDB)
        Else
            Return False
        End If
    End Function

    Public Function SaveOrderHeader(ByVal ConnDB As SqlClient.SqlConnection) As Boolean
        Dim objCmd As SqlClient.SqlCommand
        Dim objParam As SqlClient.SqlParameter

        Try

            objCmd = New SqlClient.SqlCommand("p_customer_order_header_write", ConnDB)
            With objCmd
                .CommandType = CommandType.StoredProcedure
                'Create Parameters
                objParam = .Parameters.Add("@return", SqlDbType.Int)
                objParam.Direction = ParameterDirection.ReturnValue

                objParam = .Parameters.Add("@order_header_id", SqlDbType.Int)
                objParam.Direction = ParameterDirection.InputOutput
                objParam.Value = OrderId

                objParam = .Parameters.Add("@ts", SqlDbType.Timestamp)
                objParam.Direction = ParameterDirection.InputOutput
                objParam.Value = IIf(TS Is Nothing, System.DBNull.Value, TS)

                objParam = .Parameters.Add("@order_num", SqlDbType.VarChar, 30)
                objParam.Direction = ParameterDirection.InputOutput
                objParam.Value = IIf(OrderNum Is Nothing, System.DBNull.Value, OrderNum)

                objParam = .Parameters.Add("@customer_id", SqlDbType.BigInt)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = CustomerId

                objParam = .Parameters.Add("@depot_id", SqlDbType.Int)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = DepotId

                objParam = .Parameters.Add("@supplier", SqlDbType.VarChar, 30)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = Supplier

                objParam = .Parameters.Add("@delivery_date", SqlDbType.Date)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = DeliveryDate

                objParam = .Parameters.Add("@notes", SqlDbType.VarChar, 2000)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = Notes

                objParam = .Parameters.Add("@current_user_id", SqlDbType.Int)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = UserId

                objParam = .Parameters.Add("@date_last_update", SqlDbType.DateTime)
                objParam.Direction = ParameterDirection.Output

                objParam = .Parameters.Add("@person_updated", SqlDbType.VarChar, 30)
                objParam.Direction = ParameterDirection.Output

                .ExecuteNonQuery()

                Select Case .Parameters("@return").Value
                    Case 0
                        OrderId = CType(.Parameters("@order_header_id").Value, Integer)
                        TS = CType(.Parameters("@ts").Value, Byte())
                        OrderNum = Nz(Of String)(.Parameters("@order_num").Value, Nothing)
                        LastUpdated = CType(.Parameters("@date_last_update").Value, Date)
                        UserName = .Parameters("@person_updated").Value.ToString
                    Case 1 ' timestamp
                        MsgBox("An update has taken place on this data since it was last loaded.", MsgBoxStyle.Information, OrderNum)
                        Return False
                    Case 2 ' already exists with this name
                        MsgBox("Update failed. Please try again." & vbCrLf & "If the problem persists please contact support.", MsgBoxStyle.Information, Nz(Of String)(OrderNum, ""))
                        Return False
                    Case 3 ' already exists with this name
                        MsgBox("Insert failed. Please try again." & vbCrLf & "If the problem persists please contact support.", MsgBoxStyle.Information, Nz(Of String)(OrderNum, ""))
                        Return False
                    Case Else ' other errors
                        ' will be reported by exception handler
                        Return False
                End Select

            End With

            Return True

        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "SaveOrderHeader")
            Return False
        End Try
    End Function

    Public Function SaveOrderLines(ByRef conDB As SqlClient.SqlConnection) As Boolean
        Dim objCmd As SqlClient.SqlCommand = PrepareDetailsSave(conDB)
        Dim clsLine As COrderDetail

        Try

            If Not objCmd Is Nothing Then
                For Each clsLine In Me.List
                    With objCmd.Parameters
                        .Item("@ts").Value = IIf(clsLine.TS Is Nothing, System.DBNull.Value, clsLine.TS)
                        .Item("@record_id").Value = clsLine.ID
                        .Item("@product_id").Value = clsLine.ProductId
                        .Item("@qty").Value = clsLine.Qty
                        If clsLine.SOQty = 0 Then
                            .Item("@so_qty").Value = DBNull.Value
                        Else
                            .Item("@so_qty").Value = clsLine.SOQty
                        End If

                        If String.IsNullOrEmpty(clsLine.ProductCode) Then
                            .Item("@code").Value = DBNull.Value
                        Else
                            .Item("@code").Value = clsLine.ProductCode
                        End If

                        If String.IsNullOrEmpty(clsLine.Product) Then
                            .Item("@product").Value = DBNull.Value
                        Else
                            .Item("@product").Value = clsLine.Product
                        End If

                        objCmd.ExecuteNonQuery()

                        Select Case .Item("@return").Value
                            Case 0
                                If clsLine.ProductId > 0 Then
                                    clsLine.ID = CType(.Item("@record_id").Value, Integer)
                                    clsLine.TS = CType(.Item("@ts").Value, Byte())
                                    clsLine.ProductCode = .Item("@code").Value.ToString
                                    clsLine.Product = .Item("@product").Value.ToString
                                Else
                                    'items deleted, has be to removed from the order
                                    clsLine.ID = 0

                                End If
                            Case 1 ' timestamp
                                MsgBox("An update has taken place on this data since it was last loaded.", MsgBoxStyle.Information, OrderNum & " - " & clsLine.ProductCode & " - " & clsLine.Product)
                                Return False
                            Case 2 ' update failed
                                MsgBox("Failed to update. Please try again." & vbCrLf & "If the problem persists please contact support.", MsgBoxStyle.Information, OrderNum & " - " & clsLine.ProductCode & " - " & clsLine.Product)
                                Return False
                            Case 3 ' Insert failed
                                MsgBox("Failed to Insert. Please try again." & vbCrLf & "If the problem persists please contact support.", MsgBoxStyle.Information, OrderNum & " - " & clsLine.ProductCode & " - " & clsLine.Product)
                                Return False
                            Case Else ' other errors
                                ' will be reported by exception handler
                                Return False
                        End Select
                    End With

                Next

                Me.Clean() 'to remove deleted items with id = 0

                Return True
            End If
        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "SaveOrderLines")
            Return False
        End Try
    End Function

    Private Function PrepareDetailsSave(ByRef conDB As SqlClient.SqlConnection) As SqlClient.SqlCommand

        Dim objCmd As SqlClient.SqlCommand = Nothing
        Dim objParam As SqlClient.SqlParameter

        Try
            objCmd = New SqlClient.SqlCommand("p_customer_order_detail_write", conDB)

            If Not objCmd Is Nothing Then
                With objCmd
                    .CommandType = CommandType.StoredProcedure

                    'Create Parameters
                    objParam = .Parameters.Add("@return", SqlDbType.Int)
                    objParam.Direction = ParameterDirection.ReturnValue

                    objParam = .Parameters.Add("@record_id", SqlDbType.Int)
                    objParam.Direction = ParameterDirection.InputOutput

                    objParam = .Parameters.Add("@ts", SqlDbType.Timestamp)
                    objParam.Direction = ParameterDirection.InputOutput


                    objParam = objCmd.Parameters.Add("@header_id", SqlDbType.Int)
                    objParam.Direction = ParameterDirection.Input
                    'assign header id here as it will be useD for all detail lines
                    objParam.Value = Me.OrderId

                    objParam = objCmd.Parameters.Add("@product_id", SqlDbType.Int)
                    objParam.Direction = ParameterDirection.Input

                    objParam = objCmd.Parameters.Add("@qty", SqlDbType.Decimal)
                    objParam.Direction = ParameterDirection.Input

                    objParam = objCmd.Parameters.Add("@so_qty", SqlDbType.Decimal)
                    objParam.Direction = ParameterDirection.Input

                    objParam = objCmd.Parameters.Add("@code", SqlDbType.VarChar, 16)
                    objParam.Direction = ParameterDirection.InputOutput

                    objParam = objCmd.Parameters.Add("@product", SqlDbType.VarChar, 30)
                    objParam.Direction = ParameterDirection.InputOutput
                End With
            End If

            Return objCmd
        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "PrepareDetailsSave")
            Return Nothing
        Finally

        End Try
    End Function

    Public Sub UpdateDateEmailed(ByVal dtDateTime As Date)
        Dim cmdCommand As SqlClient.SqlCommand
        If ConnDB Is Nothing Then
            Dim clsDB As New DB
            clsDB.Open()
            ConnDB = clsDB.Connection
        End If

        Try
            With New System.Text.StringBuilder
                .AppendLine("UPDATE dbo.customer_order_headers SET coh_date_emailed = GETDATE() ")
                .AppendLine("WHERE coh_customer_order_header_id = @order_id")
                cmdCommand = New SqlClient.SqlCommand(.ToString, ConnDB)
            End With
            cmdCommand.Parameters.AddWithValue("@order_id", OrderId)
            cmdCommand.ExecuteNonQuery()
            '
            cmdCommand.Dispose()

            mdtEmailedDateTime = dtDateTime

        Catch ex As Exception
            ReportError(ex.Message, "COrderHeader", "UpdateDateEmailed")
        End Try

    End Sub
End Class

