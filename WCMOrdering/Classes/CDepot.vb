Public Class CDepot

    Private Const _MeName As String = "CDepot"

    Public DepotId As Long = 0
    Public Depot As String = ""
    Public DepotCode As String = ""
    Public DepotEmail As String = ""
    Public DepotAltEmail As String = ""
    Public DepotAltEmail2 As String = ""
    Public Supplier As String = ""
    Public SupplierID As Long = 0
    Public ServingCode As String = ""

    Public EDIOrdersFlag As Boolean = 0

    Public DT_CutOff As DataTable
    Public Sub New()
        CreateDataTable()
    End Sub
    Public Sub New(pDepotID As Integer)
        CreateDataTable()
        DepotId = pDepotID
        GetDepotInfo()
    End Sub

    Private Sub CreateDataTable()
        Try
            DT_CutOff = New DataTable("CutOff")
            With DT_CutOff.Columns
                .Add(New DataColumn("product_type_id", GetType(Integer)))
                .Add(New DataColumn("cutoff_day", GetType(Byte)))
                .Add(New DataColumn("cutoff_time", GetType(String)))
                .Add(New DataColumn("cutoff_datetime", GetType(String)))
            End With

        Catch ex As Exception
            ReportError(ex.Message, _MeName, System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        End Try
    End Sub

    Public Function CutOffDateTimeGet(p_productType As Integer) As String
        Try
            For Each dRow As DataRow In DT_CutOff.Rows
                If dRow("product_type_id") = p_productType Then
                    Return dRow("cutoff_datetime").ToString
                End If
            Next
        Catch ex As Exception
            ReportError(ex.Message, _MeName, System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        End Try
        Return String.Empty
    End Function
    Public Sub GetDepotInfo()
        Dim cmd As SqlClient.SqlCommand _
        , dr As SqlClient.SqlDataReader _
        , l_DB As DB

        Try

            DT_CutOff.Clear()

            l_DB = New DB
            l_DB.Open()

            With New System.Text.StringBuilder
                .AppendLine("SELECT de_depot, de_email, de_alt_email, de_alt_email2, de_code ")
                .AppendLine(", de_edi_orders")
                .AppendLine("FROM depots ")
                .Append("WHERE de_depot_id = ") : .Append(DepotId.ToString)

                cmd = New SqlClient.SqlCommand(.ToString, l_DB.Connection)
            End With
            dr = cmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    Depot = dr("de_depot").ToString
                    DepotCode = Nz(Of String)(dr("de_code"), "")
                    DepotEmail = Nz(Of String)(dr("de_email"), "")
                    DepotAltEmail = Nz(Of String)(dr("de_alt_email"), "")
                    DepotAltEmail2 = Nz(Of String)(dr("de_alt_email2"), "")
                    EDIOrdersFlag = Nz(Of Boolean)(dr("de_edi_orders"), 0)
                End If
                dr.Close()

                Call GetDepotCutOff(l_DB.Connection)
            End If

        Catch ex As Exception
            ReportError(ex.Message, _MeName, System.Reflection.MethodInfo.GetCurrentMethod.ToString)

        Finally
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then
                dr.Close()
            End If
        End Try
    End Sub
    Public Sub GetSupplier(pCustomerID As Long, pDeliveryDate As Date)
        Dim lCmd As SqlClient.SqlCommand,
            lParam As SqlClient.SqlParameter,
            lResult As Object = Nothing,
            lDB As DB,
            Sql As String = ""

        Try

            lDB = New DB
            lDB.Open()
            lCmd = New SqlClient.SqlCommand("p_customer_ordering_info_get", lDB.Connection)
            With lCmd
                .CommandType = CommandType.StoredProcedure

                'Create Parameters
                lParam = .Parameters.Add("@return", SqlDbType.Int)
                lParam.Direction = ParameterDirection.ReturnValue

                lParam = .Parameters.Add("@cust_id", SqlDbType.BigInt)
                lParam.Value = pCustomerID

                lParam = .Parameters.Add("@delivery_date", SqlDbType.Date)
                lParam.Value = pDeliveryDate

                lParam = .Parameters.Add("@depot_id", SqlDbType.Int)
                lParam.Direction = ParameterDirection.InputOutput
                lParam.Value = DepotId

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

                ''If pIsProduce Then
                ''    lParam = .Parameters.Add("@supply_type", SqlDbType.Int)
                ''Else
                ''    'defaults to 1
                ''End If
                lParam = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                lParam.Direction = ParameterDirection.Output

                .ExecuteNonQuery()

                'note it returns customer_price_info_id
                Select Case .Parameters("@return").Value
                    Case 0
                        '
                    Case Is > 0
                        ServingCode = Nz(Of String)(.Parameters("@serv_code").Value, "")
                        SupplierID = Nz(Of Long)(.Parameters("@supplier_id").Value, 0)
                        Supplier = Nz(Of String)(.Parameters("@supplier").Value, "")
                    Case Else ' other errors
                        MsgBox(.Parameters("@err_msg").Value, MsgBoxStyle.Information, Depot)

                End Select
            End With


        Catch ex As Exception
            ReportError(ex.Message, _MeName, System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        End Try
    End Sub

    Public Sub CutOffDateTimeSet(pDeliveryDate As Date)
        Dim dtCutOff As Date

        Try

            For Each dRow As DataRow In DT_CutOff.Rows
                If dRow("cutoff_day") = 0 Then
                    dRow("cutoff_datetime") = String.Empty
                Else
                    dtCutOff = DateAdd(DateInterval.Day, -dRow("cutoff_day"), pDeliveryDate)
                    If DatePart(DateInterval.Weekday, dtCutOff, FirstDayOfWeek.Sunday) = 1 Then
                        dtCutOff = DateAdd(DateInterval.Day, -1, dtCutOff)
                    End If
                    dRow("cutoff_datetime") = CType(sDate(dtCutOff) + " " + dRow("cutoff_time"), Date).ToString("ddd dd/MM HH:mm")
                End If
            Next

        Catch ex As Exception
            ReportError(ex.Message, _MeName, System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        End Try
    End Sub

    Private Function GetDepotCutOff(ByVal Conn As SqlClient.SqlConnection) As Boolean
        Dim cmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader = Nothing
        Dim dtRow As DataRow

        Try

            cmd = New SqlClient.SqlCommand("SELECT * FROM depot_cut_off WHERE dco_depot_id = " & DepotId.ToString, Conn)

            dr = cmd.ExecuteReader

            If dr.HasRows Then
                Do While dr.Read
                    dtRow = DT_CutOff.NewRow()
                    dtRow("product_type_id") = dr("dco_product_type_id")
                    dtRow("cutoff_day") = dr("dco_cutoff_day")
                    If IsDBNull(dr("dco_cutoff_time")) Then
                        dtRow("cutoff_time") = "17:00"
                    Else
                        dtRow("cutoff_time") = dr("dco_cutoff_time").ToString
                    End If
                    DT_CutOff.Rows.Add(dtRow)
                Loop
            End If
            Return True
        Catch ex As Exception
            ReportError(ex.Message, _MeName, System.Reflection.MethodInfo.GetCurrentMethod.ToString)

        Finally
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then
                dr.Close()
            End If
        End Try
        Return False
    End Function
End Class
