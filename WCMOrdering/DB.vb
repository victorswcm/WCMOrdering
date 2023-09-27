Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Enum AuthenticationMode
    WindowsAuthentication
    SqlServerAuthentication
End Enum

Public Enum Protocol
    NamedPipes
    TCP_IP
End Enum

Public Class DB
    Implements IDisposable

    Protected mobjConn As SqlClient.SqlConnection

    Public Property Server As String = My.Settings.SQLServer
    Public Property Database As String = My.Settings.Database
    Public Property Port As String = 1433
    Public Property Username As String = My.Settings.DBUser
    Public Property Password As String = My.Settings.PWD
    Public Property Protocol As Protocol = Global.WCMOrdering.Protocol.NamedPipes
    Public Property AuthenticationMode As AuthenticationMode = Global.WCMOrdering.AuthenticationMode.SqlServerAuthentication


    Public Sub New()
        Database = My.Settings.Database
        If WCMOrdering._Test_Mode Then
            Database &= "_test"
        End If
    End Sub

    Public Function Test() As Boolean

        Try
            Me.Open()
            Return True
        Catch ex As Exception
            Return False
        Finally
            Me.Close()
        End Try

    End Function

    Protected Function GetConnectString() As String
        Dim connectionString As String = "Server=" & Me.Server & ";Database=" & Me.Database & ";APP=WCMOrdering"

        If Me.AuthenticationMode = AuthenticationMode.SqlServerAuthentication Then
            connectionString &= ";UID=" & Me.Username & ";PWD=" & Me.Password
        Else
            connectionString &= ";integrated security=SSPI;persist security info=False;Trusted_Connection=Yes"
        End If


        If Me.Protocol = Protocol.TCP_IP Then
            connectionString &= ";Address=" & Me.Server & "," & Me.Port & ";Network=DBMSSOCN"
        End If

        Return connectionString
    End Function

    Public Sub Open()
        'Code
        mobjConn = New SqlClient.SqlConnection(GetConnectString())
        Try
            mobjConn.Open()
        Catch ex As Exception
            Throw New Exception("Could not connect to database:" & vbCrLf & ex.Message)
        End Try
        'if the connection did not open sucessfully, redirect user to error page
        If mobjConn.State <> ConnectionState.Open Then
            'TODO: handle DB error
        End If
    End Sub

    Public Sub Close()
        'closes the database connection if it is open
        If Not mobjConn Is Nothing Then
            If mobjConn.State = ConnectionState.Open Then
                mobjConn.Close()
                mobjConn.Dispose()
            End If
        End If
    End Sub

    Public ReadOnly Property Connection() As SqlClient.SqlConnection
        Get
            Connection = mobjConn
        End Get
    End Property

    Public ReadOnly Property ConnectionString() As String
        Get
            Return GetConnectString()
        End Get
    End Property

    Public Function GetDataTable(ByVal sql As String) As DataTable
        Dim dt As New DataTable("")

        Me.Fill(sql, dt)

        Return dt

    End Function

    Public Function GetDataTable(ByVal sql As String, ByVal tableName As String, ByVal ParamArray parameters() As SqlParameter) As DataTable
        Dim dt As New DataTable(tableName)

        Me.Fill(sql, dt, parameters)
        Return dt

    End Function

    Public Function GetDataTable(ByVal cmd As SqlClient.SqlCommand, Optional ByVal tableName As String = "") As DataTable
        Dim dt As New DataTable(tableName)

        Me.Fill(cmd, dt)

        Return dt

    End Function

    Public Sub Fill(ByVal sql As String, ByRef dt As DataTable, ByVal ParamArray parameters() As SqlParameter)
        Dim cmd As SqlClient.SqlCommand
        Dim drData As SqlDataReader
        'Dim da As SqlClient.SqlDataAdapter

        cmd = Me.SqlCommand(sql, parameters)
        cmd.CommandTimeout = 6000
        'da = New SqlClient.SqlDataAdapter(cmd)
        'da.Fill(dt)
        drData = cmd.ExecuteReader
        dt = ToDataTable(drData)

    End Sub

    Public Sub Fill(ByVal cmd As SqlClient.SqlCommand, ByRef dt As DataTable)
        'Dim da As SqlClient.SqlDataAdapter
        Dim drData As SqlDataReader

        'da = New SqlClient.SqlDataAdapter(cmd)
        'da.Fill(dt)

        drData = cmd.ExecuteReader
        dt = ToDataTable(drData)

    End Sub

    ''' <summary>
    ''' A substitute for the DataAdapter, to resolve issues found when using the DataAdapter.
    ''' The DataReader is closed at the end of the conversion.
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ToDataTable(ByVal dr As SqlDataReader, Optional ByVal tableName As String = "") As DataTable
        Dim dt As New DataTable(tableName)
        Dim fieldIndex As Integer
        Dim row As DataRow
        Try
            For fieldIndex = 0 To dr.FieldCount - 1
                dt.Columns.Add(New DataColumn(dr.GetName(fieldIndex), dr.GetFieldType(fieldIndex)))
            Next
            While (dr.Read())
                row = dt.NewRow()
                For fieldIndex = 0 To dr.FieldCount - 1
                    row(fieldIndex) = dr(fieldIndex)
                Next
                dt.Rows.Add(row)
                row = Nothing
            End While
            Return dt
        Finally
            If Not dr Is Nothing Then
                dr.Close()
            End If
        End Try
    End Function

    Public Function SqlCommand(ByVal sql As String, ByVal ParamArray parameters() As SqlParameter) As SqlClient.SqlCommand
        Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand(sql, mobjConn)
        Dim parameter As SqlParameter

        For Each parameter In parameters
            cmd.Parameters.Add(parameter)
        Next

        Return cmd
    End Function

    Public Function ExecuteDataReader(ByVal sql As String, ByVal ParamArray parameters() As SqlParameter) As SqlClient.SqlDataReader
        Return Me.SqlCommand(sql, parameters).ExecuteReader()
    End Function

    Public Function ExecuteDataReader(ByVal sql As String, ByVal behaviour As System.Data.CommandBehavior, ByVal ParamArray parameters() As SqlParameter) As SqlClient.SqlDataReader
        Return Me.SqlCommand(sql, parameters).ExecuteReader(behaviour)
    End Function

    Public Function ExecuteNonQuery(ByVal sql As String, ByVal ParamArray parameters() As SqlParameter) As Integer
        Return Me.SqlCommand(sql, parameters).ExecuteNonQuery()
    End Function

    Public Function ExecuteScalar(ByVal sql As String, ByVal ParamArray parameters() As SqlParameter) As Object
        Return Me.SqlCommand(sql, parameters).ExecuteScalar()
    End Function

    Public Shared Function GetAuthenticationMode(ByVal expression As String) As AuthenticationMode
        If (expression.ToLower() = AuthenticationMode.SqlServerAuthentication.ToString().ToLower()) Then
            Return AuthenticationMode.SqlServerAuthentication
        End If

        Return AuthenticationMode.WindowsAuthentication
    End Function

    Public Shared Function GetProtocol(ByVal expression As String) As Protocol
        If (expression.ToLower() = Protocol.TCP_IP.ToString().ToLower()) Then
            Return Protocol.TCP_IP
        End If

        Return Protocol.NamedPipes
    End Function

    Public Function Nz(Of objDataType)(ByVal objVal As Object, ByVal objRet As objDataType) As objDataType
        If IsDBNull(objVal) OrElse IsNothing(objVal) Then
            Return objRet
        Else
            Return CType(objVal, objDataType)
        End If
    End Function
#Region "IDisposable Support"

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
            End If
            If Not mobjConn Is Nothing Then
                mobjConn.Dispose()
            End If
        End If
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class