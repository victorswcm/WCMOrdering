Imports System.Runtime.InteropServices
Imports System.Xml
Imports System.IO
Imports System.Net.Mail
Imports System.Data
Imports System.Globalization
'Imports System.Net
Imports System.Text

Public Class WCMOrdering
    Private Const _MeName As String = "WCMOrdering"

#Region "Declarations --------------------------------------------------------------"
    Private Const c_MYSOURCE As String = "WCMOrderingMySource"
    Private Const c_MYLOG As String = "WCMOrderingLog"

    Private Property _workStartTime As Date = DateTime.MinValue

    Private Property _eventId_Common As Integer = 0

    ' Private Property _eventId_Elior As Integer = 1001           ' 7  is not live 
    Private Property _eventId_Cypad As Integer = 15001          ' 9
    Private Property _eventId_Bourne As Integer = 20001         '11
    Private Property _eventId_Medina As Integer = 25001         '10
    Private Property _eventId_Interserve As Integer = 30001     '12
    Private Property _eventId_Compass As Integer = 35001        ' 5
    Private Property _eventId_FoodBuy_Online As Integer = 40001 ' 8
    Private Property _eventId_Email_Orders As Integer = 45001   '99
    Private Property _eventId_DN_Grahams As Integer = 50001     '98
    Private Property _eventId_Poundland As Integer = 55001      '13
    Private Property _eventId_CN_CSV As Integer = 60001         ' 6
    Private Property _eventId_McColls As Integer = 10001         '18 
    Private Property _eventId_Zupa As Integer = 1001           '16 --Zupa, re-use Elior range
    Private Property _eventId_Weezy As Integer = 65001        '19
    Private Property _eventId_Order_Upload As Integer = 101     'not required as this is not the order upload

    Public Shared _Test_Mode As Boolean = False
    Private Property _OrderRequestId As Integer
    Private Property _AccNum As String
    Private Property _VendorID As String
    Private Property _OrderNum As String
    Private Property _DeliveryNoteNum As String

    Private Property _DeliveryDate As Date
    Private Property _Status As String
    Private Property _Order_DateTime_Created As String

    Private Property _RecentResponseFileName As String = String.Empty ' ELIOR

    'ELIOR ----------
    Private __LastOrderReceivedDatetime_Elior As Date
    Private __LastWarningEmailDatetime_Elior As Date

    'COMPASS--------- and POUNDLAND
    Private Property _DeliveryDateRequested As Date = Date.MinValue
    Private Property _Order_Lines As Integer
    Private Property _Order_Value As Decimal


    'FoodBuy_Online ----------
    Private __LastOrderReceivedDatetime_FoodBuy_Online As Date
    Private __LastWarningEmailDatetime_FoodBuy_Online As Date

    'CYPAD------------
    Private Property _OrderGuid As String
    Private __LastOrderReceivedDatetime_Cypad As Date
    Private __LastWarningEmailDatetime_Cypad As Date
    '-----------------

    'BOURNE------------
    Private __LastOrderReceivedDatetime_Bourne As Date
    Private __LastWarningEmailDatetime_Bourne As Date
    '-----------------

    'MEDINA------------
    Private __LastOrderReceivedDatetime_Medina As Date
    Private __LastWarningEmailDatetime_Medina As Date
    '-----------------

    'INTERSERVE------------
    Private __LastOrderReceivedDatetime_Interserve As Date
    Private __LastWarningEmailDatetime_Interserve As Date
    '-----------------
    'POUNDLAND------------
    Private __LastOrderReceivedDatetime_Poundland As Date
    Private __LastWarningEmailDatetime_Poundland As Date
    'CAFE NERO CSV------------
    Private __LastOrderReceivedDatetime_CN_CSV As Date
    Private __LastWarningEmailDatetime_CN_CSV As Date
    '-----------------
    'ZUPA ----------
    Private __LastOrderReceivedDatetime_Zupa As Date
    Private __LastWarningEmailDatetime_Zupa As Date

    'MCCOLLS ----------
    Private __LastOrderReceivedDatetime_McColls As Date
    Private __LastWarningEmailDatetime_McColls As Date

    Private _DairyData_JN_done As Date = "2020-01-01"
    Private _AR_done As Date = "2020-01-01"
    Private _DairyData_FP_done As Date = "2020-01-01"
    Private _DairyData_MM_done As Date = "2020-01-01"

    Private _DairyData_DHT_done As Date = "2020-01-01"
    Private _DairyData_Paynes_done As Date = "2020-01-01"
    Private _Johal_done As Date = "2020-01-01"
    Private _Grahams_done As Date = "2020-01-01"
    Private _DairyData_Broadland_done As Date = "2020-01-01"
    Private _Chew_Valley_done As Date = "2020-01-01"
    Private _DairyData_Medina_done As Date = "2020-01-01"
    Private _Order_Alert_OfficeDrop_done As Date = "2020-01-01"
    Private _JJWilson_done As Date = "2020-01-01"
    Private _BR003_reminder_done As Date = "2020-01-01"
    Private _Saffron_ASN_done As Date = "2020-01-01"


    '-----------------

    'FOLDERS for orders and responses/acknowlegements and also deliveries ( grahams for now)
    Private _ORDER_IN As String = String.Empty
    Private _ORDER_ARCHIVED As String = String.Empty
    Private _ORDER_FAILED As String = String.Empty
    Private _RESPONSE_OUT As String = String.Empty
    Private _RESPONSE_ARCHIVED As String = String.Empty
    Private _ORDER_OUT As String = String.Empty
    Private _SUPPLIER_ID As Long = 0

    Private WithEvents mOrder As COrderHeader

    Private _Orders_ToEmail As Stack(Of Integer)
    Private _SO_ToEmail As Stack(Of Integer)

    Public Enum ServiceState
        SERVICE_STOPPED = 1
        SERVICE_START_PENDING = 2
        SERVICE_STOP_PENDING = 3
        SERVICE_RUNNING = 4
        SERVICE_CONTINUE_PENDING = 5
        SERVICE_PAUSE_PENDING = 6
        SERVICE_PAUSED = 7
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure ServiceStatus
        Public dwServiceType As Long
        Public dwCurrentState As ServiceState
        Public dwControlsAccepted As Long
        Public dwWin32ExitCode As Long
        Public dwServiceSpecificExitCode As Long
        Public dwCheckPoint As Long
        Public dwWaitHint As Long
    End Structure

    Declare Auto Function SetServiceStatus Lib "advapi32.dll" (ByVal handle As IntPtr, ByRef serviceStatus As ServiceStatus) As Boolean

#End Region

#Region "Constructors --------------------------------------------------------------"
    Public Sub New()
        MyBase.New()
        ' This call is required by the designer.
        InitializeComponent()

        MyEventLog = New EventLog
        If Not EventLog.SourceExists(c_MYSOURCE) Then
            EventLog.CreateEventSource(c_MYSOURCE, c_MYLOG)
        End If
        MyEventLog.Source = c_MYSOURCE
        MyEventLog.Log = c_MYLOG

        __LastOrderReceivedDatetime_Elior = Date.Now
        __LastOrderReceivedDatetime_FoodBuy_Online = Date.Now
        __LastOrderReceivedDatetime_Cypad = Date.Now
        __LastOrderReceivedDatetime_Bourne = Date.Now
        __LastOrderReceivedDatetime_Medina = Date.Now
        __LastOrderReceivedDatetime_Interserve = Date.Now
        __LastOrderReceivedDatetime_Poundland = Date.Now
        __LastOrderReceivedDatetime_Zupa = Date.Now
        __LastOrderReceivedDatetime_McColls = Date.Now
    End Sub

#End Region
#Region "On-Actions --------------------------------------------------------------"
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things in motion so your service can do its work.
        Try
            MyEventLog.WriteEntry("WCMOrdering Started", EventLogEntryType.Information, GetEventID(0))

            ' Update the service state to Start Pending.
            Dim serviceStatus As ServiceStatus = New ServiceStatus()
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING
            serviceStatus.dwWaitHint = 100000
            SetServiceStatus(Me.ServiceHandle, serviceStatus)

            ' Set up a timer to trigger every minute.
            Dim timer As System.Timers.Timer = New System.Timers.Timer()
            If My.Settings.TimerInterval > 10000 Then
                timer.Interval = My.Settings.TimerInterval ' 120 seconds
            Else
                timer.Interval = 120000 ' default 120 seconds
            End If
            AddHandler timer.Elapsed, AddressOf Me.OnTimer

            __LastOrderReceivedDatetime_Elior = Date.Now
            __LastOrderReceivedDatetime_FoodBuy_Online = Date.Now
            __LastOrderReceivedDatetime_Cypad = Date.Now
            __LastOrderReceivedDatetime_Bourne = Date.Now
            __LastOrderReceivedDatetime_Medina = Date.Now
            __LastOrderReceivedDatetime_Interserve = Date.Now
            __LastOrderReceivedDatetime_Poundland = Date.Now
            __LastOrderReceivedDatetime_Zupa = Date.Now
            __LastOrderReceivedDatetime_McColls = Date.Now

            timer.Start()

            'Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING
            SetServiceStatus(Me.ServiceHandle, serviceStatus)

            _EmailServiceMessage("WCMOrdering service started " & Date.Now)

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID(0))
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
    End Sub

    Protected Overrides Sub OnContinue()
        Try
            ' MyEventLog.WriteEntry("In OnContinue.")

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID(0))
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Try
            MyEventLog.WriteEntry("WCMOrdering Service Stopped", EventLogEntryType.Information, GetEventID(0))

            _EmailServiceMessage(Date.Now & " " & "WCMOrdering Service Stopped")
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID(0))
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
    End Sub

    ' Insert monitoring activities here.
    Private Sub OnTimer(sender As Object, e As Timers.ElapsedEventArgs)
        Dim l_MonitoringInterval_Minutes As Integer = My.Settings.WarningInterval_19to5 '6 hours 
        Dim l_Warning As String = String.Empty
        Dim l_Interval_Elior As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_Elior, Date.Now)
        Dim l_Interval_FoodBuy_Online As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_FoodBuy_Online, Date.Now)
        Dim l_Interval_Cypad As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_Cypad, Date.Now)
        Dim l_Interval_Bourne As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_Bourne, Date.Now)
        Dim l_Interval_Medina As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_Medina, Date.Now)
        Dim l_Interval_Interserve As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_Interserve, Date.Now)
        Dim l_Interval_Poundland As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_Poundland, Date.Now)
        Dim l_Interval_Zupa As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_Zupa, Date.Now)
        Dim l_Interval_McColls As Integer = DateDiff(DateInterval.Minute, __LastOrderReceivedDatetime_McColls, Date.Now)

        Try
            ' Insert monitoring activities here.
            Select Case Hour(Date.Now)
                Case 6 To 10
                    l_MonitoringInterval_Minutes = My.Settings.WarningInterval_6to10
                Case 11 To 18
                    l_MonitoringInterval_Minutes = My.Settings.WarningInterval_11to18
                Case Else
                    If Weekday(Date.Now) = 6 Then
                        l_MonitoringInterval_Minutes = My.Settings.WarningInterval_weekend
                    End If
            End Select
            If Weekday(Date.Now) = 1 Then
                'sunday
                l_MonitoringInterval_Minutes = My.Settings.WarningInterval_weekend
            End If

            'If My.Settings.Switch_Elior = 2 AndAlso l_Interval_Elior >= l_MonitoringInterval_Minutes _
            '            AndAlso DateDiff(DateInterval.Minute, __LastWarningEmailDatetime_Elior, Date.Now) >= l_MonitoringInterval_Minutes Then
            '    l_Warning = vbNewLine & vbNewLine & "WARNING!  NO ELIOR ORDER REQUESTS DETECTED FOR " & l_Interval_Elior & " minutes"
            '    __LastWarningEmailDatetime_Elior = Date.Now
            'End If


            If Not String.IsNullOrEmpty(l_Warning) Then
                MyEventLog.WriteEntry(l_Warning, EventLogEntryType.Warning, GetEventID(0))
                _EmailServiceMessage(Date.Now & "WARNING!  " & l_Warning)
            End If

            If _workStartTime <> DateTime.MinValue Then
                MyEventLog.WriteEntry("WARNING! Worker process busy since " & _workStartTime.ToLongTimeString, EventLogEntryType.Warning, GetEventID(0))
            Else
                _workStartTime = Date.Now


                ''If GetSetting_PushEmail("WEBAPP") Then
                ''    If GetOrdersNotEmailed() Then ' webapp not emailed
                ''        PushEmailOrders()
                ''    End If
                ''    If GetStandingNotEmailed() Then ' SO not emailed
                ''        PushEmailStandingOrders()
                ''    End If
                ''End If
                ''If GetSetting_PushEmail("P2P") Then
                ''    If GetOrdersNotEmailed() Then ' P2P not emailed
                ''        PushEmailOrders()
                ''    End If
                ''End If

                ''If GetSetting_Foodbuy_Online() Then
                ''    ' former compass
                ''    Process_FoodBuy_Online()
                ''End If

                ''If GetSetting_Bourne() Then
                ''    Process_Bourne()
                ''End If

                If GetSetting_Interserve_saffron() Then
                    Process_Interserve()

                    'process Saffron ASN for Standing Orders (sub-buying group Debra)
                    'If DateDiff(DateInterval.Day, _Saffron_ASN_done, Now.Date) > 0 Then
                    '    If Date.Now.Hour = 19 AndAlso (Date.Now.Minute > 30 AndAlso Date.Now.Minute < 35) Then
                    For idx As Integer = 0 To 14
                        Process_Saffron_ASN(DateAdd(DateInterval.Day, idx, CDate("11 Mar 2024")))
                        System.Threading.Thread.Sleep(100)
                        '_Saffron_ASN_done = Now.Date
                        ''    End If
                        ''End
                    Next
                End If

                'If GetSetting_DN_Grahams() Then
                '    Process_DN_Grahams()
                'End If


                'If GetSetting_CN_CrunchTime() Then
                '    Process_CN_CrunchTime()
                'End If


                'Process_DairyData_MillsMilk()

                'Process_AllanReeder()


                'Process_DairyData_Paynes()

                'Process_DairyData_Broadland()

                'Order_Alert_OfficeDrop()

                'Order_Alert_BR003()

                'Process_Johal()

                'Process_Grahams()

                'Process_JJWison()

                _workStartTime = Date.MinValue
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID(0))
        End Try
    End Sub

#End Region

#Region "Order_Alert ---------------------------------------------------------------------"
    Private Sub Order_Alert_OfficeDrop()
        Dim lMessage As String
        Dim lSubject As String
        Dim lTo As String = "Orders@wcmilk.co.uk"
        Dim lCC As String = ""

        If DateDiff(DateInterval.Day, _Order_Alert_OfficeDrop_done, Now.Date) = 0 Then Return

        lSubject = "Office Drop Order Check"
        lMessage = "Please check if the following site has placed an order" & vbNewLine & vbNewLine & "Site: MOFD0005 " & vbNewLine & vbNewLine & "If no order placed please copy from the week prior"
        If Date.Now.DayOfWeek = (DayOfWeek.Monday) OrElse Date.Now.DayOfWeek = (DayOfWeek.Friday) Then
            If Date.Now.Hour = 13 AndAlso (Date.Now.Minute > 55 And Date.Now.Minute < 59) Then
                _EmailServiceMessage(lSubject, lMessage, lTo, lCC)
                _Order_Alert_OfficeDrop_done = Now.Date
            End If
        End If
    End Sub
    Private Sub Order_Alert_BR003()
        Dim lMessage As String
        Dim lSubject As String
        Dim lTo As String = "Orders@wcmilk.co.uk"
        Dim lCC As String = ""

        If DateDiff(DateInterval.Day, _Order_Alert_OfficeDrop_done, Now.Date) = 0 Then Return

        lSubject = "REMINDER: check Order for BR003"
        lMessage = "Please check if the following site has placed an order" & vbNewLine & vbNewLine & "Site: BR003 - Brakes " & vbNewLine & vbNewLine & "If no order placed please copy from the prior week "
        If DateDiff(DateInterval.Day, _BR003_reminder_done, Now.Date) > 0 Then
            If Date.Now.Hour = 14 AndAlso (Date.Now.Minute > 0 And Date.Now.Minute < 5) Then
                _EmailServiceMessage(lSubject, lMessage, lTo, lCC)
                _BR003_reminder_done = Now.Date
            End If
        End If
    End Sub
#End Region

#Region "Methods ELIOR--------------------------------------------------------------"
    Private Sub Process_Elior()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer

        Dim bFailedToSendResponces As Boolean = False
        Dim lMsg As String = String.Empty
        Dim lResult As MsgBoxResult
        Try

            ''MyEventLog.WriteEntry("Monitoring the System", EventLogEntryType.Information,GetEventID(0)) 

            ''CHECK if recent responses have been sent away
            If _RecentResponseFileName.Length > 0 Then
                asFiles = Directory.GetFiles(_RESPONSE_OUT)
                For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                    l_File = Path.GetFileName(asFiles(idx))
                    If bFailedToSendResponces = False Then
                        If l_File.Equals(_RecentResponseFileName, comparisonType:=StringComparison.InvariantCultureIgnoreCase) Then
                            bFailedToSendResponces = True
                            lMsg = "AS2 ERROR  " & Date.Now & ".  FAILED to send file(s):" & vbCrLf
                        End If
                    End If
                    If bFailedToSendResponces Then
                        lMsg &= l_File & vbCrLf
                    End If
                Next
                If bFailedToSendResponces Then
                    MyEventLog.WriteEntry(lMsg, EventLogEntryType.Error, GetEventID)

                    _EmailServiceMessage(lMsg, True)
                End If

                _RecentResponseFileName = String.Empty
            End If

            'PROCESS NEW ORDERS
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))
                'If l_File.StartsWith("POR") AndAlso Not l_File.EndsWith("_response.xml") Then
                MyEventLog.WriteEntry("ORDER REQUEST RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)

                If Not _Test_Mode Then
                    __LastOrderReceivedDatetime_Elior = Date.Now
                    __LastWarningEmailDatetime_Elior = __LastOrderReceivedDatetime_Elior
                End If

                lResult = UploadFile_Elior(_ORDER_IN & "\" & l_File)
                If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                    Return
                End If
                'ElseIf Not l_File.EndsWith("_response.xml") Then
                'MyEventLog.WriteEntry("INVALID FILE RECEIVED:  " & l_File)
                'MoveFile(_ORDER_IN & "\" & l_File, _ORDER_IN & "\Failed_Orders\" & l_File)
                '_EmailServiceMessage(Date.Now & " " & "INVALID FILE RECEIVED:  " & l_File, True)
                'End If
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
    End Sub

    Private Function UploadFile_Elior(ByVal pFileName As String) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand
        Dim lXMLContents As String = String.Empty
        Dim lReader As StreamReader
        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim dt As DataTable = Nothing
        Dim lNewOrderID As Integer = 0

        Try
            l_DB.Open()

            lReader = New StreamReader(pFileName)
            lXMLContents = lReader.ReadToEnd()
            lReader.Close()

            lXMLContents = lXMLContents.Replace("<sh:", "<").Replace("<eanucc:", "<").Replace("<order:", "<")
            lXMLContents = lXMLContents.Replace("</sh:", "</").Replace("</eanucc:", "</").Replace("</order:", "</")
            lXMLContents = lXMLContents.Replace("encoding=""utf-8""", "")
            cmd = l_DB.SqlCommand("p_P2P_order_import")

            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 200)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@xml_order", SqlDbType.Xml)
                param.Direction = ParameterDirection.Input
                param.Value = lXMLContents

                param = .Parameters.Add("@order_num", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@vendor_id", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@datetime_created_string", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@status", SqlDbType.VarChar, 16)
                param.Direction = ParameterDirection.InputOutput
                param.Value = "ACCEPTED"

                param = .Parameters.Add("@customer_order_header_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                .ExecuteNonQuery()
                lRetVal = CType(.Parameters("@ret").Value, Integer)

                _OrderRequestId = .Parameters("@record_id").Value
                _OrderNum = Nz(Of String)(.Parameters("@order_num").Value, "")
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "")
                _VendorID = Nz(Of String)(.Parameters("@vendor_id").Value, "")
                _DeliveryDate = Nz(Of Date)(.Parameters("@delivery_date").Value, Date.MinValue)
                _Status = Nz(Of String)(.Parameters("@status").Value, "")
                _Order_DateTime_Created = Nz(Of String)(.Parameters("@datetime_created_string").Value, "")

                lNewOrderID = Nz(Of Integer)(.Parameters("@customer_order_header_id").Value, 0)

                lErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")

                Select Case lRetVal
                    Case 0 ' Success
                        If _Status = "MODIFIED_LINES" Then
                            dt = GetOrderAmendments()
                            _Status = "MODIFIED"
                        End If
                        If CreateOrderResponse("", dt) Then
                            MyEventLog.WriteEntry("Order Response Created: " & _OrderNum & " {Request id = " & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                            If _Test_Mode Then
                                _EmailServiceMessage(Date.Now & " ORDER RESPONSE CREATED: " & _OrderNum & " {Request id = " & _OrderRequestId & "} for " & _AccNum)
                            End If
                            System.Threading.Thread.Sleep(50)

                            If EmailOrder(lNewOrderID) Then
                                Return MsgBoxResult.Ok
                            End If
                        End If

                        Return MsgBoxResult.Ok
                        'Case 1 To 333
                        '    _Status = "REJECTED"
                        '    Return CreateOrderResponse(lErrMsg)
                    Case Else
                        _Status = "REJECTED"
                        Dim lRejectCode As String = lErrMsg
                        If lErrMsg.Contains("~") Then
                            lRejectCode = lErrMsg.Substring(0, lErrMsg.IndexOf("~"))
                        End If
                        If CreateOrderResponse(lRejectCode) Then
                            If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "

                            MyEventLog.WriteEntry("Order Response Created: REJECTED " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg, EventLogEntryType.Warning, GetEventID)

                            _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg)


                            System.Threading.Thread.Sleep(50)
                        End If
                End Select
            End With
            cmd.Dispose()
            l_DB.Close()

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function

    Private Function GetOrderAmendments() As DataTable
        Return GetOrderAmendments(0)
    End Function

    Private Function GetOrderAmendments(p_All As Boolean) As DataTable
        Dim l_DB = New DB
        Dim cmd As SqlClient.SqlCommand = Nothing
        Dim dt As DataTable = New DataTable("modified_lines")
        Try
            l_DB.Open()
            cmd = l_DB.SqlCommand("p_P2P_order_amendments")
            With cmd
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.Add("@record_id", Data.SqlDbType.Int)
                .Parameters("@record_id").Value = _OrderRequestId
                If p_All Then
                    .Parameters.Add("@all", Data.SqlDbType.Bit)
                    .Parameters("@all").Value = 1
                End If
            End With
            l_DB.Fill(cmd, dt)
            cmd.Dispose()
            l_DB.Close()
        Catch ex As Exception
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
        Return dt
    End Function

    Private Function CreateOrderResponse() As Boolean
        Return CreateOrderResponse(String.Empty, Nothing)
    End Function

    Private Function CreateOrderResponse(p_RejectReason As String) As Boolean
        Return CreateOrderResponse(p_RejectReason, Nothing)
    End Function

    Private Function CreateOrderResponse(p_RejectReason As String, ByVal dtLines As DataTable) As Boolean
        Dim xmlWriter As XmlTextWriter = Nothing
        Dim l_File As String = _OrderNum & "_response.xml"
        'Dim p_Status As String = "ACCEPTED" ' "MODIFIED", "REJECTED"
        ''Dim l_RejectReason As String = "CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST"
        ''Dim l_GLN_WCM As String = My.Settings.GLN_WCM
        Dim row As DataRow 'invoice detail datarow
        ''Dim l_OrderNum As String = "POR123456"
        ''Dim l_Order_Datetime_created_string As String = '"2016-01-21T09:00:00+01:00"
        ''Dim l_AccNum As String = "Q0S676"
        ''Dim l_VendorID As String = "GBR01234BB"
        ''Dim l_DeliveryDate As Date = "10 Mar 2016"

        ''Dim l_LineSeq As Integer = 1
        ''Dim l_Qty As Integer = 10
        ''Dim l_ProductCode_Original As String = "10214"
        ''Dim l_ProductCode_Substitude As String = "10282"

        Try


            If File.Exists(_RESPONSE_OUT & "\" & l_File) Then
                File.Delete(_RESPONSE_OUT & "\" & l_File)
            End If

            xmlWriter = New XmlTextWriter(_RESPONSE_OUT & "\" & l_File, System.Text.Encoding.UTF8)

            With xmlWriter
                .WriteStartDocument(True)
                .WriteStartElement("sh:StandardBusinessDocument")
                AddAttribute(xmlWriter, "xmlns:sh", "http://www.unece.org/cefact/namespaces/StandardBusinessDocumentHeader")
                AddAttribute(xmlWriter, "xmlns:order", "urn:ean.ucc:order:2")
                AddAttribute(xmlWriter, "xmlns:eanucc", "urn:ean.ucc:2")
                AddAttribute(xmlWriter, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
                AddAttribute(xmlWriter, "xsi:schemaLocation", "http://www.unece.org/cefact/namespaces/StandardBusinessDocumentHeader ../Schemas/sbdh/StandardBusinessDocumentHeader.xsd urn:ean.ucc:2 ../Schemas/OrderResponseProxy.xsd")

                .WriteStartElement("sh:StandardBusinessDocumentHeader")

                .WriteStartElement("sh:HeaderVersion")
                .WriteString("2.3")
                .WriteEndElement()

                .WriteStartElement("sh:Sender")
                .WriteStartElement("sh:Identifier") : AddAttribute(xmlWriter, "Authority", "EAN.UCC")
                .WriteString("Westcountry Milk")
                .WriteEndElement()
                .WriteEndElement()

                .WriteStartElement("sh:Receiver")
                .WriteStartElement("sh:Identifier") : AddAttribute(xmlWriter, "Authority", "EAN.UCC")
                .WriteString("Elior UK Limited")
                .WriteEndElement()
                .WriteEndElement()


                .WriteStartElement("sh:DocumentIdentification")
                .WriteStartElement("sh:Standard") : .WriteString("EAN.UCC") : .WriteEndElement()
                .WriteStartElement("sh:TypeVersion") : .WriteString("2.3") : .WriteEndElement()
                .WriteStartElement("sh:InstanceIdentifier") : .WriteString("1111") : .WriteEndElement()
                .WriteStartElement("sh:Type") : .WriteString("OrderResponse") : .WriteEndElement()
                .WriteStartElement("sh:CreationDateAndTime") : .WriteString(Now.Date.ToString("yyyy-MM-dd") & "T" & Now.ToString("HH:mm:ss") & "+00:00") : .WriteEndElement()
                .WriteEndElement()

                .WriteEndElement()  'sh:StandardBusinessDocumentHeader

                .WriteStartElement("eanucc:message")
                .WriteStartElement("entityIdentification")
                .WriteStartElement("uniqueCreatorIdentification") : .WriteString("2222") : .WriteEndElement()
                .WriteStartElement("contentOwner")
                .WriteStartElement("gln") : .WriteString(My.Settings.GLN_WCM) : .WriteEndElement() '
                .WriteEndElement()
                .WriteEndElement() 'entityIdentification

                .WriteStartElement("order:orderResponse")
                AddAttribute(xmlWriter, "creationDateTime", Now.Date.ToString("yyyy-MM-dd") & "T" & Now.ToString("HH:mm:ss") & "+00:00")
                AddAttribute(xmlWriter, "documentStatus", "ORIGINAL")
                AddAttribute(xmlWriter, "responseStatusType", _Status)

                .WriteStartElement("responseIdentification")
                .WriteStartElement("uniqueCreatorIdentification") : .WriteString(_OrderNum) : .WriteEndElement() 'supplier sales order number 
                .WriteStartElement("contentOwner")
                .WriteStartElement("gln") : .WriteString(My.Settings.GLN_WCM) : .WriteEndElement()
                .WriteEndElement()
                .WriteEndElement() ' responseIdentification

                .WriteStartElement("responseToOriginalDocument")
                AddAttribute(xmlWriter, "referenceDateTime", _Order_DateTime_Created)
                AddAttribute(xmlWriter, "referenceIdentification", _OrderNum)
                AddAttribute(xmlWriter, "referenceDocumentType", "35")
                .WriteEndElement()

                .WriteStartElement("buyer")
                .WriteStartElement("gln") : .WriteString("5055902800000") : .WriteEndElement()
                .WriteStartElement("additionalPartyIdentification")
                .WriteStartElement("additionalPartyIdentificationValue") : .WriteString(_AccNum) : .WriteEndElement() 'Elior customer account number
                .WriteStartElement("additionalPartyIdentificationType") : .WriteString("BUYER_ASSIGNED_IDENTIFIER_FOR_A_PARTY") : .WriteEndElement()
                .WriteEndElement()
                .WriteStartElement("additionalPartyIdentification")
                .WriteStartElement("additionalPartyIdentificationValue") : .WriteString(_AccNum) : .WriteEndElement() 'mapped wcm customer account number (it is the same as buyer for Elior)
                .WriteStartElement("additionalPartyIdentificationType") : .WriteString("SELLER_ASSIGNED_IDENTIFIER_FOR_A_PARTY") : .WriteEndElement()
                .WriteEndElement()
                .WriteEndElement() ' buyer

                .WriteStartElement("seller")
                .WriteStartElement("gln") : .WriteString(My.Settings.GLN_WCM) : .WriteEndElement()
                .WriteStartElement("additionalPartyIdentification")
                .WriteStartElement("additionalPartyIdentificationValue") : .WriteString(_VendorID) : .WriteEndElement() 'Elior customer account number
                .WriteStartElement("additionalPartyIdentificationType") : .WriteString("SELLER_ASSIGNED_IDENTIFIER_FOR_A_PARTY") : .WriteEndElement()
                .WriteEndElement()
                .WriteEndElement() ' seller

                If String.Equals(_Status, "REJECTED", StringComparison.CurrentCultureIgnoreCase) Then
                    .WriteStartElement("orderResponseReasonCode")
                    .WriteString(p_RejectReason)
                    .WriteEndElement()

                ElseIf String.Equals(_Status, "MODIFIED", StringComparison.CurrentCultureIgnoreCase) Then
                    .WriteStartElement("orderModification")

                    .WriteStartElement("amendedDateTimeValue")
                    .WriteStartElement("requestedDeliveryDate")
                    .WriteStartElement("date") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd")) : .WriteEndElement()
                    .WriteStartElement("time") : .WriteString("09:00:00") : .WriteEndElement()
                    .WriteEndElement()
                    .WriteEndElement() 'amendedDateTimeValue

                    If dtLines IsNot Nothing Then
                        For Each row In dtLines.Rows
                            .WriteStartElement("orderModificationLineItemLevel")
                            If row("Product_Code_Requested") <> row("Product_Code") AndAlso row("Qty") <> 0 Then ' new code
                                'IF substituted Then insert this block -------------------------------
                                ''If row("Product_Code_Requested") <> row("Product_Code") Then

                                .WriteStartElement("substituteItemIdentification")
                                .WriteStartElement("gtin") : .WriteString("00000000000000") : .WriteEndElement()

                                .WriteStartElement("additionalTradeItemIdentification")
                                .WriteStartElement("additionalTradeItemIdentificationValue") : .WriteString(row("Product_Code_Requested")) : .WriteEndElement()
                                .WriteStartElement("additionalTradeItemIdentificationType") : .WriteString("SUPPLIER_ASSIGNED") : .WriteEndElement()
                                .WriteEndElement() 'additionalTradeItemIdentification>

                                .WriteEndElement() 'substituteItemIdentification
                            End If ' end of substituted item block --------------------------------------------------------------------


                            .WriteStartElement("modifiedOrderInformation") : AddAttribute(xmlWriter, "number", row("Line_Seq").ToString)
                            .WriteStartElement("requestedQuantity")
                            .WriteStartElement("value") : .WriteString(row("Qty").ToString) : .WriteEndElement() ' NOTE Qty 0 = can not supply this item and no substitude available
                            .WriteStartElement("unitOfMeasure")
                            .WriteStartElement("measurementUnitCodeValue") : .WriteString("EA") : .WriteEndElement()
                            .WriteEndElement() 'unitOfMeasure
                            .WriteEndElement() 'requestedQuantity


                            .WriteStartElement("tradeItemIdentification")
                            .WriteStartElement("gtin") : .WriteString("00000000000000") : .WriteEndElement() ' The supplier substituted item has a GTIN of 05024333127970
                            .WriteStartElement("additionalTradeItemIdentification")
                            If row("Qty") <> 0 AndAlso Not String.IsNullOrEmpty(row("Product_Code")) Then
                                .WriteStartElement("additionalTradeItemIdentificationValue") : .WriteString(row("Product_Code")) : .WriteEndElement() ''ProductCode_Substitude
                            Else
                                .WriteStartElement("additionalTradeItemIdentificationValue") : .WriteString(row("Product_Code_Requested")) : .WriteEndElement() '' Original product code
                            End If
                            .WriteStartElement("additionalTradeItemIdentificationType") : .WriteString("SUPPLIER_ASSIGNED") : .WriteEndElement()
                            .WriteEndElement() 'additionalTradeItemIdentification
                            .WriteEndElement() 'tradeItemIdentification


                            .WriteEndElement() 'modifiedOrderInformation
                            .WriteStartElement("orderResponseReasonCode") : .WriteString("PRODUCT_OUT_OF_STOCK") : .WriteEndElement()

                            .WriteEndElement() 'orderModificationLineItemLevel
                        Next row 'end for each item
                    End If
                    .WriteEndElement() 'orderModification
                End If

                .WriteEndElement() ' order:orderResponse

                .WriteEndElement() ' eanucc:message

                .WriteEndElement()  'end of Root element: sh:StandardBusinessDocument
                .WriteEndDocument()
                .Flush()
                .Close()
            End With

            _RecentResponseFileName = l_File

            If String.Equals(_Status, "REJECTED", StringComparison.CurrentCultureIgnoreCase) Then
                File.Move(_RESPONSE_OUT & "\" & l_File, _RESPONSE_ARCHIVED & "\" & l_File)
            Else
                File.Copy(_RESPONSE_OUT & "\" & l_File, _RESPONSE_ARCHIVED & "\" & l_File, True)
            End If

            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function
#End Region

#Region "Methods Zupa--------------------------------------------------------------"
    Private Sub Process_Zupa()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer

        Dim bFailedToSendResponces As Boolean = False
        Dim lMsg As String = String.Empty
        Dim lResult As MsgBoxResult

        Try

            ''MyEventLog.WriteEntry("Monitoring the System", EventLogEntryType.Information,GetEventID(0)) 

            '' ''CHECK if recent responses have been sent away
            ''If _RecentResponseFileName.Length > 0 Then
            ''    asFiles = Directory.GetFiles(_RESPONSE_OUT)
            ''    For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
            ''        l_File = Path.GetFileName(asFiles(idx))
            ''        If bFailedToSendResponces = False Then
            ''            If l_File.Equals(_RecentResponseFileName, comparisonType:=StringComparison.InvariantCultureIgnoreCase) Then
            ''                bFailedToSendResponces = True
            ''                lMsg = "AS2 ERROR  " & Date.Now & ".  FAILED to send file(s):" & vbCrLf
            ''            End If
            ''        End If
            ''        If bFailedToSendResponces Then
            ''            lMsg &= l_File & vbCrLf
            ''        End If
            ''    Next
            ''    If bFailedToSendResponces Then
            ''        MyEventLog.WriteEntry(lMsg, EventLogEntryType.Error, GetEventID)

            ''        _EmailServiceMessage(lMsg, True)
            ''    End If

            ''    _RecentResponseFileName = String.Empty
            ''End If

            'PROCESS NEW ORDERS
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))
                'If l_File.StartsWith("POR") AndAlso Not l_File.EndsWith("_response.xml") Then
                MyEventLog.WriteEntry("ORDER REQUEST RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)

                If Not _Test_Mode Then
                    __LastOrderReceivedDatetime_Zupa = Date.Now
                    __LastWarningEmailDatetime_Zupa = __LastOrderReceivedDatetime_Zupa
                End If

                lResult = UploadFile_Zupa(_ORDER_IN & "\" & l_File)
                If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                    Return
                End If
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
    End Sub

    Private Function UploadFile_Zupa(ByVal pFileName As String) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand
        Dim lXMLContents As String = String.Empty
        Dim lReader As StreamReader
        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim dt As DataTable = Nothing
        Dim lNewOrderID As Integer = 0

        Try
            l_DB.Open()

            lReader = New StreamReader(pFileName)
            lXMLContents = lReader.ReadToEnd()
            lReader.Close()

            lXMLContents = lXMLContents.Replace("<sh:", "<").Replace("<eanucc:", "<").Replace("<order:", "<")
            lXMLContents = lXMLContents.Replace("</sh:", "</").Replace("</eanucc:", "</").Replace("</order:", "</")
            lXMLContents = lXMLContents.Replace("encoding=""utf-8""", "")
            cmd = l_DB.SqlCommand("p_P2P_order_import_Zupa")

            With cmd
                If _Test_Mode Then
                    .CommandTimeout = 120
                End If
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 200)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@xml_order", SqlDbType.Xml)
                param.Direction = ParameterDirection.Input
                param.Value = lXMLContents

                param = .Parameters.Add("@order_num", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@vendor_id", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@datetime_created_string", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@status", SqlDbType.VarChar, 16)
                param.Direction = ParameterDirection.InputOutput
                param.Value = "ACCEPTED"

                param = .Parameters.Add("@customer_order_header_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                .ExecuteNonQuery()
                lRetVal = CType(.Parameters("@ret").Value, Integer)

                _OrderRequestId = .Parameters("@record_id").Value
                _OrderNum = Nz(Of String)(.Parameters("@order_num").Value, "")
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "").Trim
                _VendorID = Nz(Of String)(.Parameters("@vendor_id").Value, "")
                _DeliveryDate = Nz(Of Date)(.Parameters("@delivery_date").Value, Date.MinValue)
                _Status = Nz(Of String)(.Parameters("@status").Value, "")
                _Order_DateTime_Created = Nz(Of String)(.Parameters("@datetime_created_string").Value, "")

                lNewOrderID = Nz(Of Integer)(.Parameters("@customer_order_header_id").Value, 0)

                lErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")

                Select Case lRetVal
                    Case 0 ' Success
                        If _Status = "MODIFIED_LINES" Then
                            dt = GetOrderAmendments()
                            _Status = "MODIFIED"
                        End If
                        If CreateOrderResponse_Zupa("", dt) Then
                            MyEventLog.WriteEntry("Order Response Created: " & _OrderNum & " {Request id = " & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                            If _Test_Mode Then
                                _EmailServiceMessage(Date.Now & " ORDER RESPONSE CREATED: " & _OrderNum & " {Request id = " & _OrderRequestId & "} for " & _AccNum)
                            End If
                            System.Threading.Thread.Sleep(50)

                            If EmailOrder(lNewOrderID) Then
                                Return MsgBoxResult.Ok
                            End If
                        End If

                        Return MsgBoxResult.Ok
                        'Case 1 To 333
                        '    _Status = "REJECTED"
                        '    Return CreateOrderResponse(lErrMsg)
                    Case Else
                        _Status = "REJECTED"
                        Dim lRejectCode As String = lErrMsg
                        If lErrMsg.Contains("~") Then
                            lRejectCode = lErrMsg.Substring(0, lErrMsg.IndexOf("~"))
                        End If
                        If lRetVal = 4 OrElse CreateOrderResponse_Zupa(lRejectCode, Nothing) Then ' if 4 = Duplicate Order : Already processed - do not create response
                            If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "
                            MyEventLog.WriteEntry("Order Response Created: REJECTED " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & " " & lErrMsg, EventLogEntryType.Warning, GetEventID)

                            _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & " " & lErrMsg)

                            System.Threading.Thread.Sleep(50)
                        End If
                End Select
            End With
            cmd.Dispose()
            l_DB.Close()
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function

    Private Function CreateOrderResponse_Zupa(p_RejectReason As String, ByVal dtLines As DataTable) As Boolean
        Dim xmlWriter As XmlTextWriter = Nothing
        Dim l_File As String = _OrderNum & "_response.xml"
        Dim row As DataRow 'invoice detail datarow

        Try

            If File.Exists(_RESPONSE_OUT & "\" & l_File) Then
                File.Delete(_RESPONSE_OUT & "\" & l_File)
            End If

            xmlWriter = New XmlTextWriter(_RESPONSE_OUT & "\" & l_File, System.Text.Encoding.UTF8)

            With xmlWriter
                .WriteStartDocument(True)
                .WriteStartElement("sh:StandardBusinessDocument")
                AddAttribute(xmlWriter, "xmlns:sh", "http://www.unece.org/cefact/namespaces/StandardBusinessDocumentHeader")
                AddAttribute(xmlWriter, "xmlns:order", "urn:ean.ucc:order:2")
                AddAttribute(xmlWriter, "xmlns:eanucc", "urn:ean.ucc:2")
                AddAttribute(xmlWriter, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
                AddAttribute(xmlWriter, "xsi:schemaLocation", "http://www.unece.org/cefact/namespaces/StandardBusinessDocumentHeader ../Schemas/sbdh/StandardBusinessDocumentHeader.xsd urn:ean.ucc:2 ../Schemas/OrderResponseProxy.xsd")

                .WriteStartElement("sh:StandardBusinessDocumentHeader")

                .WriteStartElement("sh:HeaderVersion")
                .WriteString("2.3")
                .WriteEndElement()

                .WriteStartElement("sh:Sender")
                .WriteStartElement("sh:Identifier") : AddAttribute(xmlWriter, "Authority", "EAN.UCC")
                .WriteString("West Country Milk")
                .WriteEndElement()
                .WriteEndElement()

                .WriteStartElement("sh:Receiver")
                .WriteStartElement("sh:Identifier") : AddAttribute(xmlWriter, "Authority", "EAN.UCC")
                .WriteString("CHC001") ' TODO Ask Matthew from Zupa
                .WriteEndElement()
                .WriteEndElement()


                .WriteStartElement("sh:DocumentIdentification")
                .WriteStartElement("sh:Standard") : .WriteString("EAN.UCC") : .WriteEndElement()
                .WriteStartElement("sh:TypeVersion") : .WriteString("2.3") : .WriteEndElement()
                .WriteStartElement("sh:InstanceIdentifier") : .WriteString("1111") : .WriteEndElement()
                .WriteStartElement("sh:Type") : .WriteString("OrderResponse") : .WriteEndElement()
                .WriteStartElement("sh:CreationDateAndTime") : .WriteString(Now.Date.ToString("yyyy-MM-dd") & "T" & Now.ToString("HH:mm:ss") & "+00:00") : .WriteEndElement()
                .WriteEndElement()

                .WriteEndElement()  'sh:StandardBusinessDocumentHeader

                .WriteStartElement("eanucc:message")
                .WriteStartElement("entityIdentification")
                .WriteStartElement("uniqueCreatorIdentification") : .WriteString("2222") : .WriteEndElement()
                .WriteStartElement("contentOwner")
                .WriteStartElement("gln") : .WriteString(My.Settings.GLN_WCM) : .WriteEndElement() '
                .WriteEndElement()
                .WriteEndElement() 'entityIdentification

                .WriteStartElement("order:orderResponse")
                AddAttribute(xmlWriter, "creationDateTime", Now.Date.ToString("yyyy-MM-dd") & "T" & Now.ToString("HH:mm:ss") & "+00:00")
                AddAttribute(xmlWriter, "documentStatus", "ORIGINAL")
                AddAttribute(xmlWriter, "responseStatusType", _Status)

                .WriteStartElement("responseIdentification")
                .WriteStartElement("uniqueCreatorIdentification") : .WriteString(_OrderNum) : .WriteEndElement()
                .WriteStartElement("contentOwner")
                .WriteStartElement("gln") : .WriteString(My.Settings.GLN_WCM) : .WriteEndElement()
                .WriteEndElement()
                .WriteEndElement() ' responseIdentification

                .WriteStartElement("responseToOriginalDocument")
                AddAttribute(xmlWriter, "referenceDocumentType", "35")
                AddAttribute(xmlWriter, "referenceIdentification", _OrderNum)
                AddAttribute(xmlWriter, "referenceDateTime", _Order_DateTime_Created)
                .WriteEndElement()

                .WriteStartElement("buyer")
                .WriteStartElement("gln") : .WriteString("5060260760064") : .WriteEndElement() ' verify with Matthew
                .WriteStartElement("additionalPartyIdentification")
                .WriteStartElement("additionalPartyIdentificationValue") : .WriteString(_AccNum) : .WriteEndElement() 'Ch & Co customer account number
                .WriteStartElement("additionalPartyIdentificationType") : .WriteString("BUYER_ASSIGNED_IDENTIFIER_FOR_A_PARTY") : .WriteEndElement()
                .WriteEndElement()
                .WriteStartElement("additionalPartyIdentification")
                .WriteStartElement("additionalPartyIdentificationValue") : .WriteString(_AccNum) : .WriteEndElement() 'mapped wcm customer account number (it is the same as buyer for Ch & Co)
                .WriteStartElement("additionalPartyIdentificationType") : .WriteString("SELLER_ASSIGNED_IDENTIFIER_FOR_A_PARTY") : .WriteEndElement()
                .WriteEndElement()
                .WriteEndElement() ' buyer

                .WriteStartElement("seller")
                .WriteStartElement("gln") : .WriteString(My.Settings.GLN_WCM) : .WriteEndElement()
                .WriteStartElement("additionalPartyIdentification")
                '' .WriteStartElement("additionalPartyIdentificationValue") : .WriteString(_VendorID) : .WriteEndElement() 'Ch & Co customer account number
                .WriteStartElement("additionalPartyIdentificationValue") : .WriteString(_AccNum) : .WriteEndElement() 'Ch & Co customer account number 
                .WriteStartElement("additionalPartyIdentificationType") : .WriteString("SELLER_ASSIGNED_IDENTIFIER_FOR_A_PARTY") : .WriteEndElement()
                .WriteEndElement()
                .WriteEndElement() ' seller

                If String.Equals(_Status, "REJECTED", StringComparison.CurrentCultureIgnoreCase) Then
                    .WriteStartElement("orderResponseReasonCode")
                    .WriteString(p_RejectReason)
                    .WriteEndElement()

                ElseIf String.Equals(_Status, "MODIFIED", StringComparison.CurrentCultureIgnoreCase) Then
                    .WriteStartElement("orderModification")

                    .WriteStartElement("amendedDateTimeValue")
                    .WriteStartElement("requestedDeliveryDate")
                    .WriteStartElement("date") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd")) : .WriteEndElement()
                    .WriteStartElement("time") : .WriteString("09:00:00") : .WriteEndElement()
                    .WriteEndElement()
                    .WriteEndElement() 'amendedDateTimeValue

                    If dtLines IsNot Nothing Then
                        For Each row In dtLines.Rows
                            .WriteStartElement("orderModificationLineItemLevel")
                            If row("Product_Code_Requested") <> row("Product_Code") AndAlso row("Qty") <> 0 Then ' new code
                                'IF substituted Then insert this block -------------------------------
                                ''If row("Product_Code_Requested") <> row("Product_Code") Then
                                If Not String.IsNullOrEmpty(row("Product_Code_Requested")) Then
                                    ' added delivery charge generates line with blank "ordered" code - do not output this line
                                    .WriteStartElement("substituteItemIdentification")
                                    .WriteStartElement("gtin") : .WriteString("55555555555555") : .WriteEndElement()

                                    .WriteStartElement("additionalTradeItemIdentification")
                                    .WriteStartElement("additionalTradeItemIdentificationValue") : .WriteString(row("Product_Code_Requested")) : .WriteEndElement()
                                    .WriteStartElement("additionalTradeItemIdentificationType") : .WriteString("SUPPLIER_ASSIGNED") : .WriteEndElement()
                                    .WriteEndElement() 'additionalTradeItemIdentification>

                                    .WriteEndElement() 'substituteItemIdentification
                                End If
                            End If ' end of substituted item block --------------------------------------------------------------------

                            .WriteStartElement("modifiedOrderInformation") : AddAttribute(xmlWriter, "number", row("Line_Seq").ToString)
                            .WriteStartElement("requestedQuantity")
                            .WriteStartElement("value") : .WriteString(row("Qty").ToString) : .WriteEndElement() ' NOTE Qty 0 = can not supply this item and no substitude available
                            .WriteStartElement("unitOfMeasure")
                            .WriteStartElement("measurementUnitCodeValue") : .WriteString("EA") : .WriteEndElement()
                            .WriteEndElement() 'unitOfMeasure
                            .WriteEndElement() 'requestedQuantity

                            .WriteStartElement("tradeItemIdentification")
                            .WriteStartElement("gtin") : .WriteString("00000000000000") : .WriteEndElement() ' The supplier substituted item has a GTIN of 05024333133919
                            .WriteStartElement("additionalTradeItemIdentification")
                            If row("Qty") <> 0 AndAlso Not String.IsNullOrEmpty(row("Product_Code")) Then
                                .WriteStartElement("additionalTradeItemIdentificationValue") : .WriteString(row("Product_Code")) : .WriteEndElement() ''ProductCode_Substitude
                            Else
                                .WriteStartElement("additionalTradeItemIdentificationValue") : .WriteString(row("Product_Code_Requested")) : .WriteEndElement() '' Original product code
                            End If
                            .WriteStartElement("additionalTradeItemIdentificationType") : .WriteString("SUPPLIER_ASSIGNED") : .WriteEndElement()
                            .WriteEndElement() 'additionalTradeItemIdentification
                            .WriteEndElement() 'tradeItemIdentification


                            .WriteEndElement() 'modifiedOrderInformation
                            .WriteStartElement("orderResponseReasonCode") : .WriteString("PRODUCT_OUT_OF_STOCK") : .WriteEndElement()

                            .WriteEndElement() 'orderModificationLineItemLevel
                        Next row 'end for each item
                    End If
                    .WriteEndElement() 'orderModification
                End If

                .WriteEndElement() ' order:orderResponse

                .WriteEndElement() ' eanucc:message

                .WriteEndElement()  'end of Root element: sh:StandardBusinessDocument
                .WriteEndDocument()
                .Flush()
                .Close()
            End With

            _RecentResponseFileName = l_File

            File.Copy(_RESPONSE_OUT & "\" & l_File, _RESPONSE_ARCHIVED & "\" & l_File, True)

            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function
#End Region
#Region "Methods FoodBuy_Online--------------------------------------------------------------"

    Private Sub Process_FoodBuy_Online()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer
        'Dim l_Directory As String = My.Settings.ELIOR_Waterfall_IN

        'Dim l_OrderAcknowledgementDirectory As String = My.Settings.ELIOR_Waterfall_OUT & "\"
        'Dim l_OrderResponseDirectory As String = My.Settings.ELIOR_Waterfall_OUT & "\"

        Dim lResult As MsgBoxResult

        Try

            'PROCESS NEW ORDERS
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))
                MyEventLog.WriteEntry("ORDER REQUEST RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)
                If Not _Test_Mode Then
                    __LastOrderReceivedDatetime_FoodBuy_Online = Date.Now
                    __LastWarningEmailDatetime_FoodBuy_Online = __LastOrderReceivedDatetime_FoodBuy_Online
                End If
                lResult = UploadFile_FoodBuy_Online(_ORDER_IN & "\" & l_File)
                If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                    Return
                End If
            Next


        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
    End Sub

    Private Function UploadFile_FoodBuy_Online(ByVal pFileName As String) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand
        Dim lXMLContents As String = String.Empty
        Dim lReader As StreamReader
        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim dt As DataTable = Nothing
        Dim lNewOrderID As Integer = 0

        Try
            l_DB.Open()

            lReader = New StreamReader(pFileName)
            lXMLContents = lReader.ReadToEnd()
            lReader.Close()

            lXMLContents = lXMLContents.Replace(" encoding=""UTF-8""", "")
            lXMLContents = lXMLContents.Replace("xmlns=""urn:www-basda-org/schema/purord.xml""", "")
            ''lXMLContents = lXMLContents.Replace("<ns0:", "<").Replace("</ns0:", "</")
            cmd = l_DB.SqlCommand("p_P2P_order_import_fbo")

            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 100)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@xml_order", SqlDbType.Xml)
                param.Direction = ParameterDirection.Input
                param.Value = lXMLContents

                param = .Parameters.Add("@order_num", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@vendor_id", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@delivery_date_requested", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@datetime_created_string", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@status", SqlDbType.VarChar, 16)
                param.Direction = ParameterDirection.InputOutput
                param.Value = "ACCEPTED"

                param = .Parameters.Add("@customer_order_header_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@order_lines", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@order_value", SqlDbType.Decimal)
                param.Precision = 10
                param.Scale = 4
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                .ExecuteNonQuery()
                lRetVal = CType(.Parameters("@ret").Value, Integer)

                _OrderRequestId = .Parameters("@record_id").Value
                _OrderNum = Nz(Of String)(.Parameters("@order_num").Value, "")
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "")
                _VendorID = Nz(Of String)(.Parameters("@vendor_id").Value, "")
                _DeliveryDateRequested = Nz(Of Date)(.Parameters("@delivery_date_requested").Value, Date.MinValue)
                _DeliveryDate = Nz(Of Date)(.Parameters("@delivery_date").Value, Date.MinValue)
                _Status = Nz(Of String)(.Parameters("@status").Value, "")
                _Order_DateTime_Created = Nz(Of String)(.Parameters("@datetime_created_string").Value, "")
                _Order_Lines = Nz(Of Integer)(.Parameters("@order_lines").Value, 0)
                _Order_Value = Nz(Of Decimal)(.Parameters("@order_value").Value, 0)

                lNewOrderID = Nz(Of Integer)(.Parameters("@customer_order_header_id").Value, 0)

                lErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")

                'If CreateOrderResponse_FoodBuy_Online(1, Nothing, lRetVal, "") Then
                '    MyEventLog.WriteEntry("Order Acknowledgement Created: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information,GetEventID)
                '    If My.Settings.TestFoodBuy_Online Then
                '        _EmailServiceMessage(Date.Now & " ORDER ACKNOWLEDGEMENT CREATED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum)
                '    End If
                System.Threading.Thread.Sleep(20)
                'End If

                dt = Nothing ' GetOrderLines() No Need for order lines for FoodBuy_Online response (no response needed anyway)
                Dim l_OK As Boolean = False
                Select Case lRetVal
                    Case 0 ' Success
                        If _Status.StartsWith("MODIFIED") Then
                            dt = GetOrderLines() ' No Need for order lines for FoodBuy_Online response (no response needed anyway), just for service email
                            _Status = "MODIFIED"
                            l_OK = CreateOrderResponse_FoodBuy_Online(1, dt, lRetVal, "", "")
                        Else
                            l_OK = CreateOrderResponse_FoodBuy_Online(0, dt, lRetVal, "", "")
                        End If

                        If l_OK Then
                            MyEventLog.WriteEntry("Order Response Created: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                            If _Test_Mode Then
                                _EmailServiceMessage(Date.Now & " ORDER RESPONSE CREATED: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum)
                            End If
                            System.Threading.Thread.Sleep(50)

                            If EmailOrder(lNewOrderID) Then

                            End If
                        End If

                        Return MsgBoxResult.Ok
                    Case Else
                        _Status = "REJECTED"
                        Dim lRejectCode As String = lErrMsg
                        If lErrMsg.Contains("~") Then
                            lRejectCode = lErrMsg.Substring(0, lErrMsg.IndexOf("~"))
                        End If
                        If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "
                        MyEventLog.WriteEntry("Order Response Created: REJECTED " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg, EventLogEntryType.Warning, GetEventID)
                        _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg)
                        If CreateOrderResponse_FoodBuy_Online(2, dt, lRetVal, lErrMsg, lRejectCode) Then


                            Return MsgBoxResult.Abort
                        End If
                End Select
            End With
            cmd.Dispose()
            l_DB.Close()
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pOrderResponseType">1=Acknowledgement, 2=order confirmation, 3=order ASN</param>
    ''' <returns></returns>
    ''' <remarks> </remarks>
    Private Function CreateOrderResponse_FoodBuy_Online(pOrderResponseType As Integer, dtLines As DataTable, pErrCode As Integer, pErrMsg As String, pAdditionalInfo As String) As Boolean
        Dim xmlWriter As XmlWriter = Nothing
        Dim dtProcessDate As Date = Now
        Dim l_File As String = ""
        Dim row As DataRow
        Dim l_rowcount As Integer = 0
        Dim dblActualTotalValue As Decimal = 0

        Try
            Select Case pOrderResponseType
                'Case 1
                '    l_File = "ACK_" & _OrderNum & ".XML"
                '    If My.Settings.TestFoodBuy_Online Then l_Directory = My.Settings.FoodBuy_Online_OUT & "\" Else l_Directory = My.Settings.FoodBuy_Online_OUT & "\"
                Case 2
                    l_File = "Order_Response_" & _OrderNum & ".XML"
                    ''Case 3 :
            End Select

            'NOTE - no need to create order response, just update database and send serving emails

            'If File.Exists(l_Directory & l_File) Then File.Delete(l_Directory & l_File)
            'xmlWriter = New XmlTextWriter(l_Directory & l_File, System.Text.Encoding.UTF8)

            'If IsNothing(xmlWriter) Then
            '    Return False
            'End If

            'If dtLines IsNot Nothing Then
            '    For Each row In dtLines.Rows
            '        dblActualTotalValue += row("Qty") * Nz(Of Decimal)(row("Unit_Price"), 0)
            '    Next
            'End If

            'With xmlWriter
            '    .WriteStartDocument()
            '    .WriteStartElement("ns0:GenericITNSupplierResponse") : AddAttribute(xmlWriter, "xmlns:ns0", "http://itn.hub.biztalk.orderapp.ext.GenericITNSupplier.schemas.OrderResponse")

            '    .WriteStartElement("OrderHeader")

            '    .WriteStartElement("OrderResponseType") : .WriteString(pOrderResponseType.ToString) : .WriteEndElement()
            '    .WriteStartElement("ItnOrderNumber") : .WriteString(_OrderNum) : .WriteEndElement()
            '    .WriteStartElement("DeliveryDate") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd") & "T09:00:00") : .WriteEndElement()

            '    If pOrderResponseType = 1 Then
            '        .WriteStartElement("ItnOrderStatus") : .WriteString("ACKNOWLEDGED") : .WriteEndElement()
            '        .WriteStartElement("OrderValue") : .WriteString(_Order_Value.ToString("0.00")) : .WriteEndElement()
            '    Else
            '        .WriteStartElement("ItnOrderStatus") : .WriteString(_Status) : .WriteEndElement() ' ACCEPTED, REJECTED, MODIFIED
            '        .WriteStartElement("OrderValue") : .WriteString(dblActualTotalValue.ToString("0.00")) : .WriteEndElement()
            '    End If

            '    .WriteStartElement("OrderResponseDate") : .WriteString(Now.Date.ToString("yyyy-MM-dd") & "T" & Now.ToString("HH:mm:ss")) : .WriteEndElement()
            '    .WriteStartElement("SupplierReasonCode")
            '    If _Status.Equals("REJECTED", StringComparison.CurrentCultureIgnoreCase) Then
            '        Select Case pErrCode
            '            ' 1=Customer account on stop; 2=Customer account is not recognised; 3=Delivery date changed
            '            Case 6 ' 'CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST'
            '                .WriteString("2") ' 2=Customer account is not recognised
            '            Case 7 'DELIVERY_SLOT_MISSED'
            '                .WriteString("3") '3=Delivery date changed
            '        End Select
            '    ElseIf _Status.Equals("MODIFIED", StringComparison.CurrentCultureIgnoreCase) Then
            '        If DateDiff(DateInterval.Day, _DeliveryDateRequested, _DeliveryDate) <> 0 Then
            '            .WriteString("3") ' 3=Delivery date changed
            '        End If
            '    End If
            '    .WriteEndElement()

            '    .WriteStartElement("OriginalOrderCreationDate")
            '    If IsDate(_Order_DateTime_Created) Then
            '        .WriteString(CType(_Order_DateTime_Created, Date).ToString("yyyy-MM-dd") & "T00:00:00")
            '    Else
            '        .WriteString(Now.Date.ToString("yyyy-MM-dd") & "T" & Now.ToString("HH:mm:ss"))
            '    End If
            '    .WriteEndElement()

            '    .WriteStartElement("OriginalOrderLines") : .WriteString(_Order_Lines.ToString) : .WriteEndElement()
            '    .WriteStartElement("CustomerAccountNumber") : .WriteString(_AccNum) : .WriteEndElement()

            '    .WriteEndElement() '/OrderHeader

            '    .WriteStartElement("SupplierDetails")
            '    .WriteStartElement("SupplierId") : .WriteString("WSTCTMKUK") : .WriteEndElement()
            '    .WriteStartElement("SupplierGLN") : .WriteString(My.Settings.GLN_WCM) : .WriteEndElement()
            '    .WriteStartElement("SupplierOrderNumber") : .WriteString(_OrderNum) : .WriteEndElement()
            '    .WriteEndElement() '/SupplierDetails

            '    .WriteStartElement("BuyerDetails")
            '    .WriteStartElement("BuyerGLN") : .WriteString(My.Settings.GLN_COMPASS) : .WriteEndElement()
            '    .WriteStartElement("ShipToId") : .WriteEndElement()
            '    .WriteStartElement("ShipToGLN") : .WriteEndElement()

            '    .WriteStartElement("BuyerGroup")
            '    .WriteStartElement("BuyerGroupGLN") : .WriteString("5013546085276") : .WriteEndElement()
            '    .WriteStartElement("BuyerGroupName") : .WriteString("COMPASSUK") : .WriteEndElement()
            '    .WriteEndElement() '/BuyerGroup

            '    .WriteEndElement() '/BuyerDetails

            '    .WriteStartElement("LineItems")

            If dtLines IsNot Nothing Then
                For Each row In dtLines.Rows
                    l_rowcount += 1
                    '            .WriteStartElement("LineItem")
                    '            .WriteStartElement("ItnLineNumber") : .WriteString(row("Line_Seq").ToString) : .WriteEndElement()

                    '            .WriteStartElement("ProductCode")
                    '            'If row("Product_Code").ToString.Length = 0 Then
                    '            .WriteString(row("Product_Code_Requested")) ' the same as Product_Code if there is no substitution 
                    '            'Else
                    '            '.WriteString(row("Product_Code")) 
                    '            'End If
                    '            .WriteEndElement()

                    '            .WriteStartElement("Quantity") : .WriteString(row("Qty").ToString) : .WriteEndElement() 'NOTE Qty 0 = can not supply this item and no substitude available
                    '            .WriteStartElement("UnitPrice") : .WriteString(Nz(Of Decimal)(row("Unit_Price"), 0).ToString("0.00")) : .WriteEndElement()
                    '            .WriteStartElement("LinePrice") : .WriteString((Nz(Of Decimal)(row("Unit_Price"), 0) * CType(row("Qty"), Decimal)).ToString("0.00")) : .WriteEndElement()

                    '            If row("Product_Code").ToString.Length = 0 Then
                    '                .WriteStartElement("LineStatus") : .WriteString("OS") : .WriteEndElement()
                    '            ElseIf row("Product_Code_Requested") <> row("Product_Code") Then
                    '                .WriteStartElement("SubstitutedProduct")
                    '                .WriteString(row("Product_Code"))
                    '                .WriteEndElement() 'SubstitutedProduct
                    '                .WriteStartElement("LineStatus") : .WriteString("OOS") : .WriteEndElement()
                    '            Else
                    '                .WriteStartElement("LineStatus") : .WriteEndElement()
                    '            End If

                    '            .WriteStartElement("UnitOfMeasure") : .WriteString(row("UOM")) : .WriteEndElement()

                    '            .WriteEndElement()
                Next
            End If
            '    Else
            '        .WriteStartElement("LineItem") : .WriteEndElement()
            '    End If

            '    .WriteEndElement() '/LineItems

            '    .WriteEndElement() ' End of ns0:GenericITNSupplierResponse
            '    .WriteEndDocument()
            '    .Flush()
            '    .Close()

            'End With

            If UpdateAcknowledgementDate() = 0 Then
                'MoveFile( _RESPONSE_OUT & "\" & l_File, _RESPONSE_ARCHIVED & "\" & l_File)
            End If


            If pOrderResponseType = 2 Then
                'MoveFile(_RESPONSE_OUT & "\" & l_File, _RESPONSE_ARCHIVED & "\" & l_File)

                If _Status.Equals("REJECTED", StringComparison.CurrentCultureIgnoreCase) Then
                    Select Case pErrCode
                        ' 1=Customer account on stop; 2=Customer account is not recognised; 3=Delivery date changed
                        Case 6 ' 'CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST'
                            _EmailServiceMessage("REJECTED: " & _OrderNum & "   " & _AccNum & "  {CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST}", True)
                        Case 7 'DELIVERY_SLOT_MISSED'
                            _EmailServiceMessage("REJECTED: " & _OrderNum & "   " & _AccNum & "  {DELIVERY_SLOT_MISSED}  " & _DeliveryDateRequested, True)
                        Case Else
                            _EmailServiceMessage("REJECTED: " & _OrderNum & "   " & _AccNum & "  " & pErrCode & " - " & pErrMsg, True)
                    End Select
                ElseIf pOrderResponseType = 1 Then
                    If _Status.Equals("MODIFIED", StringComparison.CurrentCultureIgnoreCase) Then
                        If DateDiff(DateInterval.Day, _DeliveryDateRequested, _DeliveryDate) <> 0 Then
                            _EmailServiceMessage("MODIFIED: " & _OrderNum & "   " & _AccNum & "  {DELIVERY_DATE_CHANGED}  " & _DeliveryDateRequested & " --> " & _DeliveryDate, True)
                        Else
                            _EmailServiceMessage("MODIFIED: " & _OrderNum & "   " & _AccNum & "  {MODIFIED ORDER LINES}  requested: " & _Order_Lines.ToString & "  ordered:" & l_rowcount.ToString, True)
                        End If
                    End If
                End If
            End If

            Return True

        Catch ex As Exception
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False

    End Function

#End Region

#Region "Methods Poundland--------------------------------------------------------------"

    Private Sub Process_Poundland()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer

        Dim lResult As MsgBoxResult

        Try
            ''MyEventLog.WriteEntry("Monitoring the System", EventLogEntryType.Information, GetEventID))

            'PROCESS NEW ORDERS
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))
                MyEventLog.WriteEntry("ORDER REQUEST RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)
                If Not _Test_Mode Then
                    __LastOrderReceivedDatetime_Poundland = Date.Now
                    __LastWarningEmailDatetime_Poundland = __LastOrderReceivedDatetime_Poundland
                End If
                lResult = UploadFile_Poundland(_ORDER_IN & "\" & l_File)
                If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                    Return
                End If
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
    End Sub

    Private Function UploadFile_Poundland(ByVal pFileName As String) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand
        Dim lXMLContents As String = String.Empty
        Dim lReader As StreamReader
        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim dt As DataTable = Nothing
        Dim lNewOrderID As Integer = 0

        Try
            l_DB.Open()

            lReader = New StreamReader(pFileName)
            lXMLContents = lReader.ReadToEnd()
            lReader.Close()

            lXMLContents = lXMLContents.Replace("encoding=""UTF-8""", "")
            lXMLContents = lXMLContents.Replace("<tc:", "<").Replace("</tc:", "</")

            cmd = l_DB.SqlCommand("p_P2P_order_import_tc")

            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 100)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@xml_order", SqlDbType.Xml)
                param.Direction = ParameterDirection.Input
                param.Value = lXMLContents

                param = .Parameters.Add("@order_num", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@vendor_id", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@delivery_date_requested", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@datetime_created_string", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@status", SqlDbType.VarChar, 16)
                param.Direction = ParameterDirection.InputOutput
                param.Value = "ACCEPTED"

                param = .Parameters.Add("@customer_order_header_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@order_lines", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@order_value", SqlDbType.Decimal)
                param.Precision = 10
                param.Scale = 4
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                .ExecuteNonQuery()
                lRetVal = CType(.Parameters("@ret").Value, Integer)

                _OrderRequestId = .Parameters("@record_id").Value
                _OrderNum = Nz(Of String)(.Parameters("@order_num").Value, "")
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "")
                _VendorID = Nz(Of String)(.Parameters("@vendor_id").Value, "")
                _DeliveryDateRequested = Nz(Of Date)(.Parameters("@delivery_date_requested").Value, Date.MinValue)
                _DeliveryDate = Nz(Of Date)(.Parameters("@delivery_date").Value, Date.MinValue)
                _Status = Nz(Of String)(.Parameters("@status").Value, "")
                _Order_DateTime_Created = Nz(Of String)(.Parameters("@datetime_created_string").Value, "")
                _Order_Lines = Nz(Of Integer)(.Parameters("@order_lines").Value, 0)
                _Order_Value = Nz(Of Decimal)(.Parameters("@order_value").Value, 0)

                lNewOrderID = Nz(Of Integer)(.Parameters("@customer_order_header_id").Value, 0)

                lErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")

                System.Threading.Thread.Sleep(20)

                dt = GetOrderLines() ' No Need for order lines for FoodBuy_Online response (no response needed anyway), just for service email
                Dim l_OK As Boolean = False
                Select Case lRetVal
                    Case 0 ' Success
                        If _Status.StartsWith("MODIFIED") Then
                            _Status = "MODIFIED"
                        End If
                        l_OK = CreateOrderResponse_Poundland(0, dt, lRetVal, "", "")

                        If l_OK Then
                            MyEventLog.WriteEntry("Order Response Created: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                            If _Test_Mode Then
                                _EmailServiceMessage(Date.Now & " ORDER RESPONSE CREATED: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum)
                            End If
                            System.Threading.Thread.Sleep(50)

                            Dim pAttachments As New ArrayList
                            pAttachments.Add("\\wcm-exe-fp01\OfficeShare\Poundland\Email Attachment\West Country Milk - Poundland Service Level Agreement.pdf")

                            If EmailOrder(lNewOrderID, pAttachments, "</br> </br> Please see attached SLA for this order") Then

                            End If
                        End If

                        Return MsgBoxResult.Ok
                    Case Else
                        _Status = "REJECTED"
                        Dim lRejectCode As String = lErrMsg
                        If lErrMsg.Contains("~") Then
                            lRejectCode = lErrMsg.Substring(0, lErrMsg.IndexOf("~"))
                        End If

                        If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "
                        MyEventLog.WriteEntry("Order Response Created: REJECTED " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg, EventLogEntryType.Warning, GetEventID)
                        _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg)

                        If CreateOrderResponse_Poundland(2, dt, lRetVal, lErrMsg, lRejectCode) Then

                            Return MsgBoxResult.Abort
                        End If
                End Select
            End With
            cmd.Dispose()
            l_DB.Close()
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pOrderResponseType">1=Acknowledgement, 2=order confirmation, 3=order ASN</param>
    ''' <returns></returns>
    ''' <remarks> </remarks>
    Private Function CreateOrderResponse_Poundland(pOrderResponseType As Integer, dtLines As DataTable, pErrCode As Integer, pErrMsg As String, pAdditionalInfo As String) As Boolean
        Dim xmlWriter As XmlWriter = Nothing
        Dim dtProcessDate As Date = Now
        Dim l_File As String = ""
        'Dim row As DataRow
        Dim l_rowcount As Integer = 0
        Dim dblActualTotalValue As Decimal = 0

        Try
            Select Case pOrderResponseType
                'Case 1
                '    l_File = "ACK_" & _OrderNum & ".XML"
                '    If My.Settings.TestFoodBuy_Online Then l_Directory = My.Settings.FoodBuy_Online_OUT & "\" Else l_Directory = My.Settings.FoodBuy_Online_OUT & "\"
                Case 2
                    l_File = "OrderAck_" & _OrderNum & ".XML"
                    ''Case 3 :
            End Select

            ' - FINISH CODING IF REQUIRED by Poundland
            'If File.Exists(_RESPONSE_OUT & "\" & l_File) Then
            '    File.Delete(_RESPONSE_OUT & "\" & l_File)
            'End If

            'xmlWriter = New XmlTextWriter(_RESPONSE_OUT & "\" & l_File, System.Text.Encoding.UTF8)

            'If IsNothing(xmlWriter) Then
            '    Return False
            'End If

            ''If dtLines IsNot Nothing Then
            ''    For Each row In dtLines.Rows
            ''        dblActualTotalValue += row("Qty") * Nz(Of Decimal)(row("Unit_Price"), 0)
            ''    Next
            ''End If

            'With xmlWriter
            '    .WriteStartDocument()
            '    .WriteStartElement("tc:TrueCommerceOrderAck")
            '    AddAttribute(xmlWriter, "xmlns:tc", "http://www.truecommerce.com/docs/orderAck")
            '    AddAttribute(xmlWriter, "xmlns:cm", "http://www.truecommerce.com/docs/common/components")
            '    AddAttribute(xmlWriter, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            '    AddAttribute(xmlWriter, "xsi:schemaLocation", "http://www.truecommerce.com/docs/orderAck truecommerce_orderack_v1.0.xsd")

            '    .WriteStartElement("MsgHeader")

            '    .WriteStartElement("MsgType") : .WriteString("ORDER_ACK") : .WriteEndElement()
            '    .WriteStartElement("VersionID") : .WriteString("v1.0") : .WriteEndElement()
            '    .WriteStartElement("TransmittedDate")
            '    .WriteStartElement("Date") : .WriteString(Now.Date.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString(Now.ToString("HH:mm:ss")) : .WriteEndElement()
            '    .WriteEndElement()
            '    .WriteStartElement("FileDate")
            '    .WriteStartElement("Date") : .WriteString(Now.Date.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString(Now.ToString("HH:mm:ss")) : .WriteEndElement()
            '    .WriteEndElement()
            '    .WriteStartElement("Interchange")
            '    .WriteStartElement("Standard") : .WriteString("TRUECOMMERCE") : .WriteEndElement()
            '    .WriteStartElement("SenderRef") : .WriteString("0000") : .WriteEndElement() 'Sender Transmission reference 
            '    .WriteEndElement()

            '    .WriteEndElement() ' End of MsgHeader

            '    .WriteStartElement("Document")

            '    .WriteStartElement("DocHeader")

            '    .WriteStartElement("DocType") : .WriteString("ORDER_ACK") : .WriteEndElement()
            '    .WriteStartElement("DocFunction") : .WriteString("ORIGINAL") : .WriteEndElement()

            '    .WriteStartElement("CustAddr")
            '    .WriteStartElement("Code") : .WriteString(_AccNum) : .WriteEndElement()
            '    .WriteStartElement("EAN13") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address5") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Country") : .WriteString("") : .WriteEndElement() ' get from original order if require
            '    .WriteStartElement("VAT")
            '    .WriteStartElement("Num") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteEndElement() ' End of VAT

            '    .WriteStartElement("Contact1")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact1
            '    .WriteStartElement("Contact2")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact2

            '    .WriteEndElement() ' End of CustAddr

            '    .WriteStartElement("SuppAddr")
            '    .WriteStartElement("Code") : .WriteString("") : .WriteEndElement() ' ??
            '    .WriteStartElement("EAN13") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address2") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address5") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("PostCode") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Country") : .WriteString("") : .WriteEndElement() ' get from original order if required

            '    .WriteStartElement("VAT")
            '    .WriteStartElement("Num") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteStartElement("Alpha") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteEndElement() ' End of VAT

            '    .WriteStartElement("Contact1")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Email") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact1
            '    .WriteStartElement("Contact2")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Email") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact2

            '    .WriteEndElement() ' End of SuppAddr

            '    .WriteStartElement("DocDate")
            '    .WriteStartElement("Date") : .WriteString(Now.Date.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString(Now.ToString("HH:mm:ss")) : .WriteEndElement()
            '    .WriteEndElement() ' End of DocDate

            '    .WriteStartElement("TranInfo") : .WriteEndElement()

            '    .WriteStartElement("RoutingCode") : .WriteString("") : .WriteEndElement() ' - what is it?

            '    .WriteEndElement() ' End of DocHeader


            '    .WriteStartElement("AckHeader")
            '    .WriteStartElement("AckCode") : .WriteString(_Status) : .WriteEndElement()

            '    .WriteStartElement("Ack_Date")
            '    .WriteStartElement("Date") : .WriteString(Now.Date.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString(Now.ToString("HH:mm:ss")) : .WriteEndElement()
            '    .WriteEndElement()

            '    .WriteStartElement("CustOrder") : .WriteString(_OrderNum) : .WriteEndElement()

            '    .WriteStartElement("CustOrderDate")
            '    If IsDate(_Order_DateTime_Created) Then
            '        .WriteStartElement("Date") : .WriteString(CType(_Order_DateTime_Created, Date).ToString("yyyy-MM-dd")) : .WriteEndElement()
            '        .WriteStartElement("Time") : .WriteString(CType(_Order_DateTime_Created, Date).ToString("HH:mm:ss")) : .WriteEndElement()
            '    Else
            '        .WriteStartElement("Date") : .WriteString(Now.Date.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '        .WriteStartElement("Time") : .WriteString(Now.ToString("HH:mm:ss")) : .WriteEndElement()
            '    End If
            '    .WriteEndElement()

            '    .WriteStartElement("SuppOrder") : .WriteString("") : .WriteEndElement() ' get from original order if required       SO0010004
            '    .WriteStartElement("Perishable") : .WriteString("") : .WriteEndElement() ' get from original order if required      false
            '    .WriteStartElement("BookingRef") : .WriteString("") : .WriteEndElement() ' get from original order if required      BR-001-0004
            '    .WriteStartElement("OrigCustOrder") : .WriteString("") : .WriteEndElement() ' get from original order if required   OCR-001-0004


            '    .WriteStartElement("Delivery")

            '    .WriteStartElement("DelMethod") : .WriteString("DELIVER_TO_DEPOT") : .WriteEndElement()

            '    .WriteStartElement("ReqDesp")
            '    .WriteStartElement("Date") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString("21:59:00") : .WriteEndElement() ' - read from original order ??
            '    .WriteEndElement()
            '    .WriteStartElement("EarliestDesp")
            '    .WriteStartElement("Date") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString("22:59:00") : .WriteEndElement() ' - read from original order ??
            '    .WriteEndElement()
            '    .WriteStartElement("LatestDesp")
            '    .WriteStartElement("Date") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString("22:59:00") : .WriteEndElement() ' - read from original order ??
            '    .WriteEndElement()
            '    .WriteStartElement("ReqDel")
            '    .WriteStartElement("Date") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString("22:59:00") : .WriteEndElement() ' - read from original order ??
            '    .WriteEndElement()
            '    .WriteStartElement("EarliestDel")
            '    .WriteStartElement("Date") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString("22:59:00") : .WriteEndElement() ' - read from original order ??
            '    .WriteEndElement()
            '    .WriteStartElement("LatestDel")
            '    .WriteStartElement("Date") : .WriteString(_DeliveryDate.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString("22:59:00") : .WriteEndElement() ' - read from original order ??
            '    .WriteEndElement()

            '    .WriteStartElement("DespatchFrom")
            '    .WriteStartElement("Code") : .WriteString("") : .WriteEndElement() ' ??
            '    .WriteStartElement("EAN13") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address2") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address5") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("PostCode") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Country") : .WriteString("") : .WriteEndElement() ' get from original order if required

            '    .WriteStartElement("VAT")
            '    .WriteStartElement("Num") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteStartElement("Alpha") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteEndElement() ' End of VAT

            '    .WriteStartElement("Contact1")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Email") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact1
            '    .WriteStartElement("Contact2")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Email") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact2

            '    .WriteEndElement() ' End of DespatchFrom

            '    .WriteStartElement("DeliverTo")
            '    .WriteStartElement("Code") : .WriteString("") : .WriteEndElement() ' ??
            '    .WriteStartElement("EAN13") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address2") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address5") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("PostCode") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Country") : .WriteString("") : .WriteEndElement() ' get from original order if required

            '    .WriteStartElement("VAT")
            '    .WriteStartElement("Num") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteStartElement("Alpha") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteEndElement() ' End of VAT

            '    .WriteStartElement("Contact1")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact1
            '    .WriteStartElement("Contact2")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact2

            '    .WriteEndElement() ' End of DeliverTo

            '    .WriteStartElement("Instructions")
            '    .WriteStartElement("Line1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement()

            '    .WriteStartElement("DelCombined") : .WriteString("false") : .WriteEndElement()

            '    .WriteEndElement() 'End of Delivery

            '    .WriteStartElement("Locations")

            '    .WriteStartElement("OrderBranch")
            '    .WriteStartElement("Code") : .WriteString("") : .WriteEndElement() ' ??
            '    .WriteStartElement("EAN13") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address5") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Country") : .WriteString("") : .WriteEndElement() ' get from original order if required

            '    .WriteStartElement("VAT")
            '    .WriteStartElement("Num") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteStartElement("Alpha") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteEndElement() ' End of VAT

            '    .WriteStartElement("Contact1")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact1
            '    .WriteStartElement("Contact2")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact2

            '    .WriteEndElement() 'End of OrderBranch


            '    .WriteStartElement("InvoiceTo")
            '    .WriteStartElement("Code") : .WriteString("") : .WriteEndElement() ' ??
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address5") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Country") : .WriteString("") : .WriteEndElement() ' get from original order if required

            '    .WriteStartElement("VAT")
            '    .WriteStartElement("Num") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteStartElement("Alpha") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteEndElement() ' End of VAT

            '    .WriteStartElement("Contact1")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact1
            '    .WriteStartElement("Contact2")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact2

            '    .WriteEndElement() ' End of InvoiceTo

            '    .WriteStartElement("InvoiceFrom")
            '    .WriteStartElement("Code") : .WriteString("") : .WriteEndElement() ' ??
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address2") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Address5") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("PostCode") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Country") : .WriteString("") : .WriteEndElement() ' get from original order if required

            '    .WriteStartElement("VAT")
            '    .WriteStartElement("Num") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteStartElement("Alpha") : .WriteString("") : .WriteEndElement() ' Customer VAT number  from original order if required
            '    .WriteEndElement() ' End of VAT

            '    .WriteStartElement("Contact1")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Email") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact1
            '    .WriteStartElement("Contact2")
            '    .WriteStartElement("Name") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Phone") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Email") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteStartElement("Fax") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() ' End of Contact2

            '    .WriteEndElement() ' End of InvoiceFrom

            '    .WriteEndElement() 'End of Locations

            '    .WriteStartElement("Notes")
            '    .WriteStartElement("Seq") : .WriteString("1") : .WriteEndElement()
            '    .WriteStartElement("Narrative")
            '    .WriteStartElement("Line1") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '    .WriteEndElement() 'End of Narrative
            '    .WriteEndElement() ' End of Notes

            '    .WriteStartElement("PricingDate")
            '    .WriteStartElement("Date") : .WriteString(Now.Date.ToString("yyyy-MM-dd")) : .WriteEndElement()
            '    .WriteStartElement("Time") : .WriteString("00:00:00")) : .WriteEndElement()
            '    .WriteEndElement()

            '    .WriteEndElement() ' End of AckHeader

            '    If dtLines IsNot Nothing Then
            '        For Each row In dtLines.Rows
            '            l_rowcount += 1
            '            .WriteStartElement("AckLine")

            '            .WriteStartElement("LineNo") : .WriteString(row("Line_Seq").ToString) : .WriteEndElement()
            '            .WriteStartElement("LineCode") : .WriteString(If(_Status <> "REJECTED", "ACCEPTED", _Status)) : .WriteEndElement()

            '            .WriteStartElement("OriginalRequirements")
            '            .WriteStartElement("Item")

            '            .WriteStartElement("CustItem")
            '            .WriteStartElement("OwnBrandEAN") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '            .WriteStartElement("Code") : .WriteString(row("Product_Code_Requested")) : .WriteEndElement()
            '            .WriteStartElement("SKU") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '            .WriteEndElement() 'End of CustItem

            '            .WriteStartElement("SuppItem")
            '            .WriteStartElement("Code") : .WriteString(row("Product_Code_Requested")) : .WriteEndElement()
            '            .WriteStartElement("EAN13") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '            .WriteStartElement("EAN12") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '            .WriteStartElement("DUN14") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '            .WriteStartElement("ISBN") : .WriteString("") : .WriteEndElement() ' get from original order if required
            '            .WriteStartElement("Desc1") : .WriteString("") : .WriteEndElement() ' get from original order if required

            '            .WriteEndElement() 'End of SuppItem

            '            .WriteStartElement("RetailEAN") : .WriteString("") : .WriteEndElement()
            '            .WriteStartElement("Desc1") : .WriteString(row("Product_Code_Requested")) : .WriteEndElement()
            '            .WriteEndElement() 'Item

            '            .WriteStartElement("Perishable") : .WriteString("false") : .WriteEndElement()

            '            .WriteStartElement("UnitOfOrder")
            '            .WriteStartElement("Unit") : .WriteString("1") : .WriteEndElement()
            '            .WriteStartElement("OrderMeasure") : .WriteString("0") : .WriteEndElement()
            '            .WriteEndElement() 'End of UnitOfOrder

            '            .WriteStartElement("OrderQty")
            '            .WriteStartElement("Unit") : .WriteString(row("Qty").ToString) : .WriteEndElement()
            '            .WriteStartElement("UOM") : .WriteString("EA") : .WriteEndElement()
            '            .WriteEndElement() 'End of OrderQty

            '            .WriteStartElement("PricingMeasure")
            '            .WriteStartElement("Measure") : .WriteString("EA") : .WriteEndElement()
            '            .WriteStartElement("MeasureQty") : .WriteString("1") : .WriteEndElement()
            '            .WriteEndElement() 'End of PricingMeasure

            '            '              

            '            ''.WriteString(row("Product_Code")) 

            '            .WriteEndElement() 'OriginalRequirements

            '            '.WriteStartElement("Quantity") : .WriteString(row("Qty").ToString) : .WriteEndElement() 'NOTE Qty 0 = can not supply this item and no substitude available
            '            '.WriteStartElement("UnitPrice") : .WriteString(Nz(Of Decimal)(row("Unit_Price"), 0).ToString("0.00")) : .WriteEndElement()
            '            '.WriteStartElement("LinePrice") : .WriteString((Nz(Of Decimal)(row("Unit_Price"), 0) * CType(row("Qty"), Decimal)).ToString("0.00")) : .WriteEndElement()

            '            'If row("Product_Code").ToString.Length = 0 Then
            '            '    .WriteStartElement("LineStatus") : .WriteString("OS") : .WriteEndElement()
            '            'ElseIf row("Product_Code_Requested") <> row("Product_Code") Then
            '            '    .WriteStartElement("SubstitutedProduct")
            '            '    .WriteString(row("Product_Code"))
            '            '    .WriteEndElement() 'SubstitutedProduct
            '            '    .WriteStartElement("LineStatus") : .WriteString("OOS") : .WriteEndElement()
            '            'Else
            '            '    .WriteStartElement("LineStatus") : .WriteEndElement()
            '            'End If

            '            '.WriteStartElement("UnitOfMeasure") : .WriteString(row("UOM")) : .WriteEndElement()

            '            .WriteEndElement() 'AckLine
            '        Next
            '    End If

            '    .WriteStartElement("DocTrailer")
            '    .WriteStartElement("TotalLines") : .WriteString(_Order_Lines) : .WriteEndElement()
            '    .WriteEndElement()

            '    '    .WriteEndElement() '/LineItems

            '    .WriteEndElement() ' End of <Document>

            '    .WriteStartElement("MsgTrailer")
            '    .WriteStartElement("TotalDocs") : .WriteString("1") : .WriteEndElement()
            '    .WriteEndElement()

            '    .WriteEndElement() ' End of tc:TrueCommerceOrderAck
            '    .WriteEndDocument()
            '    .Flush()
            '    .Close()

            'End With


            If UpdateAcknowledgementDate() = 0 Then
                'MoveFile( _RESPONSE_OUT & "\" & l_File, _RESPONSE_ARCHIVED & "\" & l_File)
            End If


            If pOrderResponseType = 2 Then
                'MoveFile(_RESPONSE_OUT & "\" & l_File, _RESPONSE_ARCHIVED & "\" & l_File)

                If _Status.Equals("REJECTED", StringComparison.CurrentCultureIgnoreCase) Then
                    Select Case pErrCode
                        ' 1=Customer account on stop; 2=Customer account is not recognised; 3=Delivery date changed
                        Case 6 ' 'CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST'
                            _EmailServiceMessage("REJECTED: " & _OrderNum & "   " & _AccNum & "  {CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST}", True)
                        Case 7 'DELIVERY_SLOT_MISSED'
                            _EmailServiceMessage("REJECTED: " & _OrderNum & "   " & _AccNum & "  {DELIVERY_SLOT_MISSED}  " & _DeliveryDateRequested, True)
                        Case Else
                            _EmailServiceMessage("REJECTED: " & _OrderNum & "   " & _AccNum & "  " & pErrCode & " - " & pErrMsg, True)
                    End Select
                ElseIf pOrderResponseType = 1 Then
                    If _Status.Equals("MODIFIED", StringComparison.CurrentCultureIgnoreCase) Then
                        If DateDiff(DateInterval.Day, _DeliveryDateRequested, _DeliveryDate) <> 0 Then
                            _EmailServiceMessage("MODIFIED: " & _OrderNum & "   " & _AccNum & "  {DELIVERY_DATE_CHANGED}  " & _DeliveryDateRequested & " --> " & _DeliveryDate, True)
                        Else
                            _EmailServiceMessage("MODIFIED: " & _OrderNum & "   " & _AccNum & "  {MODIFIED ORDER LINES}  requested: " & _Order_Lines.ToString & "  ordered:" & l_rowcount.ToString, True)
                        End If
                    End If
                End If
            End If

            Return True

        Catch ex As Exception
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False

    End Function

#End Region

#Region "Methods CSV - CN_CrunchTime --------------------------------------------------------------"
    Private Sub Process_CN_CrunchTime()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer

        Dim lResult As MsgBoxResult

        Try ''MyEventLog.WriteEntry("Monitoring the System", EventLogEntryType.Information, GetEventID))

            'PROCESS NEW ORDERS
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))
                MyEventLog.WriteEntry("ORDER REQUEST RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)
                lResult = UploadFile_CN_CrunchTime_CSV(_ORDER_IN & "\" & l_File)
                If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                    Return
                End If
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
    End Sub

    Private Function UploadFile_CN_CrunchTime_CSV(ByVal pFileName As String) As MsgBoxResult
        Dim lCSVContents As String = String.Empty,
            lReader As StreamReader,
            lRetVal As Integer = 0,
            lErrMsg As String = String.Empty,
            lNewOrderID As Integer = 0,
            lHeader() As String = Nothing,
            lDetail() As String = Nothing,
            lInvalidFileFormat As Boolean = False,
            lVendorLocationAccount As String,
            lPONum As String,
            lDeliveryDate As Date, lYYYYMMDD As String,
            lSpecialInstructions As String,
            lLineType As String = String.Empty,
            l_Code As String,
            l_Qty As Integer,
            dtLines As DataTable = New DataTable("DT_MEDINA_ORDER_LINES"),
            dRow As DataRow = Nothing _ '' re-use the same structure
          , lResponseMsg As String = String.Empty _
          , dtProcessed As DataTable

        Try
            With dtLines
                .Columns.Add(New DataColumn("prod_code", GetType(String)))
                .Columns.Add(New DataColumn("qty", GetType(Integer)))
                .Columns.Add(New DataColumn("unit_price", GetType(Decimal)))
            End With

            lReader = New StreamReader(pFileName)
            lCSVContents = lReader.ReadToEnd()
            lReader.Close()

            Dim lPOContent() As String = lCSVContents.Split(Constants.vbLf)

            lInvalidFileFormat = lPOContent.Count = 0

            If Not lInvalidFileFormat Then
                For idx As Integer = 0 To lPOContent.Length - 1
                    If lPOContent(idx).Trim.StartsWith("H", StringComparison.CurrentCultureIgnoreCase) Then
                        lLineType = "H"
                        dtLines.Rows.Clear()
                        lSpecialInstructions = String.Empty
                        lHeader = lPOContent(idx).Split(",")
                        If lHeader.Count > 1 Then lVendorLocationAccount = lHeader(1) Else lInvalidFileFormat = True
                        If lHeader.Count > 2 Then lPONum = lHeader(2) Else lInvalidFileFormat = True
                        If lHeader.Count > 3 Then lYYYYMMDD = lHeader(3) Else lInvalidFileFormat = True

                        If Not IsDate(lYYYYMMDD) Then lYYYYMMDD = lYYYYMMDD.Insert(4, "-").Insert(7, "-")

                        If IsDate(lYYYYMMDD) Then
                            lDeliveryDate = CType(lYYYYMMDD, Date).ToString("yyyy MMM dd")
                        Else
                            lInvalidFileFormat = True
                        End If

                        If lHeader.Count > 4 Then lSpecialInstructions = lHeader(4)

                    ElseIf lPOContent(idx).Trim.StartsWith("D", StringComparison.CurrentCultureIgnoreCase) Then
                        lLineType = "D"
                        l_Code = String.Empty : l_Qty = 0
                        lDetail = lPOContent(idx).Split(",")
                        If lDetail.Count > 1 AndAlso lVendorLocationAccount = lDetail(1) AndAlso lDetail.Count > 2 AndAlso lPONum = lDetail(2) Then
                            If lDetail.Count > 3 Then l_Code = lDetail(3)
                            If lDetail.Count > 4 Then l_Qty = lDetail(4)
                            If Not String.IsNullOrEmpty(l_Code) AndAlso l_Qty > 0 Then
                                dRow = dtLines.NewRow()
                                dRow("prod_code") = l_Code
                                dRow("qty") = l_Qty
                                dRow("unit_price") = DBNull.Value
                                dtLines.Rows.Add(dRow)
                            Else
                                'TODO - reject the whole order ?
                            End If
                        End If
                    End If
                Next

                If lLineType = "D" Then 'If D then there was at least one product
                    lRetVal = ExecCSVImportProcedure(pFileName, lPONum, lVendorLocationAccount, lDeliveryDate, Now, dtLines, lNewOrderID, lErrMsg, "p_P2P_order_import_cnuk")
                    Select Case lRetVal
                        Case 0 ' Success
                            System.Threading.Thread.Sleep(50)
                            dtProcessed = GetOrderAmendments(True)
                            If CreateOrderResponse_CN_CrunchTime(dtProcessed, lRetVal, "", lResponseMsg) Then
                                MyEventLog.WriteEntry("Order Response Created: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                                ' _EmailServiceMessage(Date.Now & " ORDER RESPONSE CREATED: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & vbNewLine & vbNewLine & lResponseMsg)
                            End If
                            EmailOrder(lNewOrderID)
                            Return MsgBoxResult.Ok
                        Case 5555
                            MyEventLog.WriteEntry("ERROR: " & lErrMsg & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
                            _EmailServiceMessage(Date.Now & " " & "ERROR: " & lErrMsg & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
                            If lErrMsg.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                                Return MsgBoxResult.Retry
                            Else
                                Return MsgBoxResult.Abort
                            End If
                        Case Else
                            _Status = "REJECTED"
                            If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "
                            MyEventLog.WriteEntry("Order REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg, EventLogEntryType.Warning, GetEventID)
                            _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg)
                            System.Threading.Thread.Sleep(50)
                            Return MsgBoxResult.Abort
                    End Select
                End If

                If Not lInvalidFileFormat Then
                    lRetVal = ExecCSVImportProcedure(pFileName, lPONum, lVendorLocationAccount, lDeliveryDate, Now, dtLines, lNewOrderID, lErrMsg, "p_P2P_order_import_cnuk")

                    Select Case lRetVal
                        Case 0 ' Success
                            If _Status = "MODIFIED_LINES" Then _Status = "MODIFIED"

                            dtProcessed = GetOrderAmendments(True)
                            If CreateOrderResponse_CN_CrunchTime(dtProcessed, lRetVal, "", lResponseMsg) Then
                                MyEventLog.WriteEntry("Order Response Created: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                                '  _EmailServiceMessage(Date.Now & " ORDER RESPONSE CREATED: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & vbNewLine & vbNewLine & lResponseMsg)
                            End If
                            System.Threading.Thread.Sleep(50)
                            EmailOrder(lNewOrderID)

                            Return MsgBoxResult.Ok
                        Case 5555
                            MyEventLog.WriteEntry("ERROR: " & lErrMsg & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
                            _EmailServiceMessage(Date.Now & " " & "ERROR: " & lErrMsg & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
                            If lErrMsg.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                                Return MsgBoxResult.Retry
                            Else
                            End If
                        Case Else
                            _Status = "REJECTED"
                            MyEventLog.WriteEntry("Order REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & " " & lErrMsg, EventLogEntryType.Warning, GetEventID)
                            'If CreateOrderResponse_Medina(2, dt, lRetVal, lErrMsg, lResponseMsg) Then
                            _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & vbNewLine & vbNewLine & lErrMsg, True)
                            System.Threading.Thread.Sleep(50)
                            'End If
                    End Select
                End If

            End If

            If lInvalidFileFormat Then
                'empty file or invalid format
                MyEventLog.WriteEntry("Invalid format " & pFileName, EventLogEntryType.Error, GetEventID)
                _EmailServiceMessage(Date.Now & " " & "ERROR: Invalid format " & pFileName, True)
                Return MsgBoxResult.Abort
            End If


        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function
    Private Function ExecCSVImportProcedure(pFileName As String, pOrderNum As String, pAccNum As String, pDeliveryDate As Date, pDateTimeCreated As String, pdtLines As DataTable, ByRef pNewOrderID As Integer, ByRef pErrMsg As String, Optional pProcedureName As String = "p_P2P_order_import_medina") As Integer
        Dim l_DB = New DB,
            param As SqlClient.SqlParameter,
            cmd As SqlClient.SqlCommand
        Try

            l_DB.Open()

            cmd = l_DB.SqlCommand(pProcedureName)

            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 200)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@order_num", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = pOrderNum

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = pAccNum

                param = .Parameters.Add("@delivery_date_requested", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                If IsDate(pDeliveryDate) Then
                    param.Value = pDeliveryDate
                Else
                    param.Value = DateAdd(DateInterval.Day, 1, Date.Today)
                End If

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@datetime_created_string", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = pDateTimeCreated

                param = .Parameters.Add("@status", SqlDbType.VarChar, 16)
                param.Direction = ParameterDirection.InputOutput
                param.Value = "ACCEPTED"

                param = .Parameters.Add("@customer_order_header_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@DT_MEDINA_ORDER_LINES", SqlDbType.Structured)
                param.Direction = ParameterDirection.Input
                param.Value = pdtLines

                param = .Parameters.Add("@order_lines", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@order_value", SqlDbType.Decimal)
                param.Precision = 10
                param.Scale = 4
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                .ExecuteNonQuery()

                _OrderRequestId = .Parameters("@record_id").Value
                _OrderNum = Nz(Of String)(.Parameters("@order_num").Value, "")
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "")
                _DeliveryDateRequested = Nz(Of Date)(.Parameters("@delivery_date_requested").Value, Date.MinValue)
                _DeliveryDate = Nz(Of Date)(.Parameters("@delivery_date").Value, Date.MinValue)
                _Status = Nz(Of String)(.Parameters("@status").Value, "")
                _Order_DateTime_Created = Nz(Of String)(.Parameters("@datetime_created_string").Value, "")
                _Order_Lines = Nz(Of Integer)(.Parameters("@order_lines").Value, 0)
                _Order_Value = Nz(Of Decimal)(.Parameters("@order_value").Value, 0)

                pNewOrderID = Nz(Of Integer)(.Parameters("@customer_order_header_id").Value, 0)

                pErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")

                Return CType(.Parameters("@ret").Value, Integer)
            End With
            cmd.Dispose()
            l_DB.Close()
        Catch ex As Exception
            pErrMsg = ex.Message
            Return 5555
        Finally

        End Try
    End Function

    Private Function CreateOrderResponse_CN_CrunchTime(dtLines As DataTable, pErrCode As Integer, pErrMsg As String, ByRef pResponse As String) As Boolean
        Dim writer As StreamWriter = Nothing _
            , l_File As String _
            , l_Total As Single = 0

        Try
            l_File = _RESPONSE_OUT & "\009-Confirm-" & _OrderNum & "-" & Format(Date.Now, "yyyyMMddHHmm") & ".txt"
            If File.Exists(l_File) Then File.Delete(l_File)
            _DeliveryNoteNum = "DN_" & _OrderNum
            With New StringBuilder
                'Header Lines
                .Append("H") : .Append(vbTab)
                .Append("WCM") : .Append(vbTab) 'Vendor Code
                .Append(_AccNum) : .Append(vbTab) 'Location Code
                .Append(_OrderNum) : .Append(vbTab) 'Purchase Order Number
                'request from Adam Tekiela <adamt@caffenero.com> : Any chance you could adjust them to be MM/DD/YYYY
                .Append(_DeliveryDate.ToString("MM/dd/yyyy")) : .Append(vbTab) 'Expected Delivery Date  // ToString("dd/MM/yyyy")
                .Append(_DeliveryNoteNum) : .Append(vbTab) 'dEDelivery NoteNumber 

                For Each row As DataRow In dtLines.Rows
                    l_Total += Math.Round(row("Qty") * row("Unit_Sale"), 2)
                Next
                .Append(_DeliveryDateRequested.ToString("MM/dd/yyyy")) : .Append(vbTab) 'Delivery Date
                .Append(Math.Round(l_Total, 2)) : .Append(vbTab) ' Total sales price
                'optional values
                .Append("TAXES PAYABLE - SALES & USE") : .Append(vbTab)
                .AppendLine("0")
                For Each row As DataRow In dtLines.Rows
                    .Append("D") : .Append(vbTab)
                    ''.Append(row("Product_Code_Requested")) : .Append(vbTab)   

                    '       Adam Tekiela | Senior Systems Analyst Caffe Nero Group Ltd commented (21.04.2023) :
                    '       "I believe in the even when substitution Is being used you can just replace the details of the item ordered with the substitution item details."
                    .Append(row("Product_Code")) : .Append(vbTab)
                    .Append("N") : .Append(vbTab)
                    .Append(Format(row("Qty"), "#")) : .Append(vbTab)
                    .Append(Math.Round(row("Unit_Sale"), 4)) : .Append(vbTab) 'Unit Sales Price
                    .Append(Math.Round(row("Qty") * row("Unit_Sale"), 2)) : .Append(vbTab) 'Invoice Total
                    .Append("0") : .Append(vbTab)   'VAT
                    If row("Product_Code_Requested") <> row("Product_Code") Then
                        .Append("Y") : .Append(vbTab) ' Substitute indicator
                        .Append(row("Product_Code")) : .Append(vbTab)
                        .Append(row("Product")) : .Append(vbTab)
                    End If
                    .AppendLine("")
                Next

                ' Open the file for writing
                writer = File.CreateText(l_File)
                writer.Write(.ToString)
            End With
            writer.Close()


            File.Copy(l_File, _RESPONSE_ARCHIVED & "\" & My.Computer.FileSystem.GetName(l_File), True)

            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False

    End Function

#End Region


#Region "Methods BOURNE--------------------------------------------------------------"

    Private Sub Process_Bourne()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer
        Dim bFailedToSendResponces As Boolean = False
        Dim lMsg As String = String.Empty
        Dim lResult As MsgBoxResult

        Try

            'PROCESS NEW ORDERS
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))
                If l_File.StartsWith("Bourne") Then
                    MyEventLog.WriteEntry("ORDER REQUEST RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)

                    If Not _Test_Mode Then
                        __LastOrderReceivedDatetime_Bourne = Date.Now
                        __LastWarningEmailDatetime_Bourne = __LastOrderReceivedDatetime_Bourne
                    End If

                    lResult = UploadFile_Bourne(_ORDER_IN & "\" & l_File)
                    If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                        Return
                    End If

                End If
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
    End Sub

    Private Function UploadFile_Bourne(ByVal pFileName As String) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand
        Dim lXMLContents As String = String.Empty
        Dim lReader As StreamReader
        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim lResponseMsg As String = String.Empty
        Dim dt As DataTable = Nothing
        Dim lNewOrderID As Integer = 0

        Try
            l_DB.Open()

            lReader = New StreamReader(pFileName)
            lXMLContents = lReader.ReadToEnd()
            lReader.Close()

            cmd = l_DB.SqlCommand("p_P2P_order_import_bourne")

            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 200)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@xml_order_as_varchar", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.Input
                param.Value = lXMLContents

                param = .Parameters.Add("@order_num", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@delivery_date_requested", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@datetime_created_string", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@status", SqlDbType.VarChar, 16)
                param.Direction = ParameterDirection.InputOutput
                param.Value = "ACCEPTED"

                param = .Parameters.Add("@customer_order_header_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@order_lines", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@order_value", SqlDbType.Decimal)
                param.Precision = 10
                param.Scale = 4
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                .ExecuteNonQuery()
                lRetVal = CType(.Parameters("@ret").Value, Integer)

                _OrderRequestId = .Parameters("@record_id").Value
                _OrderNum = Nz(Of String)(.Parameters("@order_num").Value, "")
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "")
                _DeliveryDateRequested = Nz(Of Date)(.Parameters("@delivery_date_requested").Value, Date.MinValue)
                _DeliveryDate = Nz(Of Date)(.Parameters("@delivery_date").Value, Date.MinValue)
                _Status = Nz(Of String)(.Parameters("@status").Value, "")
                _Order_DateTime_Created = Nz(Of String)(.Parameters("@datetime_created_string").Value, "")
                _Order_Lines = Nz(Of Integer)(.Parameters("@order_lines").Value, 0)
                _Order_Value = Nz(Of Decimal)(.Parameters("@order_value").Value, 0)

                lNewOrderID = Nz(Of Integer)(.Parameters("@customer_order_header_id").Value, 0)

                lErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")
            End With
            cmd.Dispose()
            l_DB.Close()

            '' dt = GetOrderLines()

            Select Case lRetVal
                Case 0 ' Success

                    If _Status.StartsWith("MODIFIED") Then
                        ''     If CreateOrderResponse_Bourne(2, dt, lRetVal, "", lResponseMsg) Then
                        MyEventLog.WriteEntry("Order Processed: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                        _EmailServiceMessage(Date.Now & " ORDER Processed: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & vbNewLine & vbNewLine & lResponseMsg)
                        ''       End If
                    End If
                    System.Threading.Thread.Sleep(50)
                    EmailOrder(lNewOrderID)
                    Return MsgBoxResult.Ok
                Case Else
                    _Status = "REJECTED"
                    If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "
                    '' If CreateOrderResponse_Bourne(2, dt, lRetVal, lErrMsg, lResponseMsg) Then
                    _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg & vbNewLine & vbNewLine & lResponseMsg)
                    System.Threading.Thread.Sleep(50)
                    ''  End If
                    MyEventLog.WriteEntry("Order REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg & vbNewLine & lResponseMsg, EventLogEntryType.Warning, GetEventID)
            End Select

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function

    '''' <summary>
    '''' 
    '''' </summary>
    '''' <param name="pOrderResponseType">1=Acknowledgement, 2=order confirmation, 3=order ASN</param>
    '''' <returns></returns>
    '''' <remarks> </remarks>
    'Private Function CreateOrderResponse_Bourne(pOrderResponseType As Integer, dtLines As DataTable, pErrCode As Integer, pErrMsg As String, ByRef pResponse As String) As Boolean
    '    Dim xmlWriter As XmlWriter = Nothing
    '    Dim dtProcessDate As Date = Now
    '    Dim l_Directory As String = ""
    '    Dim l_File_UPL As String = "", l_File_CON As String = ""
    '    ' Dim row As DataRow

    '    Try
    '        'Select Case pOrderResponseType
    '        '    Case 2
    '        '        l_File_UPL = "UPL--" & _OrderGuid & "--" & _Order_DateTime_Created & ".XML"
    '        '        l_File_CON = "CON--" & _OrderGuid & "--" & _Order_DateTime_Created & ".XML"
    '        '        ''Case 3 :
    '        'End Select

    '        'If File.Exists(_RESPONSE_OUT & "\" & l_File_UPL) Then File.Delete(_RESPONSE_OUT & "\" & l_File_UPL)
    '        'If File.Exists(_RESPONSE_OUT & "\" & l_File_CON) Then File.Delete(_RESPONSE_OUT & "\" & l_File_CON)

    '        'xmlWriter = New XmlTextWriter(_RESPONSE_OUT & "\" & l_File_UPL, System.Text.Encoding.UTF8)

    '        'If IsNothing(xmlWriter) Then
    '        '    Return False
    '        'End If

    '        'With xmlWriter
    '        '    .WriteStartDocument()
    '        '    .WriteStartElement("orderConfirmation")

    '        '    .WriteStartElement("header")

    '        '    .WriteStartElement("testStatus")
    '        '    .WriteString(If(My.Settings.TestCypad, "Y", "N"))
    '        '    .WriteEndElement()

    '        '    .WriteStartElement("purchaseOrderReference") : .WriteString(_OrderNum) : .WriteEndElement()
    '        '    .WriteStartElement("purchaseOrderDate") : .WriteString(_Order_DateTime_Created.Replace("-", "")) : .WriteEndElement()

    '        '    .WriteStartElement("orderStatus") : .WriteString(_Status) : .WriteEndElement() ' ACCEPTED, REJECTED, MODIFIED

    '        '    .WriteStartElement("orderStatusReason")
    '        If _Status.Equals("REJECTED", StringComparison.CurrentCultureIgnoreCase) Then
    '            Select Case pErrCode
    '                ' 1=Customer account on stop; 2=Customer account is not recognised; 3=Delivery date changed
    '                Case 4
    '                    pResponse &= "DUPLICATE_ORDER" ' added by VS on 13/07/2020
    '                Case 6 ' 'CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST'
    '                    pResponse &= "CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST" & vbNewLine
    '                Case 7 'DELIVERY_SLOT_MISSED'
    '                    pResponse &= "DELIVERY_SLOT_MISSED" & vbNewLine
    '                Case 8 'PRODUCT_NOT_VALID_FOR_LOCATION'
    '                    pResponse &= "PRODUCT_NOT_VALID_FOR_LOCATION" & vbNewLine
    '                Case 9 'DELIVERY_SLOT_UNCLEAR
    '                    pResponse &= "DELIVERY_SLOT_UNCLEAR: RING CUSTOMER TO CHECK REQUIRED DELIVERY DATE" & vbNewLine
    '            End Select
    '        ElseIf _Status.Equals("MODIFIED", StringComparison.CurrentCultureIgnoreCase) Then
    '            If DateDiff(DateInterval.Day, _DeliveryDateRequested, _DeliveryDate) <> 0 Then
    '                pResponse &= "DELIVERY_DATE_CHANGED" & vbNewLine
    '                pResponse &= "Confirmed Delivery Date :"
    '                If IsDate(_DeliveryDate) AndAlso _DeliveryDate <> Date.MinValue Then
    '                    pResponse &= (CType(_DeliveryDate, Date).ToString("ddd dd/MM/yyyy")) & vbNewLine
    '                ElseIf _DeliveryDate <> Date.MinValue Then
    '                    pResponse &= _DeliveryDate & vbNewLine
    '                End If
    '            End If
    '        ElseIf _Status.Equals("MODIFIED_LINES", StringComparison.CurrentCultureIgnoreCase) Then
    '            pResponse &= "ORDER ITEMS MODIFIED" & vbNewLine
    '        End If

    '        ' ''in modified order confirmation the only modified (or rejected) items get included 
    '        ''If dtLines IsNot Nothing Then
    '        ''    pResponse &= "Order Items" & vbNewLine
    '        ''    Dim lItemStatus As String, lReason As String
    '        ''    For Each row In dtLines.Rows
    '        ''        If row("Product_Code_Requested") <> row("Product_Code") Then
    '        ''            .WriteStartElement("item")

    '        ''            .WriteStartElement("itemCode")
    '        ''            .WriteString(row("Product_Code_Requested")) ' the same as Product_Code is there is no substitution 
    '        ''            .WriteEndElement() ' itemCode

    '        ''            .WriteStartElement("itemQuantity") : .WriteString(row("Qty").ToString) : .WriteEndElement() 'NOTE Qty 0 = can not supply this item and no substitude available
    '        ''            .WriteStartElement("itemPrice") : .WriteString(Nz(Of Decimal)(row("Unit_Price"), 0).ToString("0.00")) : .WriteEndElement()

    '        ''            If row("Qty") = 0 Then
    '        ''                lItemStatus = "R" 'Rejected 
    '        ''                lReason = "PRODUCT_NOT_VALID_FOR_LOCATION"
    '        ''            Else
    '        ''                lItemStatus = "C" 'Changed
    '        ''                lReason = "SUBSTITUTED"
    '        ''            End If

    '        ''            .WriteStartElement("itemStatus") : .WriteString(lItemStatus) : .WriteEndElement()
    '        ''            .WriteStartElement("itemReasonForChange") : .WriteString(lReason) : .WriteEndElement()

    '        ''            If row("Product_Code").ToString.Length > 0 Then
    '        ''                .WriteStartElement("itemSubstitute")
    '        ''                .WriteString(row("Product_Code"))
    '        ''                .WriteEndElement() 'SubstitutedProduct
    '        ''            End If

    '        ''            .WriteEndElement() '/item
    '        ''        End If
    '        ''    Next
    '        ''End If
    '        ''End If


    '        'If pOrderResponseType = 2 Then
    '        '    If _OrderRequestId <> 0 Then UpdateAcknowledgementDate()
    '        'My.Computer.FileSystem.RenameFile(_RESPONSE_OUT & "\" & l_File_UPL, l_File_CON)
    '        'File.Copy(_RESPONSE_OUT & "\" & l_File_CON, _RESPONSE_ARCHIVED & "\" & l_File_CON, True)
    '        'End If
    '        Return True

    '    Catch ex As Exception
    '        MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
    '        _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
    '    End Try
    '    Return False

    'End Function

#End Region

#Region "Methods INTERSERVE SAFFRON --------------------------------------------------------------"

    Private Sub Process_Interserve()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer
        Dim bFailedToSendResponces As Boolean = False
        Dim lMsg As String = String.Empty
        Dim lResult As MsgBoxResult
        Dim lDoProcess As Boolean = False

        Try

            'PROCESS NEW ORDERS
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))

                If l_File.StartsWith("PO_WCM") Then lDoProcess = True

                If lDoProcess Then
                    MyEventLog.WriteEntry("ORDER REQUEST RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)

                    If Not _Test_Mode Then
                        __LastOrderReceivedDatetime_Interserve = Date.Now
                        __LastWarningEmailDatetime_Interserve = __LastOrderReceivedDatetime_Interserve
                    End If

                    lResult = UploadFile_Interserve(_ORDER_IN & "\" & l_File)
                    If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                        Return
                    End If
                End If
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
    End Sub

    Private Function UploadFile_Interserve(ByVal pFileName As String) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand
        Dim lXMLContents As String = String.Empty
        Dim lReader As StreamReader
        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim lResponseMsg As String = String.Empty
        Dim dt As DataTable = Nothing
        Dim lCustomeLocationCode As String = String.Empty
        Dim lUnitCode As String = String.Empty
        Dim lArgeementCode As String = String.Empty
        Dim lNewOrderID As Integer = 0
        Dim lCustomerCode As String = ""
        Dim lDNPrefix As String = ""

        Try
            l_DB.Open()

            lReader = New StreamReader(pFileName)
            lXMLContents = lReader.ReadToEnd()
            lXMLContents = lXMLContents.Replace(" encoding=""iso-8859-1""", "")
            lReader.Close()

            cmd = l_DB.SqlCommand("p_P2P_order_import_Interserve")

            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 200)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@xml_order", SqlDbType.Xml)
                param.Direction = ParameterDirection.Input
                param.Value = lXMLContents

                param = .Parameters.Add("@order_num", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@delivery_date_requested", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@datetime_created_string", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@customer_location_code", SqlDbType.VarChar, 50)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@unit_code", SqlDbType.VarChar, 50)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@status", SqlDbType.VarChar, 16)
                param.Direction = ParameterDirection.InputOutput
                param.Value = "ACCEPTED"

                param = .Parameters.Add("@customer_order_header_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@customer_agreement_code", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@customer_code", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                .ExecuteNonQuery()
                lRetVal = CType(.Parameters("@ret").Value, Integer)

                _OrderRequestId = .Parameters("@record_id").Value
                _OrderNum = Nz(Of String)(.Parameters("@order_num").Value, "")
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "")
                _DeliveryDateRequested = Nz(Of Date)(.Parameters("@delivery_date_requested").Value, Date.MinValue)
                _DeliveryDate = Nz(Of Date)(.Parameters("@delivery_date").Value, Date.MinValue)
                _Status = Nz(Of String)(.Parameters("@status").Value, "")
                _Order_DateTime_Created = Nz(Of String)(.Parameters("@datetime_created_string").Value, "")

                lArgeementCode = Nz(Of String)(.Parameters("@customer_agreement_code").Value, "")
                lCustomeLocationCode = Nz(Of String)(.Parameters("@customer_location_code").Value, "")
                lUnitCode = Nz(Of String)(.Parameters("@unit_code").Value, "")
                lNewOrderID = Nz(Of Integer)(.Parameters("@customer_order_header_id").Value, 0)

                lErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")
                lCustomerCode = Nz(Of String)(.Parameters("@customer_code").Value, "5060642190021")
            End With

            cmd.Dispose()
            l_DB.Close()

            dt = GetOrderLines()

            Select Case lRetVal
                Case 0 ' Success

                    If CreateASN_Saffron(2, lCustomeLocationCode, lUnitCode, lArgeementCode, dt, lRetVal, "", lResponseMsg, lCustomerCode) Then
                        If dt.Rows.Count > 0 Then
                            'create additional  ASN for dual supply order
                            CreateASN_Saffron(2, lCustomeLocationCode, lUnitCode, lArgeementCode, dt, lRetVal, "", lResponseMsg, lCustomerCode)
                        End If
                        MyEventLog.WriteEntry("Order Processed: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                        If _Status.StartsWith("MODIFIED") Then
                            _EmailServiceMessage(Date.Now & " ORDER Processed: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & vbNewLine & vbNewLine & lResponseMsg)
                        End If
                    End If
                    System.Threading.Thread.Sleep(50)
                    EmailOrder(lNewOrderID)
                    Return MsgBoxResult.Ok
                Case Else
                    _Status = "REJECTED"
                    If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "
                    If CreateASN_Saffron(2, lCustomeLocationCode, lUnitCode, lArgeementCode, dt, lRetVal, lErrMsg, lResponseMsg, lCustomerCode) Then
                        If dt.Rows.Count > 0 Then
                            'create additional  ASN for dual supply order
                            CreateASN_Saffron(2, lCustomeLocationCode, lUnitCode, lArgeementCode, dt, lRetVal, "", lResponseMsg, lCustomerCode)
                        End If
                        _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg & vbNewLine & vbNewLine & lResponseMsg)
                        System.Threading.Thread.Sleep(50)
                    End If
                    MyEventLog.WriteEntry("Order REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg, EventLogEntryType.Warning, GetEventID)
            End Select

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function
    Private Function Process_Saffron_ASN() As MsgBoxResult
        Return Process_Saffron_ASN(Date.Today)
    End Function
    ''' <summary>
    ''' 'process  ASN for Standing Orders (sub-buying group Debra, System generated Daily SO from  standing orders for these site, WCM needs to generate and supplly ASN to Saffron to match order/delivery note numbers on invoices.=))
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Private Function Process_Saffron_ASN(pDate As Date) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand

        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim lResponseMsg As String = String.Empty
        Dim dt_ASN As DataTable = New DataTable("ASN")
        Dim dt_Lines As DataTable = Nothing
        Dim lCustomerLocationCode As String = String.Empty
        Dim lUnitCode As String = String.Empty
        Dim lArgeementCode As String = String.Empty
        Dim lAmendmentID As Integer = 0
        Dim lSOHID As Integer = 0
        Dim lCustomerCode As String = ""
        Dim lDNPrefix As String = ""

        Try
            _DeliveryDate = pDate
            l_DB.Open()

            cmd = l_DB.SqlCommand("p_SO_ASN_get_list")
            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.Input
                param.Value = _DeliveryDate

                param = .Parameters.Add("@buying_group_id", SqlDbType.Int)
                param.Direction = ParameterDirection.Input
                param.Value = 214 ' ISS

                param = .Parameters.Add("@sub_buying_group_id", SqlDbType.Int)
                param.Direction = ParameterDirection.Input
                param.Value = 952 ' Defra

                l_DB.Fill(cmd, dt_ASN)
                cmd.Dispose()

                _Status = "ACCEPTED"

                _DeliveryDateRequested = _DeliveryDate
                _OrderRequestId = 0
                l_DB.Close()

                For Each row As DataRow In dt_ASN.Rows
                    lAmendmentID = row("amendment_id")
                    lSOHID = row("SOH_id")
                    _OrderNum = row("order_num")
                    _AccNum = row("acc_num")
                    _Order_DateTime_Created = row("date_created")
                    lCustomerLocationCode = row("customer_location_code")
                    lUnitCode = row("customer_unit_code")
                    lArgeementCode = row("agreement_code")
                    lCustomerCode = row("customer_code")

                    If lSOHID <> 0 Then
                        dt_Lines = GetSOLines(lSOHID, _DeliveryDate)
                    Else
                        dt_Lines = GetSOAmendmentLines(lAmendmentID)
                    End If
                    If CreateASN_Saffron(2, lCustomerLocationCode, lUnitCode, lArgeementCode, dt_Lines, lRetVal, "", lResponseMsg, lCustomerCode) Then
                        If dt_Lines.Rows.Count > 0 Then
                            'create additional  ASN for dual supply order
                            CreateASN_Saffron(2, lCustomerLocationCode, lUnitCode, lArgeementCode, dt_Lines, lRetVal, "", lResponseMsg, lCustomerCode)
                        End If
                        MyEventLog.WriteEntry("ASN Processed: " & _Status & " " & _OrderNum & " {SO} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                    End If
                    System.Threading.Thread.Sleep(50)
                Next

            End With




            '_Status = "REJECTED"
            'If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "
            'If CreateASN_Saffron(2, lCustomerLocationCode, lUnitCode, lArgeementCode, dt, lRetVal, lErrMsg, lResponseMsg, lCustomerCode) Then
            '    If dt.Rows.Count > 0 Then
            '        'create additional  ASN for dual supply order
            '        CreateASN_Saffron(2, lCustomerLocationCode, lUnitCode, lArgeementCode, dt, lRetVal, "", lResponseMsg, lCustomerCode)
            '    End If
            '    _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg & vbNewLine & vbNewLine & lResponseMsg)
            '    System.Threading.Thread.Sleep(50)
            'End If
            'MyEventLog.WriteEntry("Order REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg, EventLogEntryType.Warning, GetEventID)


        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pOrderResponseType">1=Acknowledgement, 2=order ASN</param>
    ''' <returns></returns>
    ''' <remarks> </remarks>
    Private Function CreateASN_Saffron(pOrderResponseType As Integer, pCustomerLocationCode As String, pUnitCode As String, pArgeementCode As String, ByRef dtLines As DataTable, pErrCode As Integer _
            , pErrMsg As String, ByRef pResponse As String, ByVal pCustomerCode As String) As Boolean

        Dim xmlWriter As XmlWriter = Nothing
        Dim dtProcessDate As Date = Now
        Dim l_File_UPL As String = "", l_File_CON As String = ""
        Dim row As DataRow
        Dim l_DN As String = ""
        Dim idx As Integer

        Try
            If dtLines IsNot Nothing Then
                If dtLines.Rows.Count > 0 Then
                    l_DN = dtLines.Rows(0)("del_note_num")
                End If
            End If

            Select Case pOrderResponseType
                Case 2
                    l_File_UPL = "ASN-West Country Milk-131-852091-" & pCustomerCode & "-UPL_" & l_DN & ".XML"
                    l_File_CON = "ASN-West Country Milk-131-852091-" & pCustomerCode & "-CON_" & l_DN & ".XML"
                    ''Case 3 :
            End Select

            If File.Exists(_RESPONSE_OUT & "\" & l_File_UPL) Then File.Delete(_RESPONSE_OUT & "\" & l_File_UPL)
            If File.Exists(_RESPONSE_OUT & "\" & l_File_CON) Then File.Delete(_RESPONSE_OUT & "\" & l_File_CON)

            xmlWriter = New XmlTextWriter(_RESPONSE_OUT & "\" & l_File_UPL, System.Text.Encoding.GetEncoding("ISO-8859-1"))

            If IsNothing(xmlWriter) Then
                Return False
            End If

            With xmlWriter
                .WriteStartDocument()
                .WriteStartElement("FDHASNS")
                .WriteStartElement("SupplierCode") : .WriteString("131-852091") : .WriteEndElement()
                .WriteStartElement("ASN")

                .WriteStartElement("CustomerCode")
                '' .WriteString(My.Settings.GLN_INTERSERVE) ''("5060642190021") 'Changed from 5060075570247 to this one in Feb 2020
                .WriteString(pCustomerCode) ''changed 30/06/2021 : Standard (Interserve) = 5060642190021, ISS Private = 5060642190137 , ISS Public  = 5060642190083, Epsom & St Helier University Hospitals Trust = 5060642190465, 
                'CAL1 -CalMac Ferries - 5060642190267, NCL100 - Newcastle University -5060642190274
                'MITIE1, 2, 3, 4, 5, 9  506000487988 - THEY ARE GOING TO CH&Co (Gather & Gather bought by CH&Co) IT'S NOT DECIDED YET
                .WriteEndElement()

                .WriteStartElement("UnitCode") : .WriteString(pUnitCode) : .WriteEndElement()

                .WriteStartElement("UnifiedSupplierCode") : .WriteString("131-852091") : .WriteEndElement()
                .WriteStartElement("SupplierCode") : .WriteString("131-852091") : .WriteEndElement()
                .WriteStartElement("DeliveryLocationEANCode") : .WriteEndElement()


                .WriteStartElement("CustomersLocationCode") : .WriteString(pCustomerLocationCode) : .WriteEndElement()
                .WriteStartElement("SuppliersLocationCode") : .WriteString(_AccNum) : .WriteEndElement()

                .WriteStartElement("LocationName") : .WriteEndElement()
                .WriteStartElement("LocationAddress1") : .WriteEndElement()
                .WriteStartElement("LocationAddress2") : .WriteEndElement()
                .WriteStartElement("LocationAddress3") : .WriteEndElement()
                .WriteStartElement("LocationAddress4") : .WriteEndElement()
                .WriteStartElement("LocationPostalCode") : .WriteEndElement()

                .WriteStartElement("CustomersOrderNumber") : .WriteString(_OrderNum) : .WriteEndElement()

                .WriteStartElement("CustomersOrderDate") : .WriteString(_Order_DateTime_Created.Replace("-", "")) : .WriteEndElement()

                .WriteStartElement("SuppliersOrderNumber") : .WriteString(_OrderNum) : .WriteEndElement()

                .WriteStartElement("DateOrderReceivedBySupplier") : .WriteString(Date.Today.ToString("yyyyMMdd")) : .WriteEndElement()

                '' .WriteStartElement("DeliveryNoteNumber") : .WriteString("DN_" & _OrderNum) : .WriteEndElement()
                .WriteStartElement("DeliveryNoteNumber") : .WriteString(l_DN) : .WriteEndElement() 'replaced above line by VS on 31/05/2023
                .WriteStartElement("DeliveryNoteDate") : .WriteString(_DeliveryDate) : .WriteEndElement()

                .WriteStartElement("ExpectedDeliveryDateTime")

                If _Status.Equals("REJECTED", StringComparison.CurrentCultureIgnoreCase) Then
                    Select Case pErrCode
                        Case 4
                            pResponse &= "DUPLICATE_ORDER" & vbNewLine
                            '1=Customer account on stop; 2=Customer account is not recognised; 3=Delivery date changed
                        Case 6 ' 'CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST'
                            pResponse &= "CUSTOMER_IDENTIFICATION_NUMBER_DOES_NOT_EXIST" & vbNewLine
                        Case 7 'DELIVERY_SLOT_MISSED'
                            pResponse &= "DELIVERY_SLOT_MISSED" & vbNewLine
                        Case 8 'PRODUCT_NOT_VALID_FOR_LOCATION'
                            pResponse &= "PRODUCT_NOT_VALID_FOR_LOCATION" & vbNewLine
                    End Select
                ElseIf _Status.StartsWith("MODIFIED", StringComparison.CurrentCultureIgnoreCase) Then
                    If DateDiff(DateInterval.Day, _DeliveryDateRequested, _DeliveryDate) <> 0 Then
                        pResponse &= "DELIVERY_DATE_CHANGED" & vbNewLine
                        pResponse &= "Confirmed Delivery Date :"
                        If IsDate(_DeliveryDate) AndAlso _DeliveryDate <> Date.MinValue Then
                            pResponse &= (CType(_DeliveryDate, Date).ToString("ddd dd/MM/yyyy")) & vbNewLine
                        ElseIf _DeliveryDate <> Date.MinValue Then
                            pResponse &= _DeliveryDate & vbNewLine
                        End If
                    End If
                    If _Status.Equals("MODIFIED_LINES", StringComparison.CurrentCultureIgnoreCase) Then
                        pResponse &= "ORDER ITEMS MODIFIED" & vbNewLine
                    End If
                End If
                .WriteString(CType(_DeliveryDate, Date).ToString("yyyyMMdd"))
                .WriteEndElement() 'ExpectedDeliveryDateTime

                .WriteStartElement("ASNAgreementCode") : .WriteString(pArgeementCode) : .WriteEndElement()

                .WriteStartElement("ItemList")
                'in modified order confirmation the only modified (or rejected) items get included 
                If dtLines IsNot Nothing Then
                    pResponse &= "Order Items" & vbNewLine
                    For Each row In dtLines.Rows
                        If row("del_note_num") = l_DN Then
                            If row("Product_Code_Requested") <> row("Product_Code") Then
                                If Not String.IsNullOrEmpty(row("Product_Code_Requested")) Then
                                    ' added delivery charge generates line with blank "ordered" code - do not output this line
                                    .WriteStartElement("Item")
                                    .WriteStartElement("ItemEANCode") : .WriteString("") : .WriteEndElement()
                                    .WriteStartElement("SuppliersProductCode") : .WriteString(row("Product_Code_Requested")) : .WriteEndElement() ' SuppliersProductCode
                                    .WriteStartElement("Quantity") : .WriteString("0") : .WriteEndElement() 'NOTE Qty 0 = can not supply this item and no substitude available
                                    .WriteStartElement("LineAgreementCode") : .WriteString(pArgeementCode) : .WriteEndElement()
                                    .WriteStartElement("ItemDescription") : .WriteString(row("Product_Requested")) : .WriteEndElement()
                                    .WriteEndElement() '/item
                                End If
                            End If
                            'Output ordered product (or substitude)
                            If row("Qty") <> 0 Then
                                .WriteStartElement("Item")
                                .WriteStartElement("ItemEANCode") : .WriteString("") : .WriteEndElement()

                                .WriteStartElement("SuppliersProductCode") : .WriteString(row("Product_Code")) : .WriteEndElement() ' SuppliersProductCode
                                .WriteStartElement("Quantity") : .WriteString(row("Qty").ToString) : .WriteEndElement()
                                .WriteStartElement("LineAgreementCode") : .WriteString(pArgeementCode) : .WriteEndElement()
                                .WriteStartElement("ItemDescription") : .WriteString(row("Product")) : .WriteEndElement()
                                .WriteEndElement() '/item
                            End If
                        End If
                    Next
                    For idx = dtLines.Rows.Count - 1 To 0 Step -1
                        If dtLines.Rows(idx)("del_note_num") = l_DN Then
                            dtLines.Rows.RemoveAt(idx)
                        End If
                    Next
                End If

                .WriteEndElement() 'ItemList

                .WriteEndElement() 'ASN
                .WriteEndElement() 'FDHASNS
                .WriteEndDocument()
                .Flush()
                .Close()
            End With


            If pOrderResponseType = 2 Then
                If _OrderRequestId <> 0 Then UpdateAcknowledgementDate()

                My.Computer.FileSystem.RenameFile(_RESPONSE_OUT & "\" & l_File_UPL, l_File_CON)
                File.Copy(_RESPONSE_OUT & "\" & l_File_CON, _RESPONSE_ARCHIVED & "\" & l_File_CON, True)
            End If
            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False

    End Function

#End Region

#Region "Methods Delivery Notes Grahams--------------------------------------------------------------"

    Private Sub Process_DN_Grahams()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer
        Dim lMsg As String = String.Empty
        Dim lResult As MsgBoxResult

        Try

            'PROCESS NEW EXCEL DELIVERY NOTES
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))
                If l_File.StartsWith("deliverynote-") Then
                    MyEventLog.WriteEntry("DELIVERY NOTE RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)
                    _DeliveryNoteNum = String.Empty
                    lResult = UploadFile_DN_Grahams(_ORDER_IN & "\" & l_File)
                    If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                        Return
                    End If
                End If
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
    End Sub

    Private Function UploadFile_DN_Grahams(ByVal pFileName As String) As MsgBoxResult
        Dim l_DB = New DB
        Dim dt = New DataTable("DT_P2P_DELIVERY")
        Dim dRow As DataRow
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand
        Dim lCSVContents As String = String.Empty
        Dim lLines() As String
        Dim lValues() As String
        Dim lReader As StreamReader
        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim lResponseMsg As String = String.Empty
        Dim lNewOrderID As Integer = 0

        Try

            With dt.Columns
                .Add(New DataColumn("cust_num", GetType(String)))
                .Add(New DataColumn("delivery_num", GetType(String)))
                .Add(New DataColumn("order_num", GetType(String)))
                .Add(New DataColumn("prod_code", GetType(String)))
                .Add(New DataColumn("prod_desc", GetType(String)))
                .Add(New DataColumn("qty_ordered", GetType(Int32)))
                .Add(New DataColumn("qty_delivered", GetType(Decimal)))
                .Add(New DataColumn("delivery_datetime", GetType(Date)))
                .Add(New DataColumn("temperature", GetType(Decimal)))
                .Add(New DataColumn("delivery_instructions", GetType(String)))
                .Add(New DataColumn("failure_reason", GetType(String)))
            End With

            lReader = New StreamReader(pFileName)
            lCSVContents = lReader.ReadToEnd()
            lReader.Close()

            lLines = lCSVContents.Split(vbCr)

            For row As Integer = lLines.GetLowerBound(0) + 1 To lLines.GetUpperBound(0)
                lValues = lLines(row).Split(",")
                dRow = dt.NewRow()
                For col = lValues.GetLowerBound(0) To lValues.GetUpperBound(0)
                    Select Case col
                        Case 2
                            dRow(col) = Nz(Of String)(lValues(col), String.Empty)
                            _DeliveryNoteNum = dRow(col)
                        Case 5, 6, 8
                            If IsNumeric(lValues(col)) Then
                                dRow(col) = lValues(col)
                            Else
                                dRow(col) = 0
                            End If
                        Case 7
                            If IsDate(lValues(col)) Then
                                dRow(col) = lValues(col)
                            End If
                        Case Else
                            dRow(col) = Nz(Of String)(lValues(col), String.Empty)
                    End Select
                Next col
                dt.Rows.Add(dRow)
            Next row
            If dt.Rows.Count = 0 Then
                Return MsgBoxResult.Abort
            End If

            l_DB.Open()
            cmd = l_DB.SqlCommand("p_P2P_delivery_import")

            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 200)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@supplier_id", SqlDbType.Int)
                param.Direction = ParameterDirection.Input
                param.Value = My.Settings.Grahams_SUPPL_ID

                param = .Parameters.Add(New SqlClient.SqlParameter("@DT_P2P_DELIVERY", SqlDbType.Structured))
                param.Value = dt

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                .ExecuteNonQuery()
                lRetVal = CType(.Parameters("@ret").Value, Integer)

                _OrderRequestId = .Parameters("@record_id").Value
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "")

                lErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")

                cmd.Dispose()
                l_DB.Close()

                Select Case lRetVal
                    Case 0 ' Success
                        MyEventLog.WriteEntry("Delivery Note Processed: " & _DeliveryNoteNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                        System.Threading.Thread.Sleep(50)
                        Return MsgBoxResult.Ok
                    Case Else
                        MyEventLog.WriteEntry("Delivery Note FAILED: " & _DeliveryNoteNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Warning, GetEventID)
                        _EmailServiceMessage(Date.Now & " Delivery Note FAILED: " & _DeliveryNoteNum & " {" & _OrderRequestId & "} for " & _AccNum & vbNewLine & vbNewLine & lResponseMsg)
                        System.Threading.Thread.Sleep(50)
                End Select
            End With

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function

#End Region


#Region "Orders Summary for Customers:  JJWilson --------------------------------------------------------------"

    Private Sub Process_JJWison()

        If DateDiff(DateInterval.Day, _JJWilson_done, Now.Date) = 0 Then Return

        If Date.Now.DayOfWeek <> DayOfWeek.Sunday Then
            If Date.Now.Hour = 16 AndAlso (Date.Now.Minute > 1 AndAlso Date.Now.Minute < 5) Then
                If GetSetting_JJWilson() Then
                    Dim l_DateShift As Integer = 1
                    Dim l_File As String = "JJWilson_" & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd") & ".csv"
                    Dim l_Body As String = "Please find attached orders placed for delivery on " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("dd MMM yyyy")
                    If Date.Now.DayOfWeek = DayOfWeek.Saturday Then l_DateShift = 2 'We only need to run this once a day 

                    If CreateOrderSummary(_ORDER_OUT & "\" & l_File, l_DateShift) Then

                        If Email_Generic(My.Settings.JJWilson_Email, "", "Westcountry Milk Orders for " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("dd MMM yyyy") & "  [EDI] [W01]",
                                      l_Body, _ORDER_OUT & "\" & l_File, "NoReply@wcmilk.co.uk", False, "") Then
                            _JJWilson_done = Now.Date
                        End If
                    End If

                End If
            End If
        End If
    End Sub

    Private Function CreateOrderSummary(ByVal pFileName As String, ByRef pDateShift As Integer) As Boolean
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand = Nothing
        Dim lErrMsg As String = String.Empty
        Dim dt As DataTable = New DataTable("orders")

        Try
            l_DB.Open()
            cmd = l_DB.SqlCommand("p_Export_Next_Days_Orders_by_Customer")
            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                param = .Parameters.Add("@HO_ID", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.Input
                param.Value = 50793649 ''JJWSL1 JJ Wilson Shops Ltd

                param = .Parameters.Add("@ShiftDate", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.Input
                param.Value = pDateShift

                l_DB.Fill(cmd, dt)
            End With

            cmd.Dispose()
            l_DB.Close()

            System.Threading.Thread.Sleep(20)

            With (New StreamWriter(pFileName))
                .WriteLine("AccNum,SiteName,DeliveryDate,OrderNo,Code,Product,qty,notes")
                For Each row In dt.Rows
                    .WriteLine(row(0) & "," & row(1) & "," & row(2) & "," & row(3) & "," & row(4) & "," & row(5) & "," & row(6) & "," & row(7))
                Next
                .Close()
                .Dispose()
            End With

            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then

            End If
        Finally

        End Try
        Return False
    End Function
#End Region
#Region "Methods DairyData Orders CSV : AllanReeder, FreshPastures,  JNDairies, DHT, Johal, Paynes, Broadland, Chew Valley(transmission not completed), Medina --------------------------------------------------------------"

    Private Sub Process_AllanReeder()

        If DateDiff(DateInterval.Day, _AR_done, Now.Date) = 0 Then Return

        If Date.Now.DayOfWeek <> DayOfWeek.Sunday Then
            'changed from 16:01-16:05 to 15:11-15:15;  03/02/2020 15:31- 15:35
            If Date.Now.Hour = 15 AndAlso (Date.Now.Minute > 31 AndAlso Date.Now.Minute <= 35) Then
                If GetSetting_AllanReeder() Then
                    'We only need to run this once a day at the depots cut off
                    If Export_DairyData("AllanReeder_", False) Then
                        If Export_DairyData("AllanReeder_Summary_", False, True) Then
                            Dim l_DateShift As Integer = 1
                            If Date.Now.DayOfWeek = DayOfWeek.Saturday Then l_DateShift = 2
                            Dim l_File As String = "AllanReeder_Summary_" & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd") & ".csv"
                            Dim l_Body As String = "Please find attached a summary of orders placed for delivery on " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd")

                            Email_Generic(My.Settings.AllanReeder_Summary_Email, "", "Allan Reeder EDI Summary for " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd"), l_Body, _ORDER_ARCHIVED & "\" & l_File, "NoReply@wcmilk.co.uk", False, "")
                        End If
                        _AR_done = Now.Date
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Process_DairyData_MillsMilk()
        If DateDiff(DateInterval.Day, _DairyData_MM_done, Now.Date) = 0 Then Return

        If Date.Now.DayOfWeek <> DayOfWeek.Sunday Then
            If Date.Now.Hour = 12 AndAlso (Date.Now.Minute > 31 AndAlso Date.Now.Minute < 35) Then
                If GetSetting_DairyData_MillsMilk() Then
                    'We only need to run this once a day at the depots cut off
                    If Export_DairyData("MillsMilk_", False) Then 'create a file for the next day 
                        _DairyData_MM_done = Now.Date
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Process_DairyData_Paynes()
        Dim lMinute As Integer = 0

        If DateDiff(DateInterval.Day, _DairyData_Paynes_done, Now.Date) = 0 Then Return

        Select Case Date.Now.DayOfWeek
            Case DayOfWeek.Sunday
                Return
            Case DayOfWeek.Saturday
            Case Else
                lMinute = 55
        End Select

        If Date.Now.Hour = 11 AndAlso (Date.Now.Minute >= lMinute AndAlso Date.Now.Minute <= lMinute + 4) Then
            If GetSetting_DairyData_Paynes() Then
                'We only need to run this once a day at the depots cut off
                If Export_DairyData("Paynes_") Then
                    _DairyData_Paynes_done = Now.Date
                End If
            End If
        End If

    End Sub

    Private Sub Process_DairyData_Broadland()
        If DateDiff(DateInterval.Day, _DairyData_Broadland_done, Now.Date) = 0 Then Return

        If Date.Now.Hour = 12 AndAlso (Date.Now.Minute >= 0 AndAlso Date.Now.Minute <= 5) Then
            If GetSetting_DairyData_Broadland() Then
                'We only need to run this once a day at the depots cut off
                If Export_DairyData("Broadland_") Then
                    _DairyData_Broadland_done = Now.Date

                    Dim l_DateShift As Integer = 1
                    If Date.Now.DayOfWeek = DayOfWeek.Saturday Then l_DateShift = 2
                    Dim l_File As String = "Broadland_" & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd") & ".csv"
                    Dim l_Body As String = "Please find attached orders placed for delivery on " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd")

                    Email_Generic("orders@broadlandfoodservice.co.uk", "", "Westcountry Milk Order for " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd"), l_Body, _ORDER_OUT & "\" & l_File, "NoReply@wcmilk.co.uk", True, "")
                    '' Email_Generic("chris@wcmilk.co.uk", "victor@wcmilk.co.uk", "Westcountry Milk Order for " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd"), l_Body, _ORDER_OUT & "\" & l_File, "NoReply@wcmilk.co.uk", True, "")
                    'System.Threading.Thread.Sleep(50)
                    'If File.Exists(_ORDER_OUT & "\" & l_File) Then
                    '    File.Delete(_ORDER_OUT & "\" & l_File)
                    'End If


                End If
            End If
        End If

    End Sub

    Private Sub Process_Johal()
        Dim lHour As Integer = 14, lMinute As Integer = 0, l_DateShift As Integer = 1
        If DateDiff(DateInterval.Day, _Johal_done, Now.Date) = 0 Then Return

        Select Case Date.Now.DayOfWeek
            Case DayOfWeek.Sunday
                Return
            Case DayOfWeek.Saturday
                lHour = 12
                l_DateShift = 2
            Case Else
        End Select

        If Date.Now.Hour = lHour AndAlso Date.Now.Minute > lMinute AndAlso Date.Now.Minute < lMinute + 5 Then
            If GetSetting_Johal() Then
                'We only need to run this once a day at the depots cut off
                If Export_DairyData("Johal_") Then
                    Dim l_File As String = "Johal_" & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd") & ".csv"
                    Dim l_Body As String = "Please find attached orders placed for delivery on " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("dd MMM yyyy")

                    Email_Generic(My.Settings.Johal_Order_Email, "", "Westcountry Milk Order for " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("dd MMM yyyy"),
                                      l_Body, _ORDER_OUT & "\" & l_File, "NoReply@wcmilk.co.uk", False, "")


                    System.Threading.Thread.Sleep(500)
                    MoveFile(_ORDER_OUT & "\" & l_File, _ORDER_ARCHIVED & "\" & l_File)

                    _Johal_done = Now.Date
                End If
            End If
        End If
    End Sub
    Private Sub Process_Grahams()
        Dim lHour As Integer = 13, lMinute As Integer = 30
        Dim l_File As String = "Grahams_"

        If DateDiff(DateInterval.Day, _Grahams_done, Now.Date) = 0 Then Return

        Select Case Date.Now.DayOfWeek
            Case DayOfWeek.Sunday
                Return
            Case DayOfWeek.Saturday
                lHour = 12 : lMinute = 0
            Case Else

        End Select

        If Date.Now.Hour = lHour AndAlso Date.Now.Minute >= lMinute AndAlso Date.Now.Minute < lMinute + 5 Then
            If GetSetting_Grahams() Then
                'We only need to run this once a day at the depots cut off
                If Export_DairyData(l_File) Then
                    Dim l_DateShift As Integer = 1
                    If Date.Now.DayOfWeek = DayOfWeek.Saturday Then l_DateShift = 2

                    'NOTE:   this order file is collected by supplier from WCM ftp folder

                    'Dim l_Body As String = "Please find attached orders placed for delivery on " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("dd MMM yyyy")

                    ''Email_Order(My.Settings.Graham_Order_Email, "", "Westcountry Milk Order for " & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("dd MMM yyyy"),
                    ''              l_Body, _ORDER_IN & "\" & l_File, "NoReply@wcmilk.co.uk", False, "")
                    _Grahams_done = Now.Date
                End If
            End If
        End If
    End Sub
    Private Function Export_DairyData(ByRef pFileName As String, Optional pFridayForMonday As Boolean = False, Optional pSummary As Boolean = False, Optional pExt As String = ".csv") As Boolean
        Dim l_File As String = Nothing
        Dim lResult As MsgBoxResult
        Dim l_DateShift As Integer = 1


        If Date.Now.DayOfWeek = DayOfWeek.Friday AndAlso pFridayForMonday Then
            l_DateShift = 3 'Get the orders for Monday
        ElseIf Date.Now.DayOfWeek = DayOfWeek.Saturday Then
            l_DateShift = 2 'Get the orders for Monday
        Else
            l_DateShift = 1 'Get the orders for Next day"
        End If

        l_File = pFileName & DateAdd(DateInterval.Day, l_DateShift, Date.Today).ToString("yyyy-MM-dd") & pExt

        Try
            If pSummary Then
                lResult = CreateOrders_DairyData(If(String.IsNullOrEmpty(_ORDER_ARCHIVED), _ORDER_OUT, _ORDER_ARCHIVED) & "\" & l_File, l_DateShift, pSummary)
                Return lResult = MsgBoxResult.Ok
            Else
                lResult = CreateOrders_DairyData(_ORDER_OUT & "\" & l_File, l_DateShift, pFileName.Equals("Paynes_"))
                If Wrap_Result(lResult, l_File, _Test_Mode, True) Then
                    Return True
                End If
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function CreateOrders_DairyData(ByVal pFileName As String, ByRef pDateShift As Integer, Optional pSummary As Boolean = False) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand = Nothing
        Dim lErrMsg As String = String.Empty
        Dim dt As DataTable = New DataTable("orders")
        Dim lPrevServingCode As String = ""
        Dim lCount As Integer = 0
        Dim sw As StreamWriter
        Dim line As String = ""
        Dim row As DataRow
        Try
            l_DB.Open()
            cmd = l_DB.SqlCommand("p_Export_Next_Days_Orders_by_Supplier")
            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                param = .Parameters.Add("@SupplierID", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.Input
                param.Value = _SUPPLIER_ID

                param = .Parameters.Add("@ShiftDate", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.Input
                param.Value = pDateShift

                If pSummary OrElse _SUPPLIER_ID = My.Settings.Johal_SUPPL_ID Then
                    param = .Parameters.Add("@ShowProdDesc", SqlDbType.Bit)
                    param.Direction = Data.ParameterDirection.Input
                    param.Value = 1
                End If

                l_DB.Fill(cmd, dt)
            End With

            cmd.Dispose()
            l_DB.Close()
            If dt.Rows.Count > 0 Then
                System.Threading.Thread.Sleep(20)
                sw = New StreamWriter(pFileName)
                If pSummary Then
                    With sw
                        Select Case _SUPPLIER_ID
                            Case My.Settings.Broadland_SUPPL_ID, My.Settings.MillsMilk_SUPPL_ID ''BFS1 - Broadland Food Service,Fresh Pastures
                                .WriteLine("Serving_code,Site_name,DeliveryDate,OrderNo,Code,Product,qty")
                                For Each row In dt.Rows
                                    .WriteLine(row(0) & "," & row(6) & "," & row(1) & "," & row(2) & "," & row(3) & "," & row(4) & "," & row(5))
                                Next
                            Case Else
                                .WriteLine("Serving_code,DeliveryDate,OrderNo,Code,Product,qty")
                                For Each row In dt.Rows
                                    .WriteLine(row(0) & "," & row(1) & "," & row(2) & "," & row(3) & "," & row(4) & "," & row(5))
                                Next
                        End Select

                        .Close()
                        .Dispose()
                    End With
                Else
                    With sw
                        Select Case _SUPPLIER_ID
                            Case 118867089  ''BFS1 - Broadland Food Service   
                                .WriteLine("Serving_code,Site_name,DeliveryDate,OrderNo,Code,qty")
                                For Each row In dt.Rows
                                    .WriteLine(row(0) & "," & row(5) & "," & row(1) & "," & row(2) & "," & row(3) & "," & row(4))
                                Next
                            'Case My.Settings.Chew_Valley_SUPPL_ID ' 224513559  ''CHEW1 - Chew Valley Dairy
                           '    Return ToJson(pFileName, dt)
                            Case My.Settings.Johal_SUPPL_ID
                                .WriteLine("Serving_code,DeliveryDate,OrderNo,Code,Product,qty")
                                For Each row In dt.Rows
                                    .WriteLine(row(0) & "," & row(1) & "," & row(2) & "," & row(3) & "," & row(4) & "," & row(5))
                                Next
                            Case Else

                                .WriteLine("Serving_code,DeliveryDate,OrderNo,Code,qty")
                                For Each row In dt.Rows
                                    .WriteLine(row(0) & "," & row(1) & "," & row(2) & "," & row(3) & "," & row(4))
                                Next
                        End Select
                        .Close() : .Dispose()
                    End With
                End If
                Return MsgBoxResult.Ok
            Else
                Return MsgBoxResult.Abort
            End If


        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally
            If Not sw Is Nothing Then sw.Close() : sw.Dispose()
        End Try
        Return MsgBoxResult.Abort
    End Function

    Private Function ToJson(ByVal pPath As String, pDT As DataTable) As MsgBoxResult
        Dim lOrderNum As String = "" _
          , swr As StreamWriter = Nothing _
          , lNewOrder As Boolean = False

        Try
            For Each row As DataRow In pDT.Rows
                If String.IsNullOrEmpty(lOrderNum) Then
                    lOrderNum = row("OrderNo").ToString
                    lNewOrder = True
                ElseIf Not row("OrderNo").ToString.Equals(lOrderNum) AndAlso Not IsNothing(swr) Then
                    With swr
                        .WriteLine("        }")
                        .WriteLine("    ]")
                        .WriteLine("}")
                        .Close() : .Dispose()
                    End With
                    lOrderNum = row("OrderNo").ToString
                    lNewOrder = True
                End If
                If lNewOrder Then
                    lNewOrder = False
                    swr = New StreamWriter(pPath & lOrderNum & ".json")
                    With swr
                        .WriteLine("{")
                        .Write("    ""customer_code"": """) : .Write(row("Serving_code")) : .WriteLine(""",")
                        .Write("    ""comment"": """) : .Write(row("notes")) : .WriteLine(""",")
                        .Write("    ""delivery_date"": """) : .Write(CType(row("DeliveryDate"), Date).ToString("yyyy-MM-dd")) : .WriteLine(""",")
                        .Write("    ""po_number"": """) : .Write(row("OrderNo")) : .WriteLine(""",")
                        .WriteLine("    ""apply_customer_discount"": 0,")
                        .WriteLine("    ""discount"": """",")
                        .WriteLine("    ""apply_customer_delivery_charge"": 0,")
                        .WriteLine("    ""delivery_charge"": """",")
                        .WriteLine("    ""OrderDetail"": [")
                    End With
                Else
                    swr.WriteLine("     },")
                End If
                With swr
                    .WriteLine("        {")
                    .Write("            ""item_code"": """) : .Write(row("Code")) : .WriteLine(""",")
                    .Write("            ""quantity"": """) : .Write(row("Qty")) : .WriteLine("""")
                    ' .WriteLine("            ""comment"": "",")
                    '  .WriteLine("            ""item_price"": """",")
                End With
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally
            If Not IsNothing(swr) Then
                swr.Close() : swr.Dispose()
            End If
        End Try
        Return MsgBoxResult.Abort

    End Function

#End Region

#Region "Methods Email Orders--------------------------------------------------------------"

    Private Sub Process_EmailOrders()
        Dim asFiles() As String
        Dim l_File As String
        Dim idx As Integer
        Dim bFailedToSendResponces As Boolean = False
        Dim lMsg As String = String.Empty
        Dim lResult As MsgBoxResult

        Try
            'PROCESS NEW ORDERS
            asFiles = Directory.GetFiles(_ORDER_IN)

            For idx = asFiles.GetLowerBound(0) To asFiles.GetUpperBound(0)
                l_File = Path.GetFileName(asFiles(idx))
                If l_File.StartsWith("WCM_EmailOrder-") Then
                    MyEventLog.WriteEntry("ORDER REQUEST RECEIVED:  " & l_File, EventLogEntryType.Information, GetEventID)

                    lResult = UploadFile_EmailOrders(_ORDER_IN & "\" & l_File)
                    If Not Wrap_Result(lResult, l_File, _Test_Mode AndAlso (idx = asFiles.GetLowerBound(0) OrElse idx = asFiles.GetUpperBound(0))) Then
                        Return
                    End If
                End If
            Next

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
    End Sub

    Private Function UploadFile_EmailOrders(ByVal pFileName As String) As MsgBoxResult
        Dim l_DB = New DB
        Dim param As SqlClient.SqlParameter
        Dim cmd As SqlClient.SqlCommand
        Dim lXMLContents As String = String.Empty
        Dim lReader As StreamReader
        Dim lRetVal As Integer = 0
        Dim lErrMsg As String = String.Empty
        Dim lResponseMsg As String = String.Empty
        Dim dt As DataTable = Nothing
        Dim lCustomeLocationCode As String = String.Empty
        Dim lUnitCode As String = String.Empty
        Dim lArgeementCode As String = String.Empty
        Dim lNewOrderID As Integer = 0

        Try
            l_DB.Open()

            lReader = New StreamReader(pFileName)
            lXMLContents = lReader.ReadToEnd()
            lXMLContents = lXMLContents.Replace("<?xml version=""1.0"" encoding=""utf-8""?>", "")
            lReader.Close()

            cmd = l_DB.SqlCommand("p_P2P_order_import_WCM_EmailOrders")

            With cmd
                .CommandType = Data.CommandType.StoredProcedure

                'Return Value
                param = .Parameters.Add("@ret", SqlDbType.Int)
                param.Direction = Data.ParameterDirection.ReturnValue

                param = .Parameters.Add("@record_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@file_name", SqlDbType.VarChar, 200)
                param.Direction = ParameterDirection.Input
                param.Value = pFileName

                param = .Parameters.Add("@xml_order", SqlDbType.Xml)
                param.Direction = ParameterDirection.Input
                param.Value = lXMLContents

                param = .Parameters.Add("@order_num", SqlDbType.VarChar, 20)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@acc_num", SqlDbType.VarChar, 30)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                param = .Parameters.Add("@delivery_date_requested", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@delivery_date", SqlDbType.Date)
                param.Direction = ParameterDirection.InputOutput
                param.Value = Date.MinValue

                param = .Parameters.Add("@datetime_created_string", SqlDbType.VarChar, 200)
                param.Direction = ParameterDirection.InputOutput
                param.Value = DBNull.Value

                param = .Parameters.Add("@status", SqlDbType.VarChar, 16)
                param.Direction = ParameterDirection.InputOutput
                param.Value = "ACCEPTED"

                param = .Parameters.Add("@customer_order_header_id", SqlDbType.Int)
                param.Direction = ParameterDirection.InputOutput
                param.Value = 0

                param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                param.Direction = ParameterDirection.InputOutput
                param.Value = ""

                'param = .Parameters.Add("@buyer_seq", SqlDbType.Int)
                'param.Direction = ParameterDirection.Input
                'param.Value = 99 '$$MATT  will need to add a new flag in XML to switch between mapping types

                .ExecuteNonQuery()
                lRetVal = CType(.Parameters("@ret").Value, Integer)

                _OrderRequestId = .Parameters("@record_id").Value
                _OrderNum = Nz(Of String)(.Parameters("@order_num").Value, "")
                _AccNum = Nz(Of String)(.Parameters("@acc_num").Value, "")
                _DeliveryDateRequested = Nz(Of Date)(.Parameters("@delivery_date_requested").Value, Date.MinValue)
                _DeliveryDate = Nz(Of Date)(.Parameters("@delivery_date").Value, Date.MinValue)
                _Status = Nz(Of String)(.Parameters("@status").Value, "")
                _Order_DateTime_Created = Nz(Of String)(.Parameters("@datetime_created_string").Value, "")

                'lArgeementCode = Nz(Of String)(.Parameters("@customer_agreement_code").Value, "")
                'lCustomeLocationCode = Nz(Of String)(.Parameters("@customer_location_code").Value, "")
                'lUnitCode = Nz(Of String)(.Parameters("@unit_code").Value, "")
                lNewOrderID = Nz(Of Integer)(.Parameters("@customer_order_header_id").Value, 0)

                lErrMsg = Nz(Of String)(.Parameters("@err_msg").Value, "")
            End With
            cmd.Dispose()
            l_DB.Close()
            dt = GetOrderLines()

            Select Case lRetVal
                Case 0 ' Success
                    MyEventLog.WriteEntry("Order Processed: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum, EventLogEntryType.Information, GetEventID)
                    If _Status.StartsWith("MODIFIED") Then
                        _EmailServiceMessage(Date.Now & " ORDER Processed: " & _Status & " " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & vbNewLine & vbNewLine & lResponseMsg)
                    End If
                    System.Threading.Thread.Sleep(50)
                    EmailOrder(lNewOrderID)
                    Return MsgBoxResult.Ok
                Case Else
                    _Status = "REJECTED"

                    If lErrMsg = "CUSTOMER_IDENTIFICATION_NUMBER_IS_INVALID" Then lErrMsg = " SUSPENDED ACCOUNT "
                    MyEventLog.WriteEntry("Order REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg, EventLogEntryType.Warning, GetEventID)

                    _EmailServiceMessage(Date.Now & " ORDER REJECTED: " & _OrderNum & " {" & _OrderRequestId & "} for " & _AccNum & lErrMsg & vbNewLine & vbNewLine & lResponseMsg)
                    System.Threading.Thread.Sleep(50)
            End Select


        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            If ex.Message.ToString.ToUpperInvariant.Contains("TIMEOUT EXPIRED") Then
                Return MsgBoxResult.Retry
            End If
        Finally

        End Try
        Return MsgBoxResult.Abort
    End Function

#End Region

#Region "Supporting Functions --------------------------------------------------------------"

    Private Function GetOrderLines() As DataTable
        Dim l_DB = New DB
        Dim cmd As SqlClient.SqlCommand = Nothing
        Dim dt As DataTable = New DataTable("order_lines")
        Try
            l_DB.Open()
            cmd = l_DB.SqlCommand("p_P2P_order_lines_get")
            With cmd
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.Add("@record_id", SqlDbType.Int)
                .Parameters("@record_id").Value = _OrderRequestId
            End With
            l_DB.Fill(cmd, dt)
            cmd.Dispose()
            l_DB.Close()
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
        Return dt
    End Function
    Private Function GetSOAmendmentLines(pAmendmentID As Integer) As DataTable
        Dim l_DB = New DB
        Dim cmd As SqlClient.SqlCommand = Nothing
        Dim dt As DataTable = New DataTable("order_lines")
        Try
            l_DB.Open()
            cmd = l_DB.SqlCommand("p_SO_amendment_lines_get")
            With cmd
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.Add("@order_id", SqlDbType.Int)
                .Parameters("@order_id").Value = pAmendmentID
            End With
            l_DB.Fill(cmd, dt)
            cmd.Dispose()
            l_DB.Close()
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
        Return dt
    End Function
    Private Function GetSOLines(pSOHID As Integer, pDeliveryDate As Date) As DataTable
        Dim l_DB = New DB
        Dim cmd As SqlClient.SqlCommand = Nothing
        Dim dt As DataTable = New DataTable("SO_lines")
        Try
            l_DB.Open()
            cmd = l_DB.SqlCommand("p_SO_lines_get")
            With cmd
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.Add("@SOH_id", SqlDbType.Int)
                .Parameters("@SOH_id").Value = pSOHID
                .Parameters.Add("@delivery_date", SqlDbType.Date)
                .Parameters("@delivery_date").Value = pDeliveryDate
            End With
            l_DB.Fill(cmd, dt)
            cmd.Dispose()
            l_DB.Close()

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
        Return dt
    End Function
    Private Function UpdateAcknowledgementDate() As Integer
        Dim l_DB As New DB
        Dim cmd As SqlClient.SqlCommand = Nothing
        Dim l_RecAffected As Integer

        Try
            l_DB.Open()

            cmd = New SqlClient.SqlCommand("UPDATE P2P_order_headers set p2h_acknowledgement_datetime = GETDATE() WHERE p2h_record_id = @record_id", l_DB.Connection)
            cmd.CommandTimeout = 6000
            cmd.Parameters.AddWithValue("@record_id", _OrderRequestId)

            l_RecAffected = cmd.ExecuteNonQuery()
            cmd.Dispose()
            l_DB.Close()
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try

        Return l_RecAffected
    End Function

    Private Sub AddAttribute(ByRef xmlWriter As XmlTextWriter, ByVal strAttribute As String, ByVal strValue As String)
        With xmlWriter
            .WriteStartAttribute(strAttribute)
            .WriteString(strValue)
            .WriteEndAttribute()
        End With
    End Sub

    Private Function MoveFile(ByVal pSource As String, ByVal pDestination As String) As Boolean
        Try
            If File.Exists(pDestination) Then
                File.Delete(pDestination)
            End If

            File.Move(pSource, pDestination)

            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
            Return False
        End Try
    End Function

    Private Function Nz(Of objDataType)(ByVal objVal As Object, ByVal objRet As objDataType) As objDataType
        If IsDBNull(objVal) OrElse IsNothing(objVal) Then
            Return objRet
        Else
            Return CType(objVal, objDataType)
        End If
    End Function

    Private Function GetEventID(pCommon As Integer) As Integer
        If _eventId_Common > 99 Then _eventId_Common = 1 Else _eventId_Common += 1
        Return _eventId_Common
    End Function

    Private Function GetEventID() As Integer
        Select Case COrderHeader.BuyerSequence
            Case 0, 1000
                ' NOTE: eventId 500 reserved to PushEmailWEBAPP; 999 - to PushEmailP2P
                If _eventId_Order_Upload > 998 Then _eventId_Order_Upload = 101 Else _eventId_Order_Upload += 1
                If _eventId_Order_Upload = 500 Then _eventId_Order_Upload += 1
                Return _eventId_Order_Upload
            Case 7
                'If _eventId_Elior > 9999 Then _eventId_Elior = 1001 Else _eventId_Elior += 1
                'Return _eventId_Elior
            Case 9
                If _eventId_Cypad > 19999 Then _eventId_Cypad = 15001 Else _eventId_Cypad += 1
                Return _eventId_Cypad
            Case 11
                If _eventId_Bourne > 24999 Then _eventId_Bourne = 20001 Else _eventId_Bourne += 1
                Return _eventId_Bourne
            Case 10
                If _eventId_Medina > 29999 Then _eventId_Medina = 25001 Else _eventId_Medina += 1
                Return _eventId_Medina
            Case 12
                If _eventId_Interserve > 34999 Then _eventId_Interserve = 30001 Else _eventId_Interserve += 1
                Return _eventId_Interserve
            Case 5
                'redundant
                If _eventId_Compass > 39999 Then _eventId_Compass = 35001 Else _eventId_Compass += 1
                Return _eventId_Compass
            Case 8
                If _eventId_FoodBuy_Online > 44999 Then _eventId_FoodBuy_Online = 40001 Else _eventId_FoodBuy_Online += 1
                Return _eventId_FoodBuy_Online
            Case 99
                If _eventId_Email_Orders > 49999 Then _eventId_Email_Orders = 45001 Else _eventId_Email_Orders += 1
                Return _eventId_Email_Orders
            Case 98
                If _eventId_DN_Grahams > 54999 Then _eventId_DN_Grahams = 50001 Else _eventId_DN_Grahams += 1
                Return _eventId_DN_Grahams
            Case 14
                If _eventId_Poundland > 59999 Then _eventId_Poundland = 55001 Else _eventId_Poundland += 1
                Return _eventId_Poundland
            Case COrderHeader.BuyerSeq.CaffeNero_6
                If _eventId_CN_CSV > 64999 Then _eventId_CN_CSV = 60001 Else _eventId_CN_CSV += 1
                Return _eventId_CN_CSV
            Case COrderHeader.BuyerSeq.Zupa_16
                If _eventId_Zupa > 9999 Then _eventId_CN_CSV = 1001 Else _eventId_Zupa += 1
                Return _eventId_Zupa
            Case COrderHeader.BuyerSeq.McColls_18
                If _eventId_McColls > 14999 Then _eventId_McColls = 10001 Else _eventId_McColls += 1
                Return _eventId_McColls
            Case COrderHeader.BuyerSeq.Weezy_19
                ' max for  Int32 =65535
                If _eventId_Weezy > 65534 Then _eventId_Weezy = 65001 Else _eventId_Weezy += 1
                Return _eventId_Weezy
            Case 500
                Return 500 ' fixed
            Case 999
                Return 999 ' fixed
            Case Else
                If _eventId_Common > 99 Then _eventId_Common = 1 Else _eventId_Common += 1
                Return _eventId_Common
        End Select
    End Function

    Private Function GetSetting_Foodbuy_Online() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.FoodBuy_Online_8
            If My.Settings.Switch_FoodBuyOnline = 0 Then ' switched off
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _ORDER_FAILED = String.Empty
                _RESPONSE_OUT = String.Empty
                _RESPONSE_ARCHIVED = String.Empty

            ElseIf My.Settings.Switch_FoodBuyOnline = 1 Then ' test
                _ORDER_IN = My.Settings.FoodBuyOnline_IN_test
                _ORDER_ARCHIVED = My.Settings.FoodBuyOnline_Archive_test
                _ORDER_FAILED = My.Settings.FoodBuyOnline_Failed_test
                _RESPONSE_OUT = My.Settings.FoodBuyOnline_OUT
                _RESPONSE_ARCHIVED = My.Settings.FoodBuyOnline_Archive_test
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_FoodBuyOnline = 2 Then 'production
                _ORDER_IN = My.Settings.FoodBuyOnline_IN
                _ORDER_ARCHIVED = My.Settings.FoodBuyOnline_Archive
                _ORDER_FAILED = My.Settings.FoodBuyOnline_Failed
                _RESPONSE_OUT = My.Settings.FoodBuyOnline_OUT
                _RESPONSE_ARCHIVED = My.Settings.FoodBuyOnline_Archive
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function



    Private Function GetSetting_Bourne() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.BourneLeisure_11
            If My.Settings.Switch_Bourne = 0 Then ' switched off
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _ORDER_FAILED = String.Empty
                _RESPONSE_OUT = String.Empty
                _RESPONSE_ARCHIVED = String.Empty

            ElseIf My.Settings.Switch_Bourne = 1 Then ' test
                _ORDER_IN = My.Settings.BOURNE_IN_test
                _ORDER_ARCHIVED = My.Settings.BOURNE_Archive_test
                _ORDER_FAILED = My.Settings.BOURNE_Failed_test
                _RESPONSE_OUT = My.Settings.BOURNE_OUT_test
                _RESPONSE_ARCHIVED = My.Settings.BOURNE_RESPONSE_ARCHIVED_test
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_Bourne = 2 Then 'production
                _ORDER_IN = My.Settings.BOURNE_IN
                _ORDER_ARCHIVED = My.Settings.BOURNE_Archive
                _ORDER_FAILED = My.Settings.BOURNE_Failed
                _RESPONSE_OUT = My.Settings.BOURNE_OUT
                _RESPONSE_ARCHIVED = My.Settings.BOURNE_RESPONSE_ARCHIVED
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function GetSetting_Interserve_saffron() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Interserve_12
            If My.Settings.Switch_Interserve_Saffron = 0 Then ' switched off
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _ORDER_FAILED = String.Empty
                _RESPONSE_OUT = String.Empty
                _RESPONSE_ARCHIVED = String.Empty

            ElseIf My.Settings.Switch_Interserve_Saffron = 1 Then ' test
                _ORDER_IN = My.Settings.INTERSERVE_IN_test_SAFFRON
                _ORDER_ARCHIVED = My.Settings.INTERSERVE_Archive_test_SAFFRON
                _ORDER_FAILED = My.Settings.INTERSERVE_Failed_test_SAFFRON
                _RESPONSE_OUT = My.Settings.INTERSERVE_OUT_test_SAFFRON
                _RESPONSE_ARCHIVED = My.Settings.INTERSERVE_RESPONSE_ARCHIVED_test_SAFFRON
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_Interserve_Saffron = 2 Then 'production
                _ORDER_IN = My.Settings.INTERSERVE_IN_SAFFRON
                _ORDER_ARCHIVED = My.Settings.INTERSERVE_Archive_SAFFRON
                _ORDER_FAILED = My.Settings.INTERSERVE_Failed_SAFFRON
                _RESPONSE_OUT = My.Settings.INTERSERVE_OUT_SAFFRON
                _RESPONSE_ARCHIVED = My.Settings.INTERSERVE_RESPONSE_ARCHIVED_SAFFRON
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function GetSetting_CN_CrunchTime() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.CaffeNero_6
            If My.Settings.Switch_CN_CrunchTime = 0 Then ' switched off
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _ORDER_FAILED = String.Empty
                _RESPONSE_OUT = String.Empty
                _RESPONSE_ARCHIVED = String.Empty

            ElseIf My.Settings.Switch_CN_CrunchTime = 1 Then ' test
                _ORDER_IN = My.Settings.CN_CrunchTime_IN_test
                _ORDER_ARCHIVED = My.Settings.CN_CrunchTime_Archived_Test
                _ORDER_FAILED = My.Settings.CN_CrunchTime_Failed_test
                _RESPONSE_OUT = My.Settings.CN_CrunchTime_OUT_test
                _RESPONSE_ARCHIVED = My.Settings.CN_CrunchTime_RESPONSE_Archived_test
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_CN_CrunchTime = 2 Then 'production
                _ORDER_IN = My.Settings.CN_CrunchTime_IN
                _ORDER_ARCHIVED = My.Settings.CN_CrunchTime_Archive
                _ORDER_FAILED = My.Settings.CN_CrunchTime_Failed
                _RESPONSE_OUT = My.Settings.CN_CrunchTime_OUT
                _RESPONSE_ARCHIVED = My.Settings.CN_CrunchTime_RESPONSE_Archived
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function


    Private Function GetSetting_PushEmail(pType As String) As Boolean
        Try

            _ORDER_IN = String.Empty
            _ORDER_ARCHIVED = String.Empty
            _ORDER_FAILED = String.Empty
            _RESPONSE_OUT = String.Empty
            _RESPONSE_ARCHIVED = String.Empty

            Select Case pType.ToUpper
                Case "P2P"
                    COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.PushEmailP2P_999
                    If My.Settings.Switch_PushEmailP2P = 0 Then
                        Return False
                    ElseIf My.Settings.Switch_PushEmailP2P = 1 Then ' test
                        _Test_Mode = True
                        Return True
                    ElseIf My.Settings.Switch_PushEmailP2P = 2 Then
                        _Test_Mode = False
                        Return True
                    End If
                Case "WEBAPP"
                    COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.PushEmailWebApp_500
                    If My.Settings.Switch_PushEmailWebApp = 0 Then
                        Return False
                    ElseIf My.Settings.Switch_PushEmailWebApp = 1 Then ' test
                        _Test_Mode = True
                        Return True
                    ElseIf My.Settings.Switch_PushEmailWebApp = 2 Then
                        _Test_Mode = False
                        Return True
                    End If
            End Select

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function GetSetting_DN_Grahams() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.DN_Grahams_98
            If My.Settings.Switch_DN_Grahams = 0 Then ' switched off
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _ORDER_FAILED = String.Empty
                _RESPONSE_OUT = String.Empty
                _RESPONSE_ARCHIVED = String.Empty

            ElseIf My.Settings.Switch_DN_Grahams = 1 Then ' test
                _ORDER_IN = My.Settings.DN_Grahams_IN
                _ORDER_ARCHIVED = My.Settings.DN_Grahams_Archive
                _ORDER_FAILED = My.Settings.DN_Grahams_Failed
                _RESPONSE_OUT = String.Empty
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_DN_Grahams = 2 Then 'production
                _ORDER_IN = My.Settings.DN_Grahams_IN
                _ORDER_ARCHIVED = My.Settings.DN_Grahams_Archive
                _ORDER_FAILED = My.Settings.DN_Grahams_Failed
                _RESPONSE_OUT = String.Empty
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function GetSetting_Johal() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000
            If My.Settings.Switch_Johal = 0 Then ' switched off
                _ORDER_OUT = String.Empty
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _SUPPLIER_ID = 0
            ElseIf My.Settings.Switch_Johal = 1 Then ' test
                _ORDER_OUT = My.Settings.Johal_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.Johal_ARCHIVE
                _SUPPLIER_ID = My.Settings.Johal_SUPPL_ID
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_Johal = 2 Then 'production
                _ORDER_OUT = My.Settings.Johal_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.Johal_ARCHIVE
                _SUPPLIER_ID = My.Settings.Johal_SUPPL_ID
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function
    Private Function GetSetting_Grahams() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000
            If My.Settings.Switch_Grahams = 0 Then ' switched off
                _ORDER_OUT = String.Empty
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _SUPPLIER_ID = 0
            ElseIf My.Settings.Switch_Grahams = 1 Then ' test
                _ORDER_OUT = My.Settings.Grahams_OUT.Replace("Live", "Test")
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.Grahams_ARCHIVE.Replace("Live", "Test")
                _SUPPLIER_ID = My.Settings.Grahams_SUPPL_ID
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_Grahams = 2 Then 'production
                _ORDER_OUT = My.Settings.Grahams_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.Grahams_ARCHIVE
                _SUPPLIER_ID = My.Settings.Grahams_SUPPL_ID
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function


    Private Function GetSetting_DairyData_MillsMilk() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000
            If My.Settings.Switch_DairyData_MillsMilk = 0 Then ' switched off
                _ORDER_OUT = String.Empty
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _SUPPLIER_ID = 0
            ElseIf My.Settings.Switch_DairyData_MillsMilk = 1 Then ' test
                _ORDER_OUT = My.Settings.DairyData_MillsMilk_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.DairyData_MillsMilk_Archive
                _SUPPLIER_ID = My.Settings.MillsMilk_SUPPL_ID
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_DairyData_MillsMilk = 2 Then 'production
                _ORDER_OUT = My.Settings.DairyData_MillsMilk_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.DairyData_MillsMilk_Archive
                _SUPPLIER_ID = My.Settings.MillsMilk_SUPPL_ID
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function
    Private Function GetSetting_AllanReeder() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000
            If My.Settings.Switch_AllanReeder = 0 Then ' switched off
                _ORDER_OUT = String.Empty
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _SUPPLIER_ID = 0
            ElseIf My.Settings.Switch_AllanReeder = 1 Then ' test
                _ORDER_OUT = My.Settings.AllanReeder_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.AllanReeder_ARCHIVE
                _SUPPLIER_ID = My.Settings.AllanReeder_SUPPL_ID
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_AllanReeder = 2 Then 'production
                _ORDER_OUT = My.Settings.AllanReeder_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.AllanReeder_ARCHIVE
                _SUPPLIER_ID = My.Settings.AllanReeder_SUPPL_ID
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function GetSetting_DairyData_Paynes() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000
            If My.Settings.Switch_Paynes = 0 Then ' switched off
                _ORDER_OUT = String.Empty
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _SUPPLIER_ID = 0
            ElseIf My.Settings.Switch_Paynes = 1 Then ' test
                _ORDER_OUT = My.Settings.Paynes_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.Paynes_Archived
                _SUPPLIER_ID = My.Settings.Paynes_SUPPL_ID
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_Paynes = 2 Then 'production
                _ORDER_OUT = My.Settings.Paynes_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.Paynes_Archived
                _SUPPLIER_ID = My.Settings.Paynes_SUPPL_ID
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function GetSetting_DairyData_Broadland() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000
            If My.Settings.Switch_Broadland = 0 Then ' switched off
                _ORDER_OUT = String.Empty
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _SUPPLIER_ID = 0
            ElseIf My.Settings.Switch_Broadland = 1 Then ' test
                _ORDER_OUT = My.Settings.Broadland_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.Broadland_OUT & "\Archive"
                _SUPPLIER_ID = My.Settings.Broadland_SUPPL_ID
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_Broadland = 2 Then 'production
                _ORDER_OUT = My.Settings.Broadland_OUT
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.Broadland_OUT & "\Archive"
                _SUPPLIER_ID = My.Settings.Broadland_SUPPL_ID
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function GetSetting_JJWilson() As Boolean
        Try
            COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000
            If My.Settings.Switch_JJWilson = 0 Then ' switched off
                _ORDER_OUT = String.Empty
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = String.Empty
                _SUPPLIER_ID = 0
            ElseIf My.Settings.Switch_JJWilson = 1 Then ' test
                _ORDER_OUT = My.Settings.JJWilson_Orders
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.JJWilson_Archived
                _SUPPLIER_ID = 0
                _Test_Mode = True
                Return True
            ElseIf My.Settings.Switch_JJWilson = 2 Then 'production
                _ORDER_OUT = My.Settings.JJWilson_Orders
                _ORDER_IN = String.Empty
                _ORDER_ARCHIVED = My.Settings.JJWilson_Archived
                _SUPPLIER_ID = 0
                _Test_Mode = False
                Return True
            End If

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Private Function Wrap_Result(p_Result As MsgBoxResult, p_file As String, p_IsDetailedLog As Boolean, Optional p_DontMove As Boolean = False) As Boolean
        Try
            If p_Result = MsgBoxResult.Ok Then
                If COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.DN_Grahams_98 Then
                    MyEventLog.WriteEntry("DELIVERY NOTE PROCESSED:  " & _DeliveryNoteNum & " for " & _AccNum, EventLogEntryType.Information, GetEventID)
                    _EmailServiceMessage(Date.Now & " DELIVERY NOTE PROCESSED:  " & _DeliveryNoteNum & " for " & _AccNum)
                ElseIf COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000 Then
                    MyEventLog.WriteEntry("ORDER FILE GENERATED :  " & p_file, EventLogEntryType.Information, GetEventID)
                    'If p_IsDetailedLog Then
                    _EmailServiceMessage(Date.Now & " ORDER FILE GENERATED :  " & p_file)
                    'End If
                ElseIf COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.McColls_18 Then
                    If p_IsDetailedLog Then
                        MyEventLog.WriteEntry("ORDER FILE PROCESSED :  " & p_file, EventLogEntryType.Information, GetEventID)
                        _EmailServiceMessage(Date.Now & " ORDER FILE PROCESSED :  " & p_file)
                    End If
                Else
                    If p_IsDetailedLog Then
                        MyEventLog.WriteEntry("ORDER REQUEST PROCESSED:  " & _OrderNum & " for " & _AccNum, EventLogEntryType.Information, GetEventID)
                        _EmailServiceMessage(Date.Now & " ORDER REQUEST PROCESSED:  " & _OrderNum & " for " & _AccNum)
                    End If
                End If

                If Directory.Exists(_ORDER_IN) Then
                    If p_DontMove Then
                        FileCopy(_ORDER_IN & "\" & p_file, _ORDER_ARCHIVED & "\" & p_file)
                    Else
                        MoveFile(_ORDER_IN & "\" & p_file, _ORDER_ARCHIVED & "\" & p_file)
                    End If
                End If

                If COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000 AndAlso Not (Directory.Exists(_ORDER_IN)) Then
                    If Not String.IsNullOrEmpty(_ORDER_ARCHIVED) AndAlso Directory.Exists(_ORDER_ARCHIVED) Then
                        FileCopy(_ORDER_OUT & "\" & p_file, _ORDER_ARCHIVED & "\" & p_file)
                    End If
                End If

            ElseIf p_Result = MsgBoxResult.Retry Then
                If COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.DN_Grahams_98 Then
                    MyEventLog.WriteEntry("DELIVERY NOTE FAILED: " & _DeliveryNoteNum & ". NEXT ATTEMPT IN " & My.Settings.TimerInterval / 1000 & " SECONDS", EventLogEntryType.Warning, GetEventID)
                    _EmailServiceMessage(Date.Now & " DELIVERY NOTE FAILED: " & _DeliveryNoteNum & ". NEXT ATTEMPT IN " & My.Settings.TimerInterval / 1000 & " SECONDS", True)
                ElseIf COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000 Then
                    MyEventLog.WriteEntry("ORDER FILE FAILED TO GENERATE :  " & p_file, EventLogEntryType.Warning, GetEventID)
                    _EmailServiceMessage(Date.Now & " ORDER FILE FAILED TO GENERATE :  " & p_file, True)
                Else
                    MyEventLog.WriteEntry("ORDER REQUEST FAILED: " & _OrderNum & ". NEXT ATTEMPT IN " & My.Settings.TimerInterval / 1000 & " SECONDS", EventLogEntryType.Information, GetEventID)
                    _EmailServiceMessage(Date.Now & " ORDER REQUEST FAILED: " & _OrderNum & ". NEXT ATTEMPT IN " & My.Settings.TimerInterval / 1000 & " SECONDS", True)
                End If
                Return False
            Else  ' MsgBoxResult.Abort 
                If COrderHeader.BuyerSequence <> COrderHeader.BuyerSeq.Order_Upload_1000 Then
                    If File.Exists(_ORDER_IN & "\" & p_file) Then
                        MoveFile(_ORDER_IN & "\" & p_file, _ORDER_FAILED & "\" & p_file)
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID(0))
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        End Try
        Return False
    End Function

    Public Function sDate(ByVal dDate As Object, Optional ByVal sSeparator As String = " ") As String
        ' Return date in correct format for SQL string
        If IsDate(dDate) Then

            'System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-GB", True)
            'System.Threading.Thread.CurrentThread.CurrentUICulture = New System.Globalization.CultureInfo("en-GB", True)

            Return DateAndTime.Day(dDate) & sSeparator & MonthName(Month(dDate), True) & sSeparator & Year(dDate)
        Else
            Return dDate.ToString
        End If
    End Function

#End Region

#Region "Events --------------------------------------------------------------"
    Private Sub MyEventLog_EntryWritten(sender As System.Object, e As EntryWrittenEventArgs) Handles MyEventLog.EntryWritten

    End Sub

    Private Sub mOrder_Report_Error(pErrMsg As String) Handles mOrder.Report_Error
        MyEventLog.WriteEntry(pErrMsg, EventLogEntryType.Error, GetEventID)
        _EmailServiceMessage(Date.Now & pErrMsg, True)
    End Sub

#End Region

#Region "Emailing --------------------------------------------------------------"

    Private Function EmailOrder(ByVal pOrderID As Integer, Optional ByVal pFileAttachemnt As ArrayList = Nothing, Optional ByVal pStaticNotes As String = "") As Boolean
        Dim lTextAttachments As Dictionary(Of String, System.Text.StringBuilder) _
        , lTestEmail As String = "" 'My.Settings.OrderTo 
        Dim lOrderInfo As String = String.Empty

        Try
            mOrder = New COrderHeader(pOrderID, COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.PushEmailWebApp_500)
            If mOrder Is Nothing Then
                MyEventLog.WriteEntry("ERROR: Failed to retrieve Order ID = " & pOrderID.ToString & " Buyer " & COrderHeader.BuyerSeq.PushEmailWebApp_500.ToString & " in " & _MeName & ".EmailOrder", EventLogEntryType.Error, GetEventID)
                _EmailServiceMessage(Date.Now & " ERROR: Failed to retrieve Order ID = " & pOrderID.ToString & " Buyer " & COrderHeader.BuyerSeq.PushEmailWebApp_500.ToString & " in " & _MeName & ".EmailOrder", True)
                Return False
            End If

            If mOrder.EDIOrdersFlag Then
                ' commented out by VS on 27/09/2023 - applies to any depot with edi flag on
                'If mOrder.SupplierID = My.Settings.Grahams_SUPPL_ID OrElse mOrder.SupplierID_Produce = My.Settings.Grahams_SUPPL_ID _
                '        OrElse mOrder.SupplierID = My.Settings.Johal_SUPPL_ID OrElse mOrder.SupplierID = My.Settings.MillsMilk_SUPPL_ID Then
                Return True
                'End If
            End If

            lOrderInfo = "Order Number: " & mOrder.OrderNum & " for " & mOrder.CustAcc.Trim & " delivery " & sDate(mOrder.DeliveryDate)
            If _Test_Mode Then
                lOrderInfo = "TEST" & lOrderInfo
            End If
            If pStaticNotes <> "" Then
                mOrder.Notes += pStaticNotes
            End If
            mOrder.GetProductCodes()
            If mOrder.TotalQty = 0 Then
                MyEventLog.WriteEntry("Zero Product Quantity. " & lOrderInfo, EventLogEntryType.Warning, GetEventID)
                _EmailServiceMessage(Date.Now & " " & "ERROR: Zero Product Quantity. " & lOrderInfo, True)
                Return False
            Else
                If mOrder.ProduceItems.Count > 0 AndAlso mOrder.DepotId_Produce <> 0 Then
                    If mOrder.HOId = 245739642 OrElse mOrder.HOId = 245740000 OrElse mOrder.Depot.Contains("Hovis") Then 'ISS1 and ISS2
                        'Add this line in the notes for all ISS sites
                        mOrder.Notes = "</br></br></br> ---- PLEASE SUPPLY ONLY THE ABOVE. THE ORDER MUST NOT BE INCREASED OR AMENDED WITHOUT AUTHORISATION FROM WEST COUNTRY MILK LTD ---- </br></br>---- IF POSSIBLE PLEASE INCLUDE PURCHASE ORDER REFFERENCE ON DELIVERY NOTE----</br></br>"
                        mOrder.Notes &= "<b style='color:red;'>---- IMPORTANT ALLERGEN CONTROL - No Brand Substitutes Allowed on Bread Products ----</b></br></br>"
                    End If
                    lTextAttachments = PrepareOrderHTML(mOrder, "", , True)
                    If Not lTextAttachments Is Nothing Then
                        If _Test_Mode Then
                            If String.IsNullOrEmpty(lTestEmail) Then lTestEmail = "victor@wcmilk.co.uk"
                            If _EmailOrder(mOrder, lTestEmail, "", pFileAttachemnt, lTextAttachments, True, mOrder.ServingAccNum_Produce) Then
                                mOrder.UpdateDateEmailed(Now, True)
                                MyEventLog.WriteEntry(lOrderInfo & " emailed To: " & lTestEmail, EventLogEntryType.Information, GetEventID)
                            End If
                        Else
                            If _EmailOrder(mOrder, mOrder.DepotEmail_Produce, mOrder.DepotAltEmail_Produce, pFileAttachemnt, lTextAttachments, , mOrder.ServingAccNum_Produce) Then
                                mOrder.UpdateDateEmailed(Now, True)
                                MyEventLog.WriteEntry(lOrderInfo & " emailed To: " & mOrder.DepotEmail_Produce & " Cc: " & mOrder.DepotAltEmail_Produce, EventLogEntryType.Information, GetEventID)
                                '_EmailServiceMessage(Date.Now & " " & "Order Number: " & _OrderNum & " for " & _AccNum & " emailed To: " & mOrder.DepotEmail_Produce & " Cc: " & mOrder.DepotAltEmail_Produce)
                            End If
                        End If
                    End If
                End If
                If mOrder.MilkItems.Count > 0 OrElse mOrder.DepotId_Produce = 0 Then
                    If mOrder.HOId = 160973922 OrElse mOrder.HOId = 160974269 OrElse mOrder.HOId = 160974511 Then
                        'Add this line in the notes for all interserve site
                        mOrder.Notes = If(mOrder.Notes.Contains("</br></br></br>---- PLEASE SUPPLY ONLY RED TRACTOR MILK FOR THIS SITE ---- </br> </br>---- PLEASE SUPPLY ONLY THE ABOVE. THE ORDER MUST NOT BE INCREASED OR AMENDED WITHOUT AUTHORISATION FROM WEST COUNTRY MILK LTD ---- </br></br>---- IF POSSIBLE PLEASE INCLUDE PURCHASE ORDER REFFERENCE ON DELIVERY NOTE----</br></br>"), mOrder.Notes, mOrder.Notes + "</br></br></br>---- PLEASE SUPPLY ONLY RED TRACTOR MILK FOR THIS SITE ---- </br></br> ---- PLEASE SUPPLY ONLY THE ABOVE. THE ORDER MUST NOT BE INCREASED OR AMENDED WITHOUT AUTHORISATION FROM WEST COUNTRY MILK LTD ---- </br></br>---- IF POSSIBLE PLEASE INCLUDE PURCHASE ORDER REFFERENCE ON DELIVERY NOTE----</br></br>")
                    ElseIf mOrder.HOId = 245739642 OrElse mOrder.HOId = 245740000 OrElse mOrder.Depot.Contains("Hovis") Then 'ISS1 and ISS2
                        'Add this line in the notes for all ISS sites
                        mOrder.Notes = "</br></br></br> ---- PLEASE SUPPLY ONLY THE ABOVE. THE ORDER MUST NOT BE INCREASED OR AMENDED WITHOUT AUTHORISATION FROM WEST COUNTRY MILK LTD ---- </br></br>---- IF POSSIBLE PLEASE INCLUDE PURCHASE ORDER REFFERENCE ON DELIVERY NOTE----</br></br>"
                        mOrder.Notes &= "<b style='color:red;'>---- IMPORTANT ALLERGEN CONTROL - No Brand Substitutes Allowed on Bread Products ----</b></br></br>"
                    End If
                    lTextAttachments = PrepareOrderHTML(mOrder, "")
                    If Not lTextAttachments Is Nothing Then
                        If _Test_Mode Then
                            If String.IsNullOrEmpty(lTestEmail) Then lTestEmail = "victor@wcmilk.co.uk"
                            If _EmailOrder(mOrder, lTestEmail, "", pFileAttachemnt, lTextAttachments, True, mOrder.ServingAccNum) Then
                                mOrder.UpdateDateEmailed(Now)
                                MyEventLog.WriteEntry(lOrderInfo & " emailed To: " & lTestEmail, EventLogEntryType.Information, GetEventID)
                            End If
                        Else
                            If _EmailOrder(mOrder, mOrder.DepotEmail, mOrder.DepotAltEmail, pFileAttachemnt, lTextAttachments, , mOrder.ServingAccNum) Then
                                'If mOrder.BuyingGroupID = 217 Then
                                '    'strip time part to indicate that order has not been emailed to depot but redirected to "orders@wcmilk.co.uk"
                                '    mOrder.UpdateDateEmailed(Now.Date)
                                'Else
                                mOrder.UpdateDateEmailed(Now)
                                'End If

                                MyEventLog.WriteEntry(lOrderInfo & " emailed To: " & mOrder.DepotEmail & " Cc: " & mOrder.DepotAltEmail, EventLogEntryType.Information, GetEventID)
                                '_EmailServiceMessage(Date.Now & " " & "Order Number: " & _OrderNum & " for " & _AccNum & " emailed To: " &  mOrder.DepotEmail & " Cc: " & mOrder.DepotAltEmail)
                            End If
                        End If
                    End If
                End If
            End If


            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString & vbNewLine & lOrderInfo, True)
        Finally
            mOrder = Nothing
        End Try

        Return False
    End Function

    Private Function EmailStandingOrder(ByVal pOrderID As Integer) As Boolean
        Dim lTextAttachments As Dictionary(Of String, System.Text.StringBuilder) _
        , lTestEmail As String = My.Settings.OrderTo _
        , lOrderInfo As String = String.Empty

        Try
            mOrder = New COrderHeader()
            mOrder.RetrieveStandingOrder(pOrderID)
            mOrder.IsStandingOrder = True

            lOrderInfo = "Standing Order Number: " & mOrder.OrderNum & " for " & mOrder.CustAcc.Trim
            If mOrder.IsSOCancelled Then
                lOrderInfo &= " Cancelled"
                mOrder.Notes = "STANDING ORDER HAS BEEN CANCELLED WITH EFFECT FROM " & mOrder.SOLastDeliveryDate.ToString("ddd dd MMM yyyy")
            ElseIf mOrder.IsSuspended Then
                lOrderInfo &= " Suspended"
                mOrder.Notes = "STANDING ORDER HAS BEEN SUSPENDED WITH EFFECT FROM " & mOrder.DateEffective.ToString("ddd dd MMM yyyy")
            End If

            lTextAttachments = PrepareStandingOrderHTML(mOrder)
            If Not lTextAttachments Is Nothing Then
                If _Test_Mode Then
                    If String.IsNullOrEmpty(lTestEmail) Then lTestEmail = "victor@wcmilk.co.uk"
                    If _EmailOrder(mOrder, lTestEmail, "", Nothing, lTextAttachments, True) Then
                        mOrder.UpdateDateEmailed(Now)
                        MyEventLog.WriteEntry("TEST " & lOrderInfo & " emailed To: " & lTestEmail, EventLogEntryType.Information, GetEventID)
                        Return True
                    End If
                Else
                    If _EmailOrder(mOrder, mOrder.DepotEmail, mOrder.DepotAltEmail, Nothing, lTextAttachments) Then

                        mOrder.UpdateDateEmailed(Now)

                        MyEventLog.WriteEntry(lOrderInfo & " emailed To: " & mOrder.DepotEmail & " Cc: " & mOrder.DepotAltEmail, EventLogEntryType.Information, GetEventID)
                        '_EmailServiceMessage(Date.Now & " " & "Order Number: " & _OrderNum & " for " & _AccNum & " emailed To: " & mOrder.DepotEmail & " Cc: " & mOrder.DepotAltEmail)
                        Return True
                    End If
                End If
            End If

            mOrder = Nothing

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString & vbNewLine & lOrderInfo, True)
        End Try
        Return False
    End Function

    Private Function _EmailOrder(ByVal pOrder As COrderHeader, ByVal pTo As String, ByVal pCc As String, ByVal pExcelAttachments As ArrayList, ByVal pTextAttachments As Dictionary(Of String, System.Text.StringBuilder) _
                            , Optional ByVal pTest As Boolean = False, Optional pServingCode As String = "") As Boolean
        Dim oMessage As MailMessage
        Dim idx As Integer
        Dim sTest As String = " "
        Dim sGreeting As String = "Good Afternoon,"
        Dim astrCCs As String() = My.Settings.OrderTo.Split(";")
        Dim lSubject As String = "<!>"
        Dim lCHandCo As String = ""
        Try
            If String.IsNullOrEmpty(pTo) Then Return False

            If Now.Hour < 12 Then
                sGreeting = "Good Morning,"
            ElseIf Now.Hour > 17 Then
                sGreeting = "Good Evening,"
            End If

            If pTest Then
                sTest = "TEST - PLEASE IGNORE  "
            End If

            'If pOrder.EDIOrdersFlag AndAlso Not pOrder.IsStandingOrder Then
            '    sTest &= " (Depot set up for EDI Ordering)"
            '    pTo = "sentorders@wcmilk.co.uk"
            'End If

            'If pOrder.BuyingGroupID = 217 Then
            '    'redirect CH&Co orders
            '    lCHandCo = "CH & CO "
            '    pTo = "orders@wcmilk.co.uk"
            '    pCc = ""
            'End If

            oMessage = New MailMessage()
            lSubject = lCHandCo & pOrder.Customer & " \ " & If(pServingCode = "", pOrder.ServingAccNum, pServingCode) & "  To: " & pTo & "  Cc: " & pCc

            With oMessage
                If pTo.IndexOf("@") > -1 Then
                    .To.Add(New MailAddress(pTo))
                    '.To.Add(New MailAddress("daniel@wcmilk.co.uk"))
                    If pCc.IndexOf("@") > -1 Then
                        .CC.Add(New MailAddress(pCc))
                    End If
                ElseIf pCc.IndexOf("@") > -1 Then
                    .To.Add(New MailAddress(pCc))
                End If

                If pTest Then
                    For Each sWCM_Cc As String In astrCCs
                        If sWCM_Cc.IndexOf("@") > -1 Then
                            .CC.Add(New MailAddress(sWCM_Cc))
                        End If
                    Next
                Else
                    If pOrder.EDIOrdersFlag = False Then
                        .Bcc.Add(New MailAddress("sentorders@wcmilk.co.uk"))
                    End If
                End If

                If My.Settings.orderfrom <> "" Then
                    .From = New MailAddress(My.Settings.orderfrom)
                Else
                    .From = New MailAddress("orders@wcmilk.co.uk")
                    ' .From = New MailAddress(modCurrentUser.Email)
                End If

                If pOrder.IsStandingOrder Then
                    If pOrder.IsSOCancelled Then
                        .Subject = sTest & "CANCELLED Standing Order ref. " & pOrder.OrderNum & " for  " & pOrder.Customer & " \ " & If(pServingCode = "", pOrder.ServingAccNum, pServingCode) & "  effective from " & DateAdd(DateInterval.Day, 1, pOrder.SOLastDeliveryDate).ToString("ddd dd MMM yyyy")
                        .Body = sTest & vbNewLine & vbNewLine & "Please CANCEL Standing Order below: " & vbNewLine & vbNewLine
                    ElseIf pOrder.IsSuspended Then
                        .Subject = sTest & "SUSPENDED Standing Order ref. " & pOrder.OrderNum & " for  " & pOrder.Customer & " \ " & If(pServingCode = "", pOrder.ServingAccNum, pServingCode) & "  effective from " & pOrder.DateEffective.ToString("ddd dd MMM yyyy")
                        .Body = sTest & vbNewLine & vbNewLine & "Please SUSPEND Standing Order below: " & vbNewLine & vbNewLine
                    Else
                        .Subject = sTest & "Standing Order ref. " & pOrder.OrderNum & " for  " & pOrder.Customer & " \ " & If(pServingCode = "", pOrder.ServingAccNum, pServingCode) & "  effective from " & pOrder.DateEffective.ToString("ddd dd MMM yyyy")
                        If Nz(Of Date)(pOrder.SOLastDeliveryDate, Date.MinValue) <> Date.MinValue Then
                            .Subject &= " to " & pOrder.SOLastDeliveryDate.ToString("ddd dd MMM yyyy")
                        End If
                        .Body = sTest & vbNewLine & vbNewLine & "Please see Standing Order below: " & vbNewLine & vbNewLine
                    End If
                Else
                    If pOrder.IsSoAmendment Then
                        .Subject = sTest & "Amended Order for  " & pOrder.Customer & " \ " & If(pServingCode = "", pOrder.ServingAccNum, pServingCode) & "  delivery date " & pOrder.DeliveryDate.ToString("ddd dd MMM yyyy")
                        .Body = sTest & vbNewLine & vbNewLine & "Please see Amended Order below (for this delivery date only): " & vbNewLine & vbNewLine
                    Else
                        .Subject = sTest & "Order for " & pOrder.Customer & " \ " & If(pServingCode = "", pOrder.ServingAccNum, pServingCode) & " \ " & pOrder.DeliveryDate.ToString("ddd dd MMM yyyy")
                        .Body = sTest & vbNewLine & sGreeting & vbNewLine & vbNewLine & "Please find Order for Delivery Date " &
                                pOrder.DeliveryDate.ToString("ddd dd MMM yyyy") & " for:  " & vbNewLine & vbNewLine
                    End If
                End If

                lSubject = .Subject

                If Not pTextAttachments Is Nothing Then
                    For Each lTextAttachment As KeyValuePair(Of String, System.Text.StringBuilder) In pTextAttachments
                        '.Body &= vbNewLine & lTextAttachment.Key & vbNewLine
                        .Body &= vbNewLine
                        .Body &= "</pre>" & lTextAttachment.Value.ToString & "<pre>" & vbNewLine
                    Next
                End If

                .Body &= vbNewLine & "If you have any queries please contact us on the number below." & vbNewLine
                .Body &= vbNewLine
                .Body &= "Regards," & vbNewLine
                .Body &= vbNewLine
                .Body &= "West Country Milk" & vbNewLine & vbNewLine
                .Body &= "Otter Building" & vbNewLine
                .Body &= "Grenadier Road" & vbNewLine
                .Body &= "Exeter Business Park" & vbNewLine
                .Body &= "Devon" & vbNewLine
                .Body &= "EX1 3QF" & vbNewLine
                .Body &= "Tel. 01392 350000" & vbNewLine & vbNewLine

                .IsBodyHtml = True
                .Body = "<html><body><pre>" & .Body & "</pre></body></html>"

                If Not pExcelAttachments Is Nothing Then
                    For idx = 0 To pExcelAttachments.Count - 1
                        If System.IO.File.Exists(pExcelAttachments(idx).ToString) Then
                            'creating an instance of MailAttachment class and specifying the location of attachment
                            'adding the attachment to mailMessage  
                            .Attachments.Add(New Attachment(pExcelAttachments(idx).ToString))
                        End If
                    Next idx
                End If
            End With

            Dim oSMTP As New SmtpClient()

            'oSMTP.Host  defined in app.config<system.net><mailSettings> section
            'oSMTP.DeliveryMethod = SmtpDeliveryMethod.Network
            'oSMTP.DeliveryMethod = SmtpDeliveryMethod.PickupDirectoryFromIis

            'oSMTP.UseDefaultCredentials = False
            'oSMTP.Credentials = New System.Net.NetworkCredential("WCMAdmin", "M1lkAdm1nP@55!") '("WCMApplication", "M1lkyW4y!")

            lSubject &= " a)"

            oSMTP.UseDefaultCredentials = True
            lSubject &= " b)"
            oSMTP.Send(oMessage)

            lSubject &= " c)"

            Return True

        Catch ex As SmtpFailedRecipientsException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString & " 1: " & lSubject, EventLogEntryType.Error, GetEventID)
        Catch ex As SmtpFailedRecipientException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString & " 2: " & lSubject, EventLogEntryType.Error, GetEventID)
        Catch ex As SmtpException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString & " 3: " & lSubject, EventLogEntryType.Error, GetEventID)
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString & " 4: " & lSubject, EventLogEntryType.Error, GetEventID)
        Finally
            oMessage = Nothing
        End Try
        Return False
    End Function

    Private Function PrepareOrderHTML(ByVal pOrder As COrderHeader, ByVal pBDProductsOnly As String, Optional ByVal pIsAmendment As Boolean = False, Optional pIsProduce As Boolean = False) As Dictionary(Of String, System.Text.StringBuilder)
        Dim lTextAttachments As New Dictionary(Of String, System.Text.StringBuilder) _
        , lEmailBody As New System.Text.StringBuilder _
        , idx As Integer _
        , lOrderList As List(Of COrderDetail) = pOrder.MilkItems _
        , oLine As COrderDetail _
        , lSign As String = "" _
        , lShowSupplierCodes As Boolean = pOrder.ShowProductCodes_Supplier

        If pIsProduce Then
            lOrderList = pOrder.ProduceItems
            lShowSupplierCodes = pOrder.ShowProductCodes_Supplier_Produce
        End If
        With lEmailBody
            .Append("<table cellpadding=""1"" cellspacing=""1"" border=""0"">")
            .Append("<tr>")
            .Append("<td> Order Number:   ") : .Append("</td>")
            .Append("<td>") : .Append(pOrder.OrderNum) : .Append("</td>")
            .Append("</tr>")
            .Append("<tr>")
            .Append("<td> Supplier Acct Nos:   ") : .Append("</td>")
            .Append("<td>")
            If (pIsProduce) Then .Append(pOrder.ServingAccNum_Produce) Else .Append(pOrder.ServingAccNum)
            .Append("</td>")
            .Append("</tr>")
            .Append("<tr>")
            .Append("<td> Site name:   ") : .Append("</td>")
            .Append("<td>") : .Append(pOrder.Customer) : .Append(",  ").Append(pOrder.Address1) : .Append("</td>")
            .Append("</tr>")
            .Append("<tr>")
            .Append("<td> PostCode:   ") : .Append("</td>")
            .Append("<td>") : .Append(pOrder.Postcode) : .Append("</td>")
            .Append("</tr>")
            .Append("<tr>")
            .Append("</tr>")
            .Append("<tr>")
            .Append("<td>  Delivery Date:   ") : .Append("</td>")
            .Append("<td>") : .Append(pOrder.DeliveryDate.ToString("ddd dd MMM yyyy")) : .Append("</td>")
            .Append("</tr>")
            .Append("</table>")
        End With
        lTextAttachments.Add("header", lEmailBody)

        If lOrderList.Count > 0 Then
            lEmailBody = New System.Text.StringBuilder
            With lEmailBody
                .Append("<table cellpadding=""1"" cellspacing=""1"" border=""1"">")
                .Append("<tr>")
                If lShowSupplierCodes Then
                    .Append("<td nowrap=""nowrap"">Code</td>")
                End If
                .Append("<td nowrap=""nowrap"">Product</td>")
                If pIsAmendment Then
                    .Append("<td nowrap=""nowrap"" align=""center"">Standing Order</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">Amendment</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">Quantity</td>")
                Else
                    .Append("<td nowrap=""nowrap"" align=""center"">Quantity</td>")
                End If

                .Append("</tr>")

                For idx = 0 To lOrderList.Count - 1
                    oLine = DirectCast(lOrderList.Item(idx), COrderDetail)
                    If pBDProductsOnly = "BD" Then
                        If oLine.ProductId < 5385 OrElse oLine.ProductId > 5388 Then
                            'exclude all products except BD1000 BD1001 BD1002 BD1003 which are delivered by particular supplier
                            Continue For
                        End If
                    ElseIf oLine.ProductId >= 5385 AndAlso oLine.ProductId <= 5388 Then
                        'exclude products BD1000 BD1001 BD1002 BD1003 as they are delivered by particular supplier
                        Continue For
                    End If

                    If oLine.ProductId = 6056 OrElse oLine.ProductId = 6314 OrElse oLine.ProductId = 6315 Then 'AndAlso Not (pOrder.DepotId = 129 OrElse pOrder.DepotId = 341 OrElse pOrder.DepotId = 369 OrElse pOrder.DepotId = 371) Then 'For all depots bar Wells Farm
                        'exclude product WCM Delivery Charge 
                        Continue For
                    End If

                    If pIsAmendment Then
                        .Append("<tr>")
                        If lShowSupplierCodes Then
                            If String.IsNullOrEmpty(oLine.ProductCode_Supplier) Then
                                .Append("<td>") : .Append("</td>")
                                .Append("<td>" & oLine.Product) : .Append("</td>")
                            Else
                                .Append("<td>" & oLine.ProductCode_Supplier) : .Append("</td>")
                                .Append("<td>" & oLine.Product_Supplier) : .Append("</td>")
                            End If
                        Else
                            .Append("<td>" & oLine.Product) : .Append("</td>")
                        End If
                        .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.SOQty, Int16)).ToString() : .Append("</td>")
                        If oLine.Qty - oLine.SOQty > 0 Then lSign = "+" Else lSign = ""
                        .Append("<td nowrap=""nowrap"" align=""center"">" & lSign & CType(oLine.Qty - oLine.SOQty, Int16)).ToString() : .Append("</td>")
                        .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Qty, Int16)).ToString() : .Append("</td>")
                        .Append("</tr>")
                    ElseIf oLine.Qty <> 0 Then
                        .Append("<tr>")
                        If lShowSupplierCodes Then
                            If String.IsNullOrEmpty(oLine.ProductCode_Supplier) Then
                                .Append("<td>") : .Append("</td>")
                                .Append("<td>" & oLine.Product) : .Append("</td>")
                            Else
                                .Append("<td>" & oLine.ProductCode_Supplier) : .Append("</td>")
                                .Append("<td>" & oLine.Product_Supplier) : .Append("</td>")
                            End If
                        Else
                            .Append("<td>" & oLine.Product) : .Append("</td>")
                        End If
                        .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Qty, Int16)).ToString() : .Append("</td>")
                        .Append("</tr>")
                    End If
                Next
                .Append("</table>")
            End With
            lTextAttachments.Add("details", lEmailBody)
        End If

        If pOrder.OrderNum.StartsWith("WCMWA") AndAlso lOrderList.Count > 0 AndAlso pOrder.Notes.Contains("cancel") Then
            pOrder.Notes = ""
        End If
        If pOrder.Notes <> "" Then
            lEmailBody = New System.Text.StringBuilder
            With lEmailBody
                .Append("<table cellpadding=""1"" cellspacing=""1"" border=""0"">")
                .Append("<tr>")
                .Append("<td>" & pOrder.Notes) : .Append("</td>")
                .Append("</tr>")
                .Append("</table>")
            End With

            lTextAttachments.Add("notes", lEmailBody)
        End If
        Return lTextAttachments
    End Function


    Private Function PrepareStandingOrderHTML(ByVal pStandingOrder As COrderHeader) As Dictionary(Of String, System.Text.StringBuilder)
        Dim lTextAttachments As New Dictionary(Of String, System.Text.StringBuilder)
        Dim lHeader As New System.Text.StringBuilder
        Dim lDetail As System.Text.StringBuilder
        Dim lNote As System.Text.StringBuilder
        Dim idx As Integer
        Dim oLine As COrderDetail


        With lHeader
            .Append("<table cellpadding=""1"" cellspacing=""1"" border=""0"">")
            .Append("<tr>")
            .Append("<td> Supplier Acct No:    ") : .Append("</td>")
            .Append("<td>") : .Append(pStandingOrder.ServingAccNum) : .Append("</td>")
            .Append("</tr>")
            .Append("<tr>")
            .Append("<td> Site name:   ") : .Append("</td>")
            .Append("<td>") : .Append(pStandingOrder.Customer) : .Append(",  ").Append(pStandingOrder.Address1) : .Append("</td>")
            .Append("</tr>")
            .Append("<tr>")
            .Append("<td> PostCode:   ") : .Append("</td>")
            .Append("<td>") : .Append(pStandingOrder.Postcode) : .Append("</td>")
            .Append("</tr>")
            .Append("<tr>")
            If Not pStandingOrder.IsSOCancelled Then
                If Nz(Of Date)(pStandingOrder.SOLastDeliveryDate, Date.MinValue) <> Date.MinValue Then
                    .Append("<td>  Dates Effective:  ") : .Append("</td>")
                    .Append("<td>") : .Append(pStandingOrder.DateEffective.ToString("ddd dd MMM yyyy")) : .Append(" - ") : .Append(pStandingOrder.SOLastDeliveryDate.ToString("ddd dd MMM yyyy")) : .Append("</td>")
                Else
                    .Append("<td>  Date Effective:   ") : .Append("</td>")
                    .Append("<td>") : .Append(pStandingOrder.DateEffective.ToString("ddd dd MMM yyyy")) : .Append("</td>")
                End If
            End If
            .Append("</tr>")
            .Append("<tr>")
            .Append("<td>") : .Append(" ") : .Append("</td>")
            .Append("</tr>")
            .Append("</table>")
        End With
        lTextAttachments.Add("header", lHeader)

        lHeader = New System.Text.StringBuilder
        With lHeader
            .Append("<table cellpadding=""1"" cellspacing=""1"" border=""0"">")
            .Append("<tr>")
            .Append("<td> Standing Order:   ") : .Append(pStandingOrder.OrderNum) : .Append("</td>")
            .Append("</tr>")
            .Append("</table>")
        End With
        lTextAttachments.Add("header2", lHeader)


        If pStandingOrder.Count > 0 AndAlso Not mOrder.IsSOCancelled AndAlso Not mOrder.IsSuspended Then
            lDetail = New System.Text.StringBuilder
            With lDetail
                .Append("<table cellpadding=""1"" cellspacing=""1"" border=""1"">")
                .Append("<tr>")
                .Append("<td nowrap=""nowrap"">Product</td>")
                .Append("<td nowrap=""nowrap"" align=""center"">Sun</td>")
                .Append("<td nowrap=""nowrap"" align=""center"">Mon</td>")
                .Append("<td nowrap=""nowrap"" align=""center"">Tue</td>")
                .Append("<td nowrap=""nowrap"" align=""center"">Wed</td>")
                .Append("<td nowrap=""nowrap"" align=""center"">Thu</td>")
                .Append("<td nowrap=""nowrap"" align=""center"">Fri</td>")
                .Append("<td nowrap=""nowrap"" align=""center"">Sat</td>")
                .Append("</tr>")

                For idx = 0 To pStandingOrder.Count - 1
                    oLine = DirectCast(pStandingOrder.Item(idx), COrderDetail)
                    .Append("<tr>")
                    .Append("<td>" & oLine.Product) : .Append("</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Sun, Int16)).ToString() : .Append("</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Mon, Int16)).ToString() : .Append("</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Tue, Int16)).ToString() : .Append("</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Wed, Int16)).ToString() : .Append("</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Thu, Int16)).ToString() : .Append("</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Fri, Int16)).ToString() : .Append("</td>")
                    .Append("<td nowrap=""nowrap"" align=""center"">" & CType(oLine.Sat, Int16)).ToString() : .Append("</td>")
                    .Append("</tr>")
                Next
                .Append("</table>")
            End With
            lTextAttachments.Add("details", lDetail)
        End If

        If pStandingOrder.Notes <> "" Then
            lNote = New System.Text.StringBuilder
            With lNote
                .Append("<table cellpadding=""1"" cellspacing=""1"" border=""0"">")
                .Append("<tr>")
                .Append("<td>" & pStandingOrder.Notes) : .Append("</td>")
                .Append("</tr>")
                .Append("</table>")
            End With

            lTextAttachments.Add("notes", lNote)
        End If
        Return lTextAttachments
    End Function

    Private Function _EmailServiceMessage(ByVal pMessage As String, Optional ByVal pError As Boolean = False) As Boolean
        Dim oMessage As MailMessage
        Dim sTest As String = " ", sError As String = " "
        Try
            Dim astrCCs As String() = My.Settings.OrderTo.Split(";")
            Dim astrTo As String() = My.Settings.Email_Success.Split(";")

            If _Test_Mode Then
                sTest = "TEST "
            End If

            If pError Then
                sError = "ERROR "
                astrTo = My.Settings.Email_Error.Split(";")
            End If

            oMessage = New MailMessage()
            With oMessage
                For Each sWCM_To As String In astrTo
                    If sWCM_To.IndexOf("@") > -1 Then
                        .To.Add(New MailAddress(sWCM_To))
                    End If
                Next
                If .To.Count = 0 Then
                    .To.Add(New MailAddress("victor@wcmilk.co.uk"))
                End If

                If _Test_Mode Then
                    For Each sWCM_Cc As String In astrCCs
                        If sWCM_Cc.IndexOf("@") > -1 Then
                            .CC.Add(New MailAddress(sWCM_Cc))
                        End If
                    Next
                End If
                If pError AndAlso COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.Order_Upload_1000 Then
                    .CC.Add(New MailAddress("naomi@wcmilk.co.uk"))
                    .CC.Add(New MailAddress("katie@wcmilk.co.uk"))
                End If

                If My.Settings.orderfrom <> "" Then
                    .From = New MailAddress(My.Settings.orderfrom)
                Else
                    .From = New MailAddress("orders@wcmilk.co.uk")
                End If

                .Subject = sTest & sError & " WCMOrdering"
                If pError Then
                    .Body = sError & vbNewLine & vbNewLine
                End If
                .Body &= pMessage & vbNewLine & vbNewLine

                .IsBodyHtml = False

            End With

            Dim oSMTP As New SmtpClient()

            'oSMTP.Host  defined in app.config<system.net><mailSettings> section
            'oSMTP.DeliveryMethod = SmtpDeliveryMethod.Network
            'oSMTP.DeliveryMethod = SmtpDeliveryMethod.PickupDirectoryFromIis

            'oSMTP.UseDefaultCredentials = False
            'oSMTP.Credentials = New System.Net.NetworkCredential("WCMAdmin", "M1lkAdm1nP@55!") '("WCMApplication", "M1lkyW4y!")

            oSMTP.UseDefaultCredentials = True
            oSMTP.Send(oMessage)

            Return True

        Catch ex As SmtpFailedRecipientsException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Catch ex As SmtpFailedRecipientException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Catch ex As SmtpException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Finally
            oMessage = Nothing
        End Try
        Return False
    End Function

    Private Function _EmailServiceMessage(ByVal pSubject As String, ByVal pMessage As String, pTo As String, pCC As String) As Boolean
        Dim oMessage As MailMessage

        Try
            Dim astrCCs As String() = pCC.Split(";")
            Dim astrTo As String() = pTo.Split(";")


            oMessage = New MailMessage()
            With oMessage
                For Each sWCM_To As String In astrTo
                    If sWCM_To.IndexOf("@") > -1 Then
                        .To.Add(New MailAddress(sWCM_To))
                    End If
                Next

                .From = New MailAddress("orders@wcmilk.co.uk")

                .Subject = pSubject

                .Body &= pMessage & vbNewLine & vbNewLine

                .IsBodyHtml = False

            End With

            Dim oSMTP As New SmtpClient()

            oSMTP.UseDefaultCredentials = True
            oSMTP.Send(oMessage)

            Return True

        Catch ex As SmtpFailedRecipientsException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Catch ex As SmtpFailedRecipientException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Catch ex As SmtpException
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString)
        Finally
            oMessage = Nothing
        End Try
        Return False
    End Function

    Public Function Email_Generic(ByVal pTo As String, ByVal pCc As String, ByVal pSubject As String, ByVal pMessage As String, ByVal pAttachFile As String, pFrom As String, ByVal pHTML As Boolean, pBcc As String) As Boolean
        Dim Message As MailMessage
        Dim lLineBreak As String = vbNewLine
        Try
            Message = New MailMessage()
            With Message

                If _Test_Mode Then
                    pTo = "victor@wcmilk.co.uk"
                    pSubject = "TEST " & pSubject
                    pMessage = "TEST.  PLEASE IGNORE!" & vbNewLine & vbNewLine & pMessage
                End If
                For Each sTo As String In pTo.Split(";")
                    If sTo.IndexOf("@") > -1 Then
                        .To.Add(New MailAddress(sTo))
                    End If
                Next

                For Each sCc As String In pCc.Split(";")
                    If sCc.IndexOf("@") > -1 Then
                        .CC.Add(New MailAddress(sCc))
                    End If
                Next
                For Each sBcc As String In pBcc.Split(";")
                    If sBcc.IndexOf("@") > -1 Then
                        .Bcc.Add(New MailAddress(sBcc))
                    End If
                Next

                If pFrom.IndexOf("@") = -1 Then
                    If Not String.IsNullOrEmpty(My.Settings.orderfrom) Then
                        .From = New MailAddress(My.Settings.orderfrom)
                    Else
                        .From = New MailAddress("NoReply@wcmilk.co.uk")
                    End If
                Else
                    .From = New MailAddress(pFrom)
                End If

                .Subject = pSubject
                .Body = pMessage

                .IsBodyHtml = pHTML

                If Not String.IsNullOrEmpty(pAttachFile) AndAlso System.IO.File.Exists(pAttachFile) Then
                    'creating an instance of MailAttachment class and specifying the location of attachment
                    'adding the attachment to mailMessage  
                    .Attachments.Add(New Attachment(pAttachFile))
                End If
            End With

            Dim oSMTP As New SmtpClient
            'oSMTP.Host  defined in app.config<system.net><mailSettings> section
            oSMTP.UseDefaultCredentials = True
            oSMTP.Send(Message)

            Return True

        Catch ex As SmtpFailedRecipientsException
            _EmailServiceMessage("Failed to send email to:" & pTo, True)
        Catch ex As SmtpException
            _EmailServiceMessage("Failed to connect to mail server", True)
        Catch ex As Exception
            _EmailServiceMessage(ex.Message, True)
        Finally
            Message = Nothing

        End Try
        Return False
    End Function


#End Region


#Region "Email WEBAPP Orders "

    Private Function GetOrdersNotEmailed() As Boolean
        Dim l_DB As DB = Nothing _
        , Cmd As SqlClient.SqlCommand = Nothing _
        , Param As SqlClient.SqlParameter _
        , Rdr As SqlClient.SqlDataReader = Nothing _
        , lSupplier As Integer = 0

        Try
            If COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.PushEmailWebApp_500 Then
                lSupplier = 500
            ElseIf COrderHeader.BuyerSequence = COrderHeader.BuyerSeq.PushEmailP2P_999 Then
                lSupplier = 999
            Else
                Return False
            End If
            _Orders_ToEmail = New Stack(Of Integer)
            l_DB = New DB
            l_DB.Open()

            Cmd = New SqlClient.SqlCommand("p_bulk_orders_get", l_DB.Connection)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                'Create Parameters
                Param = .Parameters.Add("@order_batch", SqlDbType.VarChar, 20)
                Param.Value = DBNull.Value

                Param = .Parameters.Add("@supplier", SqlDbType.Int)
                Param.Value = lSupplier

                Param = .Parameters.Add("@order_date", SqlDbType.Date)
                Param.Value = Now.Date

                Rdr = .ExecuteReader()
            End With

            With Rdr
                If .HasRows Then
                    Do While .Read
                        _Orders_ToEmail.Push(CType(Rdr("order_id"), Integer))
                    Loop
                End If
                .Close()
            End With
            Cmd.Dispose()
            l_DB.Close()
            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
        Return False
    End Function

    Private Function GetStandingNotEmailed() As Boolean
        Dim l_DB As DB = Nothing _
        , Cmd As SqlClient.SqlCommand = Nothing _
        , Rdr As SqlClient.SqlDataReader = Nothing _
        , lSupplier As Integer = 0

        Try

            _SO_ToEmail = New Stack(Of Integer)
            l_DB = New DB
            l_DB.Open()

            Cmd = New SqlClient.SqlCommand("p_bulk_email_so_get", l_DB.Connection)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                'No Parameters

                Rdr = .ExecuteReader()
            End With

            With Rdr
                If .HasRows Then
                    Do While .Read
                        _SO_ToEmail.Push(CType(Rdr("order_id"), Integer))
                    Loop
                End If
                .Close()
            End With
            Cmd.Dispose()
            l_DB.Close()
            Return True

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
        Return False
    End Function

    Private Function PushEmailOrders() As Integer '' ByRef pFailedCount As Integer, Optional ByVal bTest As Boolean = False) As Integer
        Dim lEmailsSent As Integer = 0 _
        , lFailedOrders As New Stack(Of Integer) _
        , lOrderId As Integer

        Try

            Do While _Orders_ToEmail.Count > 0
                lOrderId = _Orders_ToEmail.Pop
                If EmailOrder(lOrderId) Then
                    lEmailsSent += 1
                Else
                    lFailedOrders.Push(lOrderId)
                End If
                System.Threading.Thread.Sleep(100)
            Loop

            '  pFailedCount = lFailedOrders.Count


            Return lEmailsSent

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
        Return False
    End Function

    Private Function PushEmailStandingOrders() As Integer
        Dim lEmailsSent As Integer = 0 _
        , lFailedOrders As New Stack(Of Integer) _
        , lOrderId As Integer

        Try

            Do While _SO_ToEmail.Count > 0
                lOrderId = _SO_ToEmail.Pop
                If EmailStandingOrder(lOrderId) Then
                    lEmailsSent += 1
                Else
                    lFailedOrders.Push(lOrderId)
                End If
                System.Threading.Thread.Sleep(100)
            Loop

            '  pFailedCount = lFailedOrders.Count


            Return lEmailsSent

        Catch ex As Exception
            MyEventLog.WriteEntry("ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, EventLogEntryType.Error, GetEventID)
            _EmailServiceMessage(Date.Now & " " & "ERROR: " & ex.Message & " in " & _MeName & "." & System.Reflection.MethodInfo.GetCurrentMethod.ToString, True)
        Finally

        End Try
        Return False
    End Function
#End Region

End Class
