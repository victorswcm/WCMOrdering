Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WCMOrdering
    Inherits System.ServiceProcess.ServiceBase

    'UserService overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()

#If DEBUG Then
        Dim myServ As New WCMOrdering()


        myServ.OnTimer(Nothing, Nothing)
        ' here Process is my Service function
		' that will run when my service onstart is call
		' you need to call your own method or function name here instead of Process();
#Else
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New WCMOrdering, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New WCMOrdering}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
#End If
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MyEventLog = New System.Diagnostics.EventLog()
        CType(Me.MyEventLog, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'MyEventLog
        '
        '
        'WCMOrdering
        '
        Me.ServiceName = "WCMOrdering"
        CType(Me.MyEventLog, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Friend WithEvents MyEventLog As System.Diagnostics.EventLog

End Class
