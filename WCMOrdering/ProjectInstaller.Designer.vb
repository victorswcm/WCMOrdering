<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.WCMOrderingServiceProcessInstaller = New System.ServiceProcess.ServiceProcessInstaller()
        Me.WCMOrderingServiceInstaller = New System.ServiceProcess.ServiceInstaller()
        '
        'WCMOrderingServiceProcessInstaller
        '
        Me.WCMOrderingServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.WCMOrderingServiceProcessInstaller.Password = Nothing
        Me.WCMOrderingServiceProcessInstaller.Username = Nothing
        '
        'WCMOrderingServiceInstaller
        '
        Me.WCMOrderingServiceInstaller.Description = "WCM Ordering Service Apllication (Elior P2P Supplier Integration) "
        Me.WCMOrderingServiceInstaller.DisplayName = "WCM Ordering"
        Me.WCMOrderingServiceInstaller.ServiceName = "WCMOrdering"
        Me.WCMOrderingServiceInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.WCMOrderingServiceProcessInstaller, Me.WCMOrderingServiceInstaller})

    End Sub
    Friend WithEvents WCMOrderingServiceProcessInstaller As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents WCMOrderingServiceInstaller As System.ServiceProcess.ServiceInstaller

End Class
