using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace VendorExpiryEmailService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();

            // Configure the existing installers from Designer file
            serviceProcessInstaller1.Account = ServiceAccount.LocalSystem;

            serviceInstaller1.ServiceName = "VendorExpiryEmailService";
            serviceInstaller1.DisplayName = "Vendor Expiry Email Notification Service";
            serviceInstaller1.Description = "Sends email notifications before and after vendor expiry.";
            serviceInstaller1.StartType = ServiceStartMode.Automatic;
        }

        private void serviceProcessInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {
            // Optional: handle post-install logic
        }

        private void serviceInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {
            // Optional: handle post-install logic
        }
    }
}
