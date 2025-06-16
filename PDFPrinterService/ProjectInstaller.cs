using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace PDFPrinterService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        private ServiceProcessInstaller serviceProcessInstaller;
        private ServiceInstaller serviceInstaller;

        public ProjectInstaller()
        {
            InitializeComponent();
            this.serviceProcessInstaller = new ServiceProcessInstaller();
            this.serviceInstaller = new ServiceInstaller();

            this.serviceProcessInstaller.Account = ServiceAccount.User;

            this.serviceInstaller.ServiceName = "PDFPrinterService";
            this.serviceInstaller.DisplayName = "PDF Printer Service";
            this.serviceInstaller.Description = "A custom service for PDF printing.";
            this.serviceInstaller.StartType = ServiceStartMode.Automatic;

            this.Installers.AddRange(new Installer[] {
                this.serviceProcessInstaller,
                this.serviceInstaller
            });
        }
    }
}
