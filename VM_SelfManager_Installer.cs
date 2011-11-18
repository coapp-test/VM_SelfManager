using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace VM_SelfManager
{
    [RunInstaller(true)]
    class VM_SelfManager_Installer : Installer
    {

        public VM_SelfManager_Installer()
        {
            ServiceProcessInstaller serviceProcessInstaller =
                               new ServiceProcessInstaller();
            ServiceInstaller serviceInstaller = new ServiceInstaller();

            //# Service Account Information

            serviceProcessInstaller.Account = ServiceAccount.LocalSystem;
            serviceProcessInstaller.Username = null;
            serviceProcessInstaller.Password = null;

            //# Service Information

            serviceInstaller.DisplayName = "Hyper-V Virtual Machine Self Management Service";
            serviceInstaller.StartType = ServiceStartMode.Automatic;

            serviceInstaller.ServiceName = "VM_SelfManager";

            this.Installers.Add(serviceProcessInstaller);
            this.Installers.Add(serviceInstaller);

        }

    }
}
