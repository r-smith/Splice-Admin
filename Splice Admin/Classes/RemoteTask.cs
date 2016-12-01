using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;

namespace Splice_Admin.Classes
{
    class RemoteTask
    {
        public static DialogResult TaskRebootComputer(string targetComputer)
        {
            // TaskRebootComputer() uses WMI to reboot the selected computer.
            // It returns a DialogResult which will be used to display the results.
            var dialog = new DialogResult();
            bool didTaskSucceed = false;

            // Setup WMI query.
            var options = new ConnectionOptions();
            if (GlobalVar.UseAlternateCredentials)
            {
                options.Username = GlobalVar.AlternateUsername;
                options.Password = GlobalVar.AlternatePassword;
                options.Authority = $"NTLMDOMAIN:{GlobalVar.AlternateDomain}";
            }
            var scope = new ManagementScope($@"\\{targetComputer}\root\CIMV2", options);
            var query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            var searcher = new ManagementObjectSearcher(scope, query);

            try
            {
                foreach (ManagementObject m in searcher.Get())
                {
                    ManagementBaseObject inParams = m.GetMethodParameters("Win32Shutdown");
                    inParams["Flags"] = 6;
                    ManagementBaseObject outParams = m.InvokeMethod("Win32Shutdown", inParams, null);
                    int returnValue;
                    bool isReturnValid = int.TryParse(outParams["ReturnValue"].ToString(), out returnValue);
                    if (isReturnValid && returnValue == 0)
                        didTaskSucceed = true;
                    break;
                }
            }
            catch
            { }

            if (didTaskSucceed)
            {
                // Reboot command was successful.  Build DialogResult to reflect success.
                dialog.DialogTitle = "Success";
                dialog.DialogBody = $"{targetComputer} is now in the process of rebooting.";
                dialog.DialogIconPath = "/Resources/success-48.png";
                dialog.ButtonIconPath = "/Resources/checkmark-24.png";
                dialog.ButtonText = "OK";
                dialog.IsCancelVisible = false;
            }
            else
            {
                // Failed to reboot computer.  Build DialogResult to reflect failure.
                dialog.DialogTitle = "Error";
                dialog.DialogBody = $"Failed to reboot {targetComputer}.";
                dialog.DialogIconPath = "/Resources/error-48.png";
                dialog.ButtonIconPath = "/Resources/checkmark-24.png";
                dialog.ButtonText = "OK";
                dialog.IsCancelVisible = false;
            }

            return dialog;
        }


        public static DialogResult TaskShutdownComputer(string targetComputer)
        {
            // TaskShutdownComputer() uses WMI to shutdown the selected computer.
            // It returns a DialogResult which will be used to display the results.
            var dialog = new DialogResult();
            bool didTaskSucceed = false;

            // Setup WMI query.
            var options = new ConnectionOptions();
            if (GlobalVar.UseAlternateCredentials)
            {
                options.Username = GlobalVar.AlternateUsername;
                options.Password = GlobalVar.AlternatePassword;
                options.Authority = $"NTLMDOMAIN:{GlobalVar.AlternateDomain}";
            }
            var scope = new ManagementScope($@"\\{targetComputer}\root\CIMV2", options);
            var query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            var searcher = new ManagementObjectSearcher(scope, query);

            try
            {
                foreach (ManagementObject m in searcher.Get())
                {
                    ManagementBaseObject inParams = m.GetMethodParameters("Win32Shutdown");
                    inParams["Flags"] = 5;
                    ManagementBaseObject outParams = m.InvokeMethod("Win32Shutdown", inParams, null);
                    int returnValue;
                    bool isReturnValid = int.TryParse(outParams["ReturnValue"].ToString(), out returnValue);
                    if (isReturnValid && returnValue == 0)
                        didTaskSucceed = true;
                    break;
                }
            }
            catch
            { }

            if (didTaskSucceed)
            {
                // Shutdown command was successful.  Build DialogResult to reflect success.
                dialog.DialogTitle = "Success";
                dialog.DialogBody = $"{targetComputer} is now in the process of shutting down.";
                dialog.DialogIconPath = "/Resources/success-48.png";
                dialog.ButtonIconPath = "/Resources/checkmark-24.png";
                dialog.ButtonText = "OK";
                dialog.IsCancelVisible = false;
            }
            else
            {
                // Failed to shutdown computer.  Build DialogResult to reflect failure.
                dialog.DialogTitle = "Error";
                dialog.DialogBody = $"Failed to shutdown {targetComputer}.";
                dialog.DialogIconPath = "/Resources/error-48.png";
                dialog.ButtonIconPath = "/Resources/checkmark-24.png";
                dialog.ButtonText = "OK";
                dialog.IsCancelVisible = false;
            }

            return dialog;
        }


        public static DialogResult TaskGpupdate(string targetComputer)
        {
            // TaskGpupdate() uses WMI to execute the GPUpdate command.
            // It returns a DialogResult which will be used to display the results.
            var dialog = new DialogResult();
            bool didTaskSucceed = false;
            const string commandLineA = "cmd /c echo n | gpupdate /target:user /force";
            const string commandLineB = "cmd /c echo n | gpupdate /target:computer /force";

            // Setup WMI query.
            var options = new ConnectionOptions();
            if (GlobalVar.UseAlternateCredentials)
            {
                options.Username = GlobalVar.AlternateUsername;
                options.Password = GlobalVar.AlternatePassword;
                options.Authority = $"NTLMDOMAIN:{GlobalVar.AlternateDomain}";
            }
            var scope = new ManagementScope($@"\\{targetComputer}\root\CIMV2", options);

            try
            {
                scope.Connect();
                var objectGetOptions = new ObjectGetOptions();
                var managementPath = new ManagementPath("Win32_Process");
                var managementClass = new ManagementClass(scope, managementPath, objectGetOptions);

                ManagementBaseObject inParams = managementClass.GetMethodParameters("Create");
                inParams["CommandLine"] = commandLineA;
                ManagementBaseObject outParams = managementClass.InvokeMethod("Create", inParams, null);

                int returnValue;
                bool isReturnValid = int.TryParse(outParams["ReturnValue"].ToString(), out returnValue);
                if (!isReturnValid || returnValue != 0)
                    throw new Exception();

                Thread.Sleep(5000);

                inParams["CommandLine"] = commandLineB;
                outParams = managementClass.InvokeMethod("Create", inParams, null);

                isReturnValid = int.TryParse(outParams["ReturnValue"].ToString(), out returnValue);
                if (!isReturnValid || returnValue != 0)
                    throw new Exception();

                Thread.Sleep(7000);

                didTaskSucceed = true;
            }
            catch
            { }

            if (didTaskSucceed)
            {
                // GPUpdate was successful.  Build DialogResult to reflect success.
                dialog.DialogTitle = "Success";
                dialog.DialogBody = $"{targetComputer} has successfuly refreshed all Group Policy Objects.";
                dialog.DialogIconPath = "/Resources/success-48.png";
                dialog.ButtonIconPath = "/Resources/checkmark-24.png";
                dialog.ButtonText = "OK";
                dialog.IsCancelVisible = false;
            }
            else
            {
                // Failed to run GPUpdate.  Build DialogResult to reflect failure.
                dialog.DialogTitle = "Error";
                dialog.DialogBody = $"Failed to refresh Group Policy Objects on {targetComputer}.";
                dialog.DialogIconPath = "/Resources/error-48.png";
                dialog.ButtonIconPath = "/Resources/checkmark-24.png";
                dialog.ButtonText = "OK";
                dialog.IsCancelVisible = false;
            }

            return dialog;
        }
    }
}
