using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.ServiceProcess;

namespace Splice_Admin.Classes
{
    class RemoteApplication
    {
        public static string ComputerName;
        public static TaskResult Result { get; set; }

        public string DisplayName { get; set; }
        public string Publisher { get; set; }
        public string Version { get; set; }
        public string InstallDate { get; set; }
        public string UninstallPath { get; set; }


        public static List<RemoteApplication> GetInstalledApplications()
        {
            var apps = new List<RemoteApplication>();
            var taskResult = new TaskResult();
            Result = taskResult;

            const string uninstallKey64 = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            const string uninstallKey32 = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
            const string serviceName = "RemoteRegistry";
            bool isLocal = ComputerName.ToUpper() == Environment.MachineName.ToUpper() ? true : false;
            bool isServiceRunning = true;

            // If the target computer is remote, then start the Remote Registry service.
            using (
                GlobalVar.UseAlternateCredentials
                ? UserImpersonation.Impersonate(GlobalVar.AlternateUsername, GlobalVar.AlternateDomain, GlobalVar.AlternatePassword)
                : null)
            using (var sc = new ServiceController(serviceName, ComputerName))
            {
                try
                {
                    if (!isLocal && sc.Status != ServiceControllerStatus.Running)
                    {
                        isServiceRunning = false;
                        sc.Start();
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    using (RegistryKey key = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, ComputerName))
                    {
                        using (RegistryKey mainKey64 = key.OpenSubKey(uninstallKey64))
                            apps.AddRange(EnumerateUninstallKeys(mainKey64));
                        using (RegistryKey mainKey32 = key.OpenSubKey(uninstallKey32))
                            apps.AddRange(EnumerateUninstallKeys(mainKey32));
                    }

                    var internetExplorerVersion = FileVersionInfo.GetVersionInfo($@"\\{ComputerName}\C$\Program Files\Internet Explorer\iexplore.exe");
                    if (internetExplorerVersion != null && internetExplorerVersion.ProductVersion.Length > 0)
                    {
                        apps.Add(new RemoteApplication
                        {
                            DisplayName = "Internet Explorer",
                            Publisher = "Microsoft Corporation",
                            Version = internetExplorerVersion.ProductVersion
                        });
                    }

                    taskResult.DidTaskSucceed = true;
                }
                catch
                {
                    taskResult.DidTaskSucceed = false;
                }


                // Cleanup.
                if (!isLocal && !isServiceRunning)
                {
                    try
                    {
                        if (sc != null)
                        {
                            sc.Stop();
                        }
                    }

                    catch (Exception)
                    {
                    }
                }
            }

            return apps;
        }


        private static List<RemoteApplication> EnumerateUninstallKeys(RegistryKey key)
        {
            var apps = new List<RemoteApplication>();
            if (key == null)
                return apps;

            foreach (string subKeyName in key.GetSubKeyNames())
            {
                using (RegistryKey subKey = key.OpenSubKey(subKeyName))
                {
                    var app = new RemoteApplication();
                    if (subKey.GetValue("SystemComponent") != null && subKey.GetValueKind("SystemComponent") == RegistryValueKind.DWord)
                    {
                        Int64 val = Convert.ToInt64(subKey.GetValue("SystemComponent").ToString());
                        if (val == 1)
                            continue;
                    }
                    if (subKey.GetValue("ParentKeyName") != null && subKey.GetValue("ParentKeyName").ToString().Length > 0)
                        continue;
                    if (subKey.GetValue("ReleaseType") != null && (subKey.GetValue("ReleaseType").ToString().Contains("Update") || subKey.GetValue("ReleaseType").ToString() == "Hotfix"))
                        continue;

                    if (subKey.GetValue("DisplayName") != null)
                    {
                        app.DisplayName = subKey.GetValue("DisplayName").ToString();
                    }
                    else
                        continue;

                    if (subKey.GetValue("Publisher") != null)
                        app.Publisher = subKey.GetValue("Publisher").ToString();
                    if (subKey.GetValue("DisplayVersion") != null)
                        app.Version = subKey.GetValue("DisplayVersion").ToString();
                    if (subKey.GetValue("UninstallString") != null)
                        app.UninstallPath = subKey.GetValue("UninstallString").ToString();

                    apps.Add(app);
                }
            }
            return apps;
        }
    }
}
