using Microsoft.Win32;
using RemoteDesktopServicesAPI;
using Splice_Admin.Classes;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Splice_Admin.Views.Desktop
{
    /// <summary>
    /// Interaction logic for BulkQueryWindow.xaml
    /// </summary>
    public partial class BulkQueryWindow : Window
    {
        private ObservableCollection<string> HasMatchCollection = new ObservableCollection<string>();
        private ObservableCollection<string> NoMatchCollection = new ObservableCollection<string>();
        private ObservableCollection<string> ErrorCollection = new ObservableCollection<string>();
        private BackgroundWorker bw;

        public BulkQueryWindow(RemoteBulkQuery bulkQuery)
        {
            InitializeComponent();

            lbHasMatch.ItemsSource = HasMatchCollection;
            lbNoMatch.ItemsSource = NoMatchCollection;
            lbError.ItemsSource = ErrorCollection;

            switch (bulkQuery.SearchType)
            {
                case RemoteBulkQuery.QueryType.File:
                    tbSearchType.Text = "File or Directory";
                    break;
                case RemoteBulkQuery.QueryType.InstalledApplication:
                    tbSearchType.Text = "Installed Application";
                    break;
                case RemoteBulkQuery.QueryType.LoggedOnUser:
                    tbSearchType.Text = "Logged On User";
                    break;
                case RemoteBulkQuery.QueryType.Service:
                    tbSearchType.Text = "Windows Service";
                    break;
                default:
                    tbSearchType.Text = "Unknown";
                    break;
            }
            tbSearchPhrase.Text = bulkQuery.SearchPhrase;

            // Setup a background thread.
            bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.DoWork += new DoWorkEventHandler(bgThread_RunQuery);
            bw.ProgressChanged += new ProgressChangedEventHandler(bgThread_RunQueryProgressChanged);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgThread_RunQueryCompleted);

            // Launch background thread to retrieve a list of running services.
            bw.RunWorkerAsync(bulkQuery);
        }

        private void bgThread_RunQuery(object sender, DoWorkEventArgs e)
        {
            var bulkQuery = e.Argument as RemoteBulkQuery;

            switch (bulkQuery.SearchType)
            {
                case RemoteBulkQuery.QueryType.File:
                    foreach (var targetComputer in bulkQuery.TargetComputerList)
                    {
                        bw.ReportProgress(
                            (int)QueryResult.Type.ProgressReport,
                            new QueryResult { ComputerName = targetComputer });
                        SearchForFile(targetComputer, bulkQuery.SearchPhrase);
                    }
                    break;
                case RemoteBulkQuery.QueryType.InstalledApplication:
                    foreach (var targetComputer in bulkQuery.TargetComputerList)
                    {
                        bw.ReportProgress(
                            (int)QueryResult.Type.ProgressReport,
                            new QueryResult { ComputerName = targetComputer });
                        SearchForInstalledApplication(targetComputer, bulkQuery.SearchPhrase);
                    }
                    break;
                case RemoteBulkQuery.QueryType.LoggedOnUser:
                    foreach (var targetComputer in bulkQuery.TargetComputerList)
                    {
                        bw.ReportProgress(
                            (int)QueryResult.Type.ProgressReport,
                            new QueryResult { ComputerName = targetComputer });
                        SearchForLoggedOnUser(targetComputer, bulkQuery.SearchPhrase);
                    }
                    break;
                case RemoteBulkQuery.QueryType.Service:
                    foreach (var targetComputer in bulkQuery.TargetComputerList)
                    {
                        bw.ReportProgress(
                            (int)QueryResult.Type.ProgressReport,
                            new QueryResult { ComputerName = targetComputer });
                        SearchForWindowsService(targetComputer, bulkQuery.SearchPhrase);
                    }
                    break;
            }
        }

        private void bgThread_RunQueryCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            txtStatus.Text = "Search complete";
        }

        private void bgThread_RunQueryProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            var queryResult = e.UserState as QueryResult;
            switch (e.ProgressPercentage)
            {
                case ((int)QueryResult.Type.HasMatch):
                    HasMatchCollection.Add(queryResult.ComputerName);
                    break;
                case ((int)QueryResult.Type.NoMatch):
                    NoMatchCollection.Add(queryResult.ComputerName);
                    break;
                case ((int)QueryResult.Type.Error):
                    ErrorCollection.Add(queryResult.ComputerName);
                    break;
                case ((int)QueryResult.Type.ProgressReport):
                    txtStatus.Text = "Querying: " + queryResult.ComputerName;
                    break;
            }
        }

        private void SearchForFile(string targetComputer, string searchPhrase)
        {
            try
            {
                var pathRoot = Path.GetPathRoot(searchPhrase);
                var pathFolder = searchPhrase.Substring(pathRoot.Length);
                var uncPath = $@"\\{targetComputer}\{pathRoot.Substring(0, 1)}$\{pathFolder}";

                if (File.Exists(uncPath) || Directory.Exists(uncPath))
                    bw.ReportProgress(
                        (int)QueryResult.Type.HasMatch,
                        new QueryResult { ComputerName = targetComputer });
                else
                    bw.ReportProgress(
                        (int)QueryResult.Type.NoMatch,
                        new QueryResult { ComputerName = targetComputer });
            }
            catch (Exception ex)
            {
                bw.ReportProgress(
                    (int)QueryResult.Type.Error,
                    new QueryResult { ComputerName = targetComputer, ResultText = ex.Message });
            }
        }

        private void SearchForWindowsService(string targetComputer, string searchPhrase)
        {
            // Setup WMI query.
            var options = new ConnectionOptions();
            var scope = new ManagementScope($@"\\{targetComputer}\root\CIMV2", options);
            var query = new ObjectQuery($"SELECT * FROM Win32_Service WHERE DisplayName LIKE '%{searchPhrase}%' OR Name LIKE '%{searchPhrase}%'");
            var searcher = new ManagementObjectSearcher(scope, query);
            
            try
            {
                if (searcher.Get().Count > 0)
                    bw.ReportProgress(
                        (int)QueryResult.Type.HasMatch,
                        new QueryResult { ComputerName = targetComputer });
                else
                    bw.ReportProgress(
                        (int)QueryResult.Type.NoMatch,
                        new QueryResult { ComputerName = targetComputer });
            }
            catch (Exception ex)
            {
                bw.ReportProgress(
                    (int)QueryResult.Type.Error,
                    new QueryResult { ComputerName = targetComputer, ResultText = ex.Message });
            }
        }

        private void SearchForInstalledApplication(string targetComputer, string searchPhrase)
        {
            const string uninstallKey64 = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            const string uninstallKey32 = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
            const string serviceName = "RemoteRegistry";
            bool isLocal = targetComputer.ToUpper() == Environment.MachineName.ToUpper() ? true : false;
            bool isServiceRunning = true;

            // If the target computer is remote, then start the Remote Registry service.
            using (var sc = new ServiceController(serviceName, targetComputer))
            {
                try
                {
                    if (!isLocal && sc.Status != ServiceControllerStatus.Running)
                    {
                        isServiceRunning = false;
                        sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    bool hasMatch = false;
                    using (RegistryKey key = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, targetComputer, RegistryView.Registry64))
                    using (RegistryKey mainKey64 = key.OpenSubKey(uninstallKey64))
                        hasMatch = EnumerateUninstallKeys(mainKey64, searchPhrase);

                    if (hasMatch == false)
                    {
                        using (RegistryKey key = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, targetComputer, RegistryView.Registry32))
                        using (RegistryKey mainKey32 = key.OpenSubKey(uninstallKey64))
                            hasMatch = EnumerateUninstallKeys(mainKey32, searchPhrase);
                    }

                    if (hasMatch)
                        bw.ReportProgress(
                            (int)QueryResult.Type.HasMatch,
                            new QueryResult { ComputerName = targetComputer });
                    else
                        bw.ReportProgress(
                            (int)QueryResult.Type.NoMatch,
                            new QueryResult { ComputerName = targetComputer });
                }
                catch (Exception ex)
                {
                    bw.ReportProgress(
                        (int)QueryResult.Type.Error,
                        new QueryResult { ComputerName = targetComputer, ResultText = ex.Message });
                }


                // Cleanup.
                if (!isLocal && !isServiceRunning)
                {
                    try
                    {
                        if (sc != null)
                            sc.Stop();
                    }

                    catch (Exception)
                    {
                    }
                }
            }
        }

        private void SearchForLoggedOnUser(string targetComputer, string searchPhrase)
        {
            UInt32 productType = 1;

            // Determine whether operating system is server or desktop edition.
            var options = new ConnectionOptions();
            var scope = new ManagementScope($@"\\{targetComputer}\root\CIMV2", options);
            var query = new ObjectQuery("SELECT ProductType FROM Win32_OperatingSystem");
            var searcher = new ManagementObjectSearcher(scope, query);

            try
            {
                foreach (ManagementObject m in searcher.Get())
                {
                    productType = (m["ProductType"] != null) ? (UInt32)m["ProductType"] : 1;
                    break;
                }
            }
            catch
            {
            }

            // If operating system is server edition, use Remote Desktop Services API to retrieve logon sessions.
            if (productType > 1)
            {
                try
                {
                    IntPtr server = WtsApi.WTSOpenServer(targetComputer);
                    if (WtsApi.IsUserLoggedOn(server, searchPhrase))
                        bw.ReportProgress(
                            (int)QueryResult.Type.HasMatch,
                            new QueryResult { ComputerName = targetComputer });
                    else
                        bw.ReportProgress(
                            (int)QueryResult.Type.NoMatch,
                            new QueryResult { ComputerName = targetComputer });
                }
                catch (Exception ex)
                {
                    bw.ReportProgress(
                        (int)QueryResult.Type.Error,
                        new QueryResult { ComputerName = targetComputer, ResultText = ex.Message });
                }
            }
            // If operating system is desktop edition, query Win32_Process for explorer.exe to determine logged on users.
            else
            {
                query = new ObjectQuery("SELECT * FROM Win32_Process WHERE Name = 'explorer.exe'");
                searcher = new ManagementObjectSearcher(scope, query);

                try
                {
                    bool hasMatch = false;
                    foreach (ManagementObject m in searcher.Get())
                    {
                        string loggedOnUser;

                        string[] argList = new string[] { string.Empty, string.Empty };
                        int returnVal = Convert.ToInt32(m.InvokeMethod("GetOwner", argList));
                        if (returnVal == 0)
                            loggedOnUser = argList[0];
                        else
                            loggedOnUser = string.Empty;

                        if (loggedOnUser.Equals(searchPhrase, StringComparison.OrdinalIgnoreCase))
                        {
                            hasMatch = true;
                            break;
                        }
                    }

                    if (hasMatch)
                        bw.ReportProgress(
                            (int)QueryResult.Type.HasMatch,
                            new QueryResult { ComputerName = targetComputer });
                    else
                        bw.ReportProgress(
                            (int)QueryResult.Type.NoMatch,
                            new QueryResult { ComputerName = targetComputer });
                }
                catch (Exception ex)
                {
                    bw.ReportProgress(
                        (int)QueryResult.Type.Error,
                        new QueryResult { ComputerName = targetComputer, ResultText = ex.Message });
                }
            }
        }


        private bool EnumerateUninstallKeys(RegistryKey key, string searchPhrase)
        {
            bool matchFound = false;

            if (key == null)
                return false;

            foreach (string subKeyName in key.GetSubKeyNames())
            {
                using (RegistryKey subKey = key.OpenSubKey(subKeyName))
                {
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
                        if (subKey.GetValue("DisplayName").ToString().ToUpper().Contains(searchPhrase.ToUpper()))
                        {
                            matchFound = true;
                            break;
                        }
                    }
                    else
                        continue;
                }
            }
            return matchFound;
        }

    }
}
