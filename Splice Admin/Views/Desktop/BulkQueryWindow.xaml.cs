using Microsoft.Win32;
using RemoteDesktopServicesAPI;
using Splice_Admin.Classes;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Management;
using System.ServiceProcess;
using System.Windows;

namespace Splice_Admin.Views.Desktop
{
    /// <summary>
    /// Interaction logic for BulkQueryWindow.xaml
    /// </summary>
    public partial class BulkQueryWindow : Window
    {
        private ObservableCollection<QueryResult> HasMatchCollection = new ObservableCollection<QueryResult>();
        private ObservableCollection<QueryResult> NoMatchCollection = new ObservableCollection<QueryResult>();
        private BackgroundWorker bw;
        private int SearchItemsRemaining;

        public BulkQueryWindow(RemoteBulkQuery bulkQuery)
        {
            InitializeComponent();

            dgHasMatch.ItemsSource = HasMatchCollection;
            dgNoMatch.ItemsSource = NoMatchCollection;

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
                case RemoteBulkQuery.QueryType.Process:
                    tbSearchType.Text = "Running Process";
                    break;
                case RemoteBulkQuery.QueryType.Service:
                    tbSearchType.Text = "Windows Service";
                    break;
                default:
                    tbSearchType.Text = "Unknown";
                    break;
            }
            tbSearchPhrase.Text = bulkQuery.SearchPhrase;
            SearchItemsRemaining = bulkQuery.TargetComputerList.Count;
            txtRemainingCount.Text = $"Remaining: {SearchItemsRemaining}";

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
                case RemoteBulkQuery.QueryType.Process:
                    foreach (var targetComputer in bulkQuery.TargetComputerList)
                    {
                        bw.ReportProgress(
                            (int)QueryResult.Type.ProgressReport,
                            new QueryResult { ComputerName = targetComputer });
                        SearchForProcess(targetComputer, bulkQuery.SearchPhrase);
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
            txtRemainingCount.Text = "";
        }

        private void bgThread_RunQueryProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            var queryResult = e.UserState as QueryResult;
            switch (e.ProgressPercentage)
            {
                case ((int)QueryResult.Type.Match):
                    HasMatchCollection.Add(queryResult);
                    break;
                case ((int)QueryResult.Type.NoMatch):
                    NoMatchCollection.Add(queryResult);
                    break;
                case ((int)QueryResult.Type.ProgressReport):
                    txtStatus.Text = $"Querying: {queryResult.ComputerName}";
                    txtRemainingCount.Text = $"Remaining: {--SearchItemsRemaining}";
                    break;
            }
        }

        private void SearchForFile(string targetComputer, string searchPhrase)
        {
            searchPhrase = searchPhrase.TrimEnd('\\');

            // Setup WMI query.
            var options = new ConnectionOptions();
            var scope = new ManagementScope($@"\\{targetComputer}\root\CIMV2", options);
            var query = new ObjectQuery($@"SELECT * FROM CIM_DataFile WHERE Name = '{searchPhrase.Replace(@"\", @"\\")}'");
            var searcher = new ManagementObjectSearcher(scope, query);

            try
            {
                if (searcher.Get().Count > 0)
                    bw.ReportProgress(
                        (int)QueryResult.Type.Match,
                        new QueryResult { ComputerName = targetComputer, ResultText = "File found." });
                else
                {
                    query = new ObjectQuery($@"SELECT * FROM Win32_Directory WHERE Name = '{searchPhrase.Replace(@"\", @"\\")}'");
                    searcher = new ManagementObjectSearcher(scope, query);

                    if (searcher.Get().Count > 0)
                        bw.ReportProgress(
                            (int)QueryResult.Type.Match,
                            new QueryResult { ComputerName = targetComputer, ResultText = "Directory found." });
                    else
                        bw.ReportProgress(
                            (int)QueryResult.Type.NoMatch,
                            new QueryResult { ComputerName = targetComputer, ResultText = "File not found." });
                }
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message;
                if (ex.Message.Contains("("))
                    errorMessage = errorMessage.Substring(0, errorMessage.IndexOf('('));

                bw.ReportProgress(
                    (int)QueryResult.Type.NoMatch,
                    new QueryResult { ComputerName = targetComputer, ResultText = errorMessage.Trim() });
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
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var displayName = (obj["DisplayName"] != null) ? obj["DisplayName"].ToString() : string.Empty;
                        var startMode = (obj["StartMode"] != null) ? obj["StartMode"].ToString() : string.Empty;
                        var state = (obj["State"] != null) ? obj["State"].ToString() : string.Empty;

                        bw.ReportProgress(
                            (int)QueryResult.Type.Match,
                            new QueryResult { ComputerName = targetComputer, ResultText = $"{displayName} ({startMode}) - {state}" });
                    }
                }
                else
                    bw.ReportProgress(
                        (int)QueryResult.Type.NoMatch,
                        new QueryResult { ComputerName = targetComputer, ResultText = "Service not found." });
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message;
                if (ex.Message.Contains("("))
                    errorMessage = errorMessage.Substring(0, errorMessage.IndexOf('('));

                bw.ReportProgress(
                    (int)QueryResult.Type.NoMatch,
                    new QueryResult { ComputerName = targetComputer, ResultText = errorMessage.Trim() });
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
                    int numberOfMatchingApplications = 0;

                    using (RegistryKey key = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, targetComputer, RegistryView.Registry64))
                    using (RegistryKey mainKey64 = key.OpenSubKey(uninstallKey64))
                        numberOfMatchingApplications += EnumerateUninstallKeys(mainKey64, searchPhrase, targetComputer);

                    using (RegistryKey key = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, targetComputer, RegistryView.Registry32))
                    using (RegistryKey mainKey32 = key.OpenSubKey(uninstallKey64))
                        numberOfMatchingApplications += EnumerateUninstallKeys(mainKey32, searchPhrase, targetComputer);

                    if (numberOfMatchingApplications == 0)
                        bw.ReportProgress(
                            (int)QueryResult.Type.NoMatch,
                            new QueryResult { ComputerName = targetComputer, ResultText = "Application not found." });
                }
                catch (Exception ex)
                {
                    string errorMessage = ex.Message;
                    if (ex.Message.Contains("("))
                        errorMessage = errorMessage.Substring(0, errorMessage.IndexOf('('));

                    bw.ReportProgress(
                        (int)QueryResult.Type.NoMatch,
                        new QueryResult { ComputerName = targetComputer, ResultText = errorMessage.Trim() });
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
                            (int)QueryResult.Type.Match,
                            new QueryResult { ComputerName = targetComputer, ResultText = "User logged in." });
                    else
                        bw.ReportProgress(
                            (int)QueryResult.Type.NoMatch,
                            new QueryResult { ComputerName = targetComputer, ResultText = "User not logged in." });
                }
                catch (Exception ex)
                {
                    string errorMessage = ex.Message;
                    if (ex.Message.Contains("("))
                        errorMessage = errorMessage.Substring(0, errorMessage.IndexOf('('));

                    bw.ReportProgress(
                        (int)QueryResult.Type.NoMatch,
                        new QueryResult { ComputerName = targetComputer, ResultText = errorMessage.Trim() });
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
                            (int)QueryResult.Type.Match,
                            new QueryResult { ComputerName = targetComputer, ResultText = "User logged in." });
                    else
                        bw.ReportProgress(
                            (int)QueryResult.Type.NoMatch,
                            new QueryResult { ComputerName = targetComputer, ResultText = "User not logged in." });
                }
                catch (Exception ex)
                {
                    string errorMessage = ex.Message;
                    if (ex.Message.Contains("("))
                        errorMessage = errorMessage.Substring(0, errorMessage.IndexOf('('));

                    bw.ReportProgress(
                        (int)QueryResult.Type.NoMatch,
                        new QueryResult { ComputerName = targetComputer, ResultText = errorMessage.Trim() });
                }
            }
        }


        private void SearchForProcess(string targetComputer, string searchPhrase)
        {
            // Setup WMI Query.
            var options = new ConnectionOptions();
            var scope = new ManagementScope($@"\\{targetComputer}\root\CIMV2", options);
            var query = new ObjectQuery($"SELECT * FROM Win32_Process WHERE Name LIKE '%{searchPhrase}%'");
            var searcher = new ManagementObjectSearcher(scope, query);
            
            try
            {
                if (searcher.Get().Count > 0)
                {
                    // Retrieve a list of running processes.
                    foreach (ManagementObject m in searcher.Get())
                    {
                        if (m["Name"] != null)
                        {
                            var resultText = m["Name"].ToString();
                            if (m["ProcessId"] != null) resultText += $" ({m["ProcessId"]})";

                            var argList = new string[] { string.Empty, string.Empty };
                            int returnVal = Convert.ToInt32(m.InvokeMethod("GetOwner", argList));

                            string processOwner = (returnVal == 0) ? argList[0] : string.Empty;

                            switch (processOwner.ToUpper())
                            {
                                case ("SYSTEM"):
                                    processOwner = "System";
                                    break;
                                case ("LOCAL SERVICE"):
                                    processOwner = "Local Service";
                                    break;
                                case ("NETWORK SERVICE"):
                                    processOwner = "Network Service";
                                    break;
                            }

                            if (processOwner.Length > 0) resultText += $" ({processOwner})";

                            bw.ReportProgress(
                                (int)QueryResult.Type.Match,
                                new QueryResult { ComputerName = targetComputer, ResultText = resultText });
                        }
                    }
                }
                else
                    bw.ReportProgress(
                        (int)QueryResult.Type.NoMatch,
                        new QueryResult { ComputerName = targetComputer, ResultText = "Process not found." });
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message;
                if (ex.Message.Contains("("))
                    errorMessage = errorMessage.Substring(0, errorMessage.IndexOf('('));

                bw.ReportProgress(
                    (int)QueryResult.Type.NoMatch,
                    new QueryResult { ComputerName = targetComputer, ResultText = errorMessage.Trim() });
            }
        }


        private int EnumerateUninstallKeys(RegistryKey key, string searchPhrase, string targetComputer)
        {
            int numberOfMatchingApplications = 0;

            if (key == null)
                return numberOfMatchingApplications;

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
                            ++numberOfMatchingApplications;

                            string resultText;
                            if (subKey.GetValue("DisplayVersion") != null)
                                resultText = $"{subKey.GetValue("DisplayName")} [{subKey.GetValue("DisplayVersion")}]";
                            else
                                resultText = subKey.GetValue("DisplayName").ToString();

                            bw.ReportProgress(
                                (int)QueryResult.Type.Match,
                                new QueryResult { ComputerName = targetComputer, ResultText = resultText });
                        }
                    }
                    else
                        continue;
                }
            }
            return numberOfMatchingApplications;
        }

    }
}
