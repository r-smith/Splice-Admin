using Splice_Admin.Classes;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Splice_Admin.Views.Desktop
{
    /// <summary>
    /// Interaction logic for BulkQueryView.xaml
    /// </summary>
    public partial class BulkQueryView : UserControl
    {
        public BulkQueryView()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var bulkQuery = new RemoteBulkQuery();
            bulkQuery.TargetComputerList = new List<string>();
            if (rdoManual.IsChecked == true)
            {
                bulkQuery.TargetComputerList = txtTargetComputer.Text.Split(',').Select(t => t.Trim()).ToList();
            }
            else if (rdoAllDomainComputers.IsChecked == true)
            {
                bulkQuery.TargetComputerList = GetDomainComputers(DomainComputer.All);
            }
            else if (rdoAllServers.IsChecked == true)
            {
                bulkQuery.TargetComputerList = GetDomainComputers(DomainComputer.Server);
            }
            else if (rdoAllWorkstations.IsChecked == true)
            {
                bulkQuery.TargetComputerList = GetDomainComputers(DomainComputer.Workstation);
            }
            
            bulkQuery.SearchPhrase = txtSearchPhrase.Text;

            switch (cboQueryType.SelectedIndex)
            {
                case 0:
                    bulkQuery.SearchType = RemoteBulkQuery.QueryType.File;
                    break;
                case 1:
                    bulkQuery.SearchType = RemoteBulkQuery.QueryType.InstalledApplication;
                    break;
                case 2:
                    bulkQuery.SearchType = RemoteBulkQuery.QueryType.LoggedOnUser;
                    break;
                case 3:
                    bulkQuery.SearchType = RemoteBulkQuery.QueryType.Service;
                    break;
                default:
                    bulkQuery.SearchType = RemoteBulkQuery.QueryType.File;
                    break;
            }

            var wnd = new BulkQueryWindow(bulkQuery);
            wnd.Show();
        }

        private List<string> GetDomainComputers(DomainComputer searchType)
        {
            List<string> computerNames = new List<string>();

            using (DirectoryEntry directoryEntry = new DirectoryEntry($"LDAP://{Environment.UserDomainName}"))
            using (DirectorySearcher directorySearcher = new DirectorySearcher(directoryEntry))
            {
                switch (searchType)
                {
                    case DomainComputer.All:
                        directorySearcher.Filter = ("(objectClass=computer)");
                        break;
                    case DomainComputer.Server:
                        directorySearcher.Filter = ("(&(objectClass=computer)(| (operatingSystem=Windows Server *)(operatingSystem=Windows 2000 Server) ))");
                        break;
                    case DomainComputer.Workstation:
                        directorySearcher.Filter = ("(&(objectClass=computer)(!operatingSystem=Windows Server *)(!operatingSystem=Windows 2000 Server))");
                        break;
                    default:
                        directorySearcher.Filter = ("(objectClass=computer)");
                        break;
                }
                directorySearcher.SizeLimit = int.MaxValue;
                directorySearcher.PageSize = int.MaxValue;

                foreach (SearchResult searchResult in directorySearcher.FindAll())
                {
                    string computerName = searchResult.GetDirectoryEntry().Name;
                    if (computerName.StartsWith("CN="))
                        computerName = computerName.Substring(3);
                    computerNames.Add(computerName);
                }
            }

            computerNames.Sort();
            return computerNames;
        }

        private enum DomainComputer
        {
            All,
            Server,
            Workstation
        }
    }
}
