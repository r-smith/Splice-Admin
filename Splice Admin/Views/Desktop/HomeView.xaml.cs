using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Splice_Admin.Views.Desktop
{
    /// <summary>
    /// Interaction logic for HomeView.xaml
    /// </summary>
    public partial class HomeView : UserControl
    {
        public HomeView()
        {
            InitializeComponent();
            
            Version version = typeof(MainWindow).Assembly.GetName().Version;
            tbVersion.Text = $"Build: {version.Major}.{version.Minor}";
        }

        private void tbChangeLog_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var wnd = new ChangeLogWindow();
            wnd.Show();
        }
    }
}
