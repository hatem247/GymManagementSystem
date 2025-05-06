using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace GymManagementSystem
{
    /// <summary>
    /// Interaction logic for LogsPage.xaml
    /// </summary>
    public partial class LogsPage : Page
    {
        public LogsPage()
        {
            InitializeComponent();
            LoadLogs("");
        }

        private void LoadLogs(string filter)
        {
            var Logs = ExcelHelper.GetLogs(filter);
            LogsDataGrid.ItemsSource = null;
            LogsDataGrid.ItemsSource = Logs;
        }

        private void Back_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new HomePage());
        }

        private void LogsFilterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (LogsFilterComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                string filter = LogsFilterComboBox.SelectedIndex == -1 ? "" : selectedItem.Content.ToString();
                LoadLogs(filter);
            }
        }
    }
}
