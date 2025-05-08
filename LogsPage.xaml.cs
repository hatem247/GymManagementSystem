using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace GymManagementSystem
{
    /// <summary>
    /// Interaction logic for LogsPage.xaml
    /// </summary>
    public partial class LogsPage : Page
    {
        public ObservableCollection<LogEntry> Logs { get; set; } = new ObservableCollection<LogEntry>();

        public LogsPage()
        {
            InitializeComponent();
            LogsDataGrid.ItemsSource = Logs;
            LogsFilterComboBox.SelectedIndex = 0;
            LoadLogs(LogsFilterComboBox.Text);
        }

        private void LoadLogs(string filter)
        {
            var logsFromExcel = ExcelHelper.GetLogs(filter);
            Logs.Clear();
            foreach (var log in logsFromExcel)
            {
                Logs.Add(log);
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            LoadLogs(LogsFilterComboBox.Text);
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService.CanGoBack)
            {
                NavigationService.GoBack();
            }
            else
            {
                NavigationService.Navigate(new HomePage());
            }
        }
    }
}