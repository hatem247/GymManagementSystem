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
            LoadLogs();
        }

        private void LoadLogs(DateTime? selectedDate = null)
        {
            var logsFromExcel = ExcelHelper.GetLogs(selectedDate);
            Logs.Clear();
            foreach (var log in logsFromExcel)
            {
                Logs.Add(log);
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            LoadLogs(LogsDatePicker.SelectedDate);
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