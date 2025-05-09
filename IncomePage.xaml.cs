using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    /// Interaction logic for IncomePage.xaml
    /// </summary>
    public partial class IncomePage : Page
    {
        public ObservableCollection<IncomeEntry> Incomes { get; set; } = new ObservableCollection<IncomeEntry>();

        public IncomePage()
        {
            InitializeComponent();
            IncomeDataGrid.ItemsSource = Incomes;
            LoadIncome();
            LoadMonthIncome();
        }

        private void LoadIncome(DateTime? dateTime = null)
        {
            var IncomesFromExcel = ExcelHelper.GetIncome(dateTime);
            int total = 0;
            Incomes.Clear();
            foreach (var income in IncomesFromExcel)
            {
                Incomes.Add(income);
                if(income.Date == DateTime.Today.ToShortDateString()) total += int.Parse(income.Amount);
            }
            totaltodaytxt.Text = $"Today's Income: {total}";
        }

        private void IncomeDatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (IncomeDatePicker.SelectedDate != null)
            {
                LoadIncome(IncomeDatePicker.SelectedDate);
                LoadMonthIncome(IncomeDatePicker.SelectedDate);
            }
        }

        private void LoadMonthIncome(DateTime? dateTime = null)
        {
            var incomesFromExcel = ExcelHelper.GetIncome();

            DateTime targetDate = dateTime ?? DateTime.Today;
            DateTime startOfMonth = new DateTime(targetDate.Year, targetDate.Month, 1);

            int total = 0;

            foreach (var income in incomesFromExcel)
            {
                if (DateTime.TryParse(income.Date, out DateTime parsedDate))
                {
                    if (parsedDate >= startOfMonth && parsedDate <= targetDate)
                    {
                        total += int.Parse(income.Amount);
                    }
                }
            }

            totalmonthtxt.Text = dateTime == null
                ? $"This Month's Income: {total}"
                : $"Income from {startOfMonth:dd/MM/yyyy} to {targetDate:dd/MM/yyyy}: {total}";
        }




        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            LoadIncome(IncomeDatePicker.SelectedDate);
            LoadMonthIncome(IncomeDatePicker.SelectedDate);
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
