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
            var incomesFromExcel = ExcelHelper.GetIncome(dateTime);
            int total = 0;

            DateTime startDate;
            DateTime endDate;

            if (dateTime == null)
            {
                startDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                endDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month,
                                       DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month));
            }
            else
            {
                startDate = dateTime.Value;
                endDate = new DateTime(startDate.Year, startDate.Month,
                                       DateTime.DaysInMonth(startDate.Year, startDate.Month));
            }

            foreach (var income in incomesFromExcel)
            {
                if (DateTime.TryParse(income.Date, out DateTime parsedDate))
                {
                    if (parsedDate >= startDate && parsedDate <= endDate)
                    {
                        if (int.TryParse(income.Amount, out int amount))
                            total += amount;
                    }
                }
            }

            totalmonthtxt.Text = dateTime == null
                ? $"This Month's Income: {total}"
                : $"Income from {startDate:dd/MM/yyyy} to {endDate:dd/MM/yyyy}: {total}";
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
