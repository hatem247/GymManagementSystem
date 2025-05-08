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
            IncomeFilterComboBox.SelectedIndex = 0;
            LoadIncome(IncomeFilterComboBox.Text);
        }

        private void LoadIncome(string filter)
        {
            var IncomesFromExcel = ExcelHelper.GetIncome(filter);
            int total = 0;
            Incomes.Clear();
            foreach (var income in IncomesFromExcel)
            {
                Incomes.Add(income);
                if(income.Date == DateTime.Today.ToShortDateString()) total += int.Parse(income.Amount);
            }
            totaltxt.Text = $"Today's Income: {total}";
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            LoadIncome(IncomeFilterComboBox.Text);
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
