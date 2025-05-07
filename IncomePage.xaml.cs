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
    /// Interaction logic for IncomePage.xaml
    /// </summary>
    public partial class IncomePage : Page
    {
        public IncomePage()
        {
            InitializeComponent();
            LoadIncome("");
        }

        private void LoadIncome(string filter)
        {
            IncomeDataGrid.ItemsSource = null;
            var Income = ExcelHelper.GetIncome(filter);
            IncomeDataGrid.ItemsSource = Income;
        }

        private void Back_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new HomePage());
        }

        private void IncomeFilterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(IncomeFilterComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                LoadIncome(IncomeFilterComboBox.Text);
            }
        }
    }
}
