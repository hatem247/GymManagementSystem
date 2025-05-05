using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using OfficeOpenXml;

namespace GymManagementSystem
{
    public partial class ScanBarcodePage : Page
    {
        public ScanBarcodePage()
        {
            InitializeComponent();
        }

        private void SearchClient_Click(object sender, RoutedEventArgs e)
        {
            Client client = ExcelHelper.search(txtBarcodeInput.Text);
            if(client != null)
            {
                NavigationService.Navigate(new ClientDetailsPage(client));
                ExcelHelper.AddLogEntry(client.FullName, client.PhoneNumber);
            }
        }
        private void Back_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new HomePage());
        }

    }
}
