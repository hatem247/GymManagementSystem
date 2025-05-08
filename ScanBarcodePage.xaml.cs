using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace GymManagementSystem
{
    public partial class ScanBarcodePage : Page
    {
        public ScanBarcodePage()
        {
            InitializeComponent();
            txtBarcodeInput.Focus();
            txtBarcodeInput.KeyDown += TxtBarcodeInput_KeyDown;
            
        }

        private void TxtBarcodeInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ProcessBarcode();
            }
        }

        private void SearchClient_Click(object sender, RoutedEventArgs e)
        {
            ProcessBarcode();
        }

        private void ProcessBarcode()
        {
            string barcode = txtBarcodeInput.Text.Trim();

            if (string.IsNullOrEmpty(barcode))
            {
                MessageBox.Show("Please scan a barcode or enter a phone number.");
                return;
            }

            Client client = ExcelHelper.Search(barcode);
            if (client != null)
            {
                ExcelHelper.AddLogEntry(client.FullName, client.PhoneNumber);
                NavigationService.Navigate(new ClientDetailsPage(client));
                txtBarcodeInput.Clear();
            }
            else
            {
                MessageBox.Show("Client not found.");
                txtBarcodeInput.SelectAll();
            }
        }

        private void Back_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new HomePage());
        }
    }
}
