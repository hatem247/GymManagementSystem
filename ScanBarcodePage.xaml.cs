using System;
using System.Linq;
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
            bool logt = false;
            if (client != null)
            {
                var logs = ExcelHelper.GetLogs();
                var filtered = logs.Where(l => l.Phone == client.PhoneNumber).ToList();
                foreach(var log in filtered)
                {
                    if(log.Date == DateTime.Today.ToShortDateString())
                    {
                        logt = true;
                        break;
                    }
                }
                if (!logt)
                {
                    ExcelHelper.AddLogEntry(client.FullName, client.PhoneNumber);
                    NavigationService.Navigate(new ClientDetailsPage(client));
                }
                else MessageBox.Show("Client Already attended today");
                txtBarcodeInput.Clear();
            }
            else
            {
                MessageBox.Show("Client not found.");
                txtBarcodeInput.SelectAll();
            }
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
