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
            string path = Path.Combine(Environment.CurrentDirectory, "Clients.xlsx");

            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var sheet = package.Workbook.Worksheets[0];
                for (int i = 2; i <= sheet.Dimension.End.Row; i++)
                {
                    string phone = sheet.Cells[i, 4].Text;
                    if (phone == txtBarcodeInput.Text)
                    {
                        Client client = new Client
                        {
                            FullName = sheet.Cells[i, 1].Text,
                            Age = int.Parse(sheet.Cells[i, 2].Text),
                            PhoneNumber = sheet.Cells[i, 4].Text,
                            Weight = double.Parse(sheet.Cells[i, 5].Text),
                            Height = double.Parse(sheet.Cells[i, 6].Text),
                            SubscriptionType = sheet.Cells[i, 7].Text,
                            SubscriptionStart = DateTime.Parse(sheet.Cells[i, 8].Text),
                            SubscriptionEnd = DateTime.Parse(sheet.Cells[i, 9].Text),
                            IsFrozen = sheet.Cells[i, 10].Text.ToLower() == "true"
                        };

                        NavigationService.Navigate(new ClientDetailsPage(client));
                        ExcelHelper.AddLogEntry(Name, phone);
                        return;
                    }
                }
                MessageBox.Show("Client not found.");
            }
        }
        private void Back_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new HomePage());
        }

    }
}
