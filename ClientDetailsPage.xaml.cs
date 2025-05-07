using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace GymManagementSystem
{
    public partial class ClientDetailsPage : Page
    {
        private Client client;

        public ClientDetailsPage(Client clientData)
        {
            InitializeComponent();
            client = clientData;
            DisplayClientInfo();
            LoadClientLogs(client.PhoneNumber);
        }

        private void DisplayClientInfo()
        {
            client = ExcelHelper.Search(client.PhoneNumber);
            txtName.Text = client.FullName;
            txtPhone.Text = client.PhoneNumber;
            txtWeight.Text = client.Weight.ToString("F1");
            txtSubscription.Text = client.SubscriptionType;
            txtStart.Text = client.SubscriptionStart.ToShortDateString();
            txtEnd.Text = client.SubscriptionEnd.ToShortDateString();

            int daysLeft = (client.SubscriptionEnd - DateTime.Today).Days - 2;
            txtDaysLeft.Text = daysLeft >= 0 ? $"{daysLeft} days" : "Expired";

            btnFreeze.Content = client.IsFrozen ? "Continue" : "Freeze";
        }

        private void Back_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new HomePage());
        }

        private void HomeIcon_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new HomePage());
        }

        private void LoadClientLogs(string phoneNumber)
        {
            var logs = ExcelHelper.GetLogs("");
            var filtered = logs.Where(l => l.Phone == phoneNumber).ToList();
            LogsDataGrid.ItemsSource = filtered;
        }

        private void FreezeButton_Click(object sender, RoutedEventArgs e)
        {
            if (client.IsFrozen)
            {
                // Unfreeze the client
                ExcelHelper.UnfreezeClient(client.PhoneNumber);
                DisplayClientInfo();
                MessageBox.Show("Client has been unfrozen.");
                client.IsFrozen = false;
                btnFreeze.Content = "Freeze";
            }
            else
            {
                // Open freeze dialog
                if (client.SubscriptionType.Contains("15 Days") || client.SubscriptionType.Contains("1 Month"))
                {
                    MessageBox.Show("Can't freeze a client because subscription bundle is less than 3 months");
                }
                else
                {
                    var freezeDialog = new FreezeDialog();
                    if (freezeDialog.ShowDialog() == true)
                    {
                        int days = freezeDialog.SelectedDays;
                        ExcelHelper.FreezeClient(client.PhoneNumber, days);
                        DisplayClientInfo();
                        MessageBox.Show($"Client frozen for {days} day(s).");
                        client.IsFrozen = true;
                        btnFreeze.Content = "Continue";
                    }
                }
            }
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            var updateDialog = new UpdateClientDialog(client);
            if (updateDialog.ShowDialog() == true)
            {
                DisplayClientInfo();
            }
        }
        
        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            var deleteDialog = new DeleteDialoge(client.PhoneNumber);
            if (deleteDialog.ShowDialog() == true)
            {
                DisplayClientInfo();
            }
        }

        private void btnRenew_Click(object sender, RoutedEventArgs e)
        {
            var renewDialog = new RenewBundleDialog();
            if (renewDialog.ShowDialog() == true)
            {
                string selectedBundle = renewDialog.BundleBox.Text;
                string selectedDuration = renewDialog.SubscipriontypeBox.Text;

                ExcelHelper.RenewClientSubscription(client.PhoneNumber, selectedBundle, selectedDuration);
                DisplayClientInfo();

                MessageBox.Show($"Subscription renewed to {selectedBundle} for {selectedDuration} month(s).");
            }
        }

        private void btnAddExtra_Click(object sender, RoutedEventArgs e)
        {
            var AddDaysDialog = new AddExtraDaysDialoge();
            if (AddDaysDialog.ShowDialog() == true)
            {
                int days = AddDaysDialog.SelectedDays;
                ExcelHelper.AddExtraDays(client.PhoneNumber, days);
                DisplayClientInfo();
                MessageBox.Show($"Client's days added by {days} day(s).");
            }
        }
    }
}
