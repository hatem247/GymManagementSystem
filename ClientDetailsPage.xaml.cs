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
            txtName.Text = client.FullName;
            txtPhone.Text = client.PhoneNumber;
            txtWeight.Text = client.Weight.ToString("F1");
            txtSubscription.Text = client.SubscriptionType;
            txtStart.Text = client.SubscriptionStart.ToShortDateString();
            txtEnd.Text = client.SubscriptionEnd.ToShortDateString();

            int daysLeft = (client.SubscriptionEnd - DateTime.Today).Days;
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
            var logs = ExcelHelper.GetLogs();
            var filtered = logs.Where(l => l.Phone == phoneNumber).ToList();
            LogsDataGrid.ItemsSource = filtered;
        }

        private void FreezeButton_Click(object sender, RoutedEventArgs e)
        {
            if (client.IsFrozen)
            {
                // Unfreeze the client
                ExcelHelper.UnfreezeClient(client.PhoneNumber);
                MessageBox.Show("Client has been unfrozen.");
                client.IsFrozen = false;
                btnFreeze.Content = "Freeze";
            }
            else
            {
                // Open freeze dialog
                var freezeDialog = new FreezeDialog();
                if (freezeDialog.ShowDialog() == true)
                {
                    int days = freezeDialog.SelectedDays;
                    ExcelHelper.FreezeClient(client.PhoneNumber, days);
                    MessageBox.Show($"Client frozen for {days} day(s).");
                    client.IsFrozen = true;
                    btnFreeze.Content = "Continue";
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
    }
}
