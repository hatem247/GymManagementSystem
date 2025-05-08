using System.Windows.Controls;
using System.Windows.Input;
using System.Collections.Generic;
using System.Windows;
using System.Linq;

namespace GymManagementSystem
{
    public partial class ClientsListPage : Page
    {
        public ClientsListPage()
        {
            InitializeComponent();
            LoadClients();
        }

        private void LoadClients()
        {
            var clients = ExcelHelper.LoadAllClients();
            ClientsList.ItemsSource = clients;
        }

        private void ClientsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ClientsList.SelectedItem is Client selectedClient)
            {
                NavigationService.Navigate(new ClientDetailsPage(selectedClient));
            }
        }
        private void Search_Click(object sender, RoutedEventArgs e)
        {
            string phone = SearchInput.Text.Trim();

            if (string.IsNullOrEmpty(phone))
            {
                MessageBox.Show("Please enter a phone number.");
                return;
            }

            Client client = ExcelHelper.Search(phone);
            if (client != null)
            {
                NavigationService.Navigate(new ClientDetailsPage(client));
                SearchInput.Clear();
            }
            else
            {
                MessageBox.Show("Client not found.");
                SearchInput.SelectAll();
            }
        }
        private void Back_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new HomePage());
        }
    }
}