using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;

namespace GymManagementSystem
{
    public partial class HomePage : Page
    {
        public HomePage()
        {
            InitializeComponent();
        }

        private void ScanBarcode_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ScanBarcodePage());
        }

        private void AddClient_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddClientPage());
        }

        private void ViewClients_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ClientsListPage());
        }
        
        private void Logs_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new LogsPage());
        }
        
        private void Income_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new IncomePage());
        }

        private void Back_Click(object sender, MouseButtonEventArgs e)
        {
            NavigationService.Navigate(new LoginPage());
        }
    }
}