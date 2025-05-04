using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace GymManagementSystem
{
    public partial class LoginPage : Page
    {
        public LoginPage()
        {
            InitializeComponent();
            ExcelHelper.AutoUnfreezeClients();
        }

        private void Login_Click(object sender, RoutedEventArgs e)
        {
            string username = UsernameTextBox.Text.Trim();
            string password = PasswordBox.Password.Trim();

            if (username == "admin" && password == "admin")
            {
                NavigationService.Navigate(new HomePage());
            }
            else
            {
                ErrorText.Text = "Invalid credentials!";
            }
        }
    }
}
