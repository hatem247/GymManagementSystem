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
    /// Interaction logic for DeleteDialoge.xaml
    /// </summary>
    public partial class DeleteDialoge : Window
    {
        public string phoneNumber;
        public DeleteDialoge(string phoneNumber)
        {
            InitializeComponent();
            this.phoneNumber = phoneNumber;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            Close();
        }
        
        void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            ExcelHelper.DeleteClient(phoneNumber);
            MessageBox.Show("Client deleted successfully");

            NavigationWindow navigationWindow = Application.Current.MainWindow as NavigationWindow;
            if (navigationWindow != null)
            {
                navigationWindow.Navigate(new HomePage());
            }

            this.DialogResult = false;
            Close();
        }
    }
}
