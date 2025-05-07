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
using System.Windows.Shapes;

namespace GymManagementSystem
{
    /// <summary>
    /// Interaction logic for AddExtraDaysDiologe.xaml
    /// </summary>
    public partial class AddExtraDaysDialoge : Window
    {
        public AddExtraDaysDialoge()
        {
            InitializeComponent();
            for (int i = 10; i <= 40; i++)
            {
                ExtraDaysComboBox.Items.Add(i);
            }
            ExtraDaysComboBox.SelectedIndex = 0;
        }
        public int SelectedDays { get; private set; }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SelectedDays = (int)ExtraDaysComboBox.SelectedItem;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}