using System.Windows;

namespace GymManagementSystem
{
    public partial class FreezeDialog : Window
    {
        public int SelectedDays { get; private set; }

        public FreezeDialog()
        {
            InitializeComponent();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
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