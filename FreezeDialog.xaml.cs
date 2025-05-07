using System.Windows;

namespace GymManagementSystem
{
    public partial class FreezeDialog : Window
    {
        public int SelectedDays { get; private set; }

        public FreezeDialog()
        {
            InitializeComponent();
            for (int i = 10; i <= 40; i++)
            {
                DaysComboBox.Items.Add(i);
            }
            DaysComboBox.SelectedIndex = 0;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SelectedDays = (int)DaysComboBox.SelectedItem;
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