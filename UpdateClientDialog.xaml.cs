using System;
using System.Windows;

namespace GymManagementSystem
{
    public partial class UpdateClientDialog : Window
    {
        private Client client;

        public UpdateClientDialog(Client clientData)
        {
            InitializeComponent();
            client = clientData;
            LoadClientData();
        }

        private void LoadClientData()
        {
            txtName.Text = client.FullName;
            txtWeight.Text = client.Weight.ToString("F1");
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            client.FullName = txtName.Text;
            client.Weight = float.Parse(txtWeight.Text);
            ExcelHelper.EditClient(client);
            MessageBox.Show("Client details updated successfully.");
            this.DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            Close();
        }
    }
}
