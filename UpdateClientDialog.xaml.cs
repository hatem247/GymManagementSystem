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
            txtAge.Text = client.Age.ToString();
            txtWeight.Text = client.Weight.ToString("F1");
            txtHeight.Text = client.Height.ToString("F1");
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            client.FullName = txtName.Text;
            client.Age = int.Parse(txtAge.Text);
            client.Weight = float.Parse(txtWeight.Text);
            client.Height = float.Parse(txtHeight.Text);
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
