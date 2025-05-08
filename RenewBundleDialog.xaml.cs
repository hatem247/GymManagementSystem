﻿using System;
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
    /// Interaction logic for RenewBundleDialog.xaml
    /// </summary>
    public partial class RenewBundleDialog : Window
    {
        public RenewBundleDialog()
        {
            InitializeComponent();
            SessionstypeBox.Items.Add("");
            SessionstypeBox.Items.Add("45");
            SessionstypeBox.Items.Add("90");
            SessionstypeBox.Items.Add("180");
            SessionstypeBox.SelectedIndex = 0;
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
