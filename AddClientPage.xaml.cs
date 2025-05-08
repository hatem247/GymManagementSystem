using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using ZXing;
using Microsoft.Win32;
using System.Windows.Input;

namespace GymManagementSystem
{
    public partial class AddClientPage : Page
    {
        string phone = "";
        public AddClientPage()
        {
            InitializeComponent();
            SessiontypeBox.Items.Add("");
            SessiontypeBox.Items.Add("45");
            SessiontypeBox.Items.Add("90");
            SessiontypeBox.Items.Add("180");
            SessiontypeBox.SelectedIndex = 0;
        }

        // When the "Save Client" button is clicked, this method is called
        private void SaveClient_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FullNameBox.Text) || string.IsNullOrEmpty(WeightBox.Text)
                || BundleBox.SelectedItem == null || SubscipriontypeBox.SelectedItem == null || string.IsNullOrEmpty(PhoneNumberBox.Text))
            {
                StatusText.Text = "Please fill in all fields.";
                StatusText.Foreground = System.Windows.Media.Brushes.Red;
                return;
            }

            string fullName = FullNameBox.Text;
            string weight = WeightBox.Text;
            string bundle = ((ComboBoxItem)BundleBox.SelectedItem)?.Content.ToString();
            string phoneNumber = PhoneNumberBox.Text;
            string subscriptiontype = ((ComboBoxItem)SubscipriontypeBox.SelectedItem)?.Content.ToString();
            int sessions = SessiontypeBox.SelectedIndex == 0 ? 0 : int.Parse(SessiontypeBox.Text);

            if((((bundle == "3 Months" && sessions != 180) || (bundle == "6 Months" && sessions != 45)) && sessions != 0) || ((bundle == "15 Days" || bundle == "1 Month") && sessions == 0))
            {
                var confirmDialoge = new ConfirmDialog();
                string sub = bundle + " " + subscriptiontype + " " + sessions + " Sessions";
                int total = ExcelHelper.GetAmount(sub);
                confirmDialoge.Messagetxt.Text = $"Client have to pay {total} EGP";
                confirmDialoge.ShowDialog();
                if (confirmDialoge.DialogResult == true)
                {
                    if (ExcelHelper.AddClient(fullName, weight, bundle, sessions, subscriptiontype, phoneNumber))
                    {
                        ExcelHelper.AddIncomeEntry(fullName, phoneNumber, sub);
                        ExcelHelper.AddLogEntry(fullName, phoneNumber);
                        StatusText.Text = "Client added successfully!";
                        GenerateBarcode(phoneNumber);
                        phone = phoneNumber;
                        Savebtn.Visibility = Visibility.Visible;
                        FullNameBox.Clear();
                        WeightBox.Clear();
                        PhoneNumberBox.Clear();
                        SubscipriontypeBox.SelectedIndex = -1;
                        BundleBox.SelectedIndex = -1;
                        SessiontypeBox.SelectedIndex = -1;
                    }
                }
            }
            else MessageBox.Show("Invlid bundle selection");
        }

        private void GenerateBarcode(string phone)
        {
            // Initialize the barcode writer from ZXing
            var barcodeWriter = new BarcodeWriter
            {
                Format = BarcodeFormat.CODE_128,
                Options = new ZXing.Common.EncodingOptions
                {
                    Height = 100,
                    Width = 300
                }
            };

            var barcodeBitmap = barcodeWriter.Write(phone);
            BarcodeImage.Source = ConvertBitmapToBitmapImage(barcodeBitmap);
        }

        public BitmapImage ConvertBitmapToBitmapImage(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, ImageFormat.Bmp);
                memory.Position = 0;

                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }

        private void SaveBarcode_Click(object sender, RoutedEventArgs e)
        {
            if (BarcodeImage.Source is BitmapSource bitmap)
            {
                if (string.IsNullOrEmpty(phone))
                {
                    MessageBox.Show("Phone number is required to name the file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "PNG Image|*.png";
                saveFileDialog.FileName = $"{phone}.png";

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (FileStream stream = new FileStream(saveFileDialog.FileName, FileMode.Create))
                    {
                        PngBitmapEncoder encoder = new PngBitmapEncoder();
                        encoder.Frames.Add(BitmapFrame.Create(bitmap));
                        encoder.Save(stream);
                    }

                    MessageBox.Show("Barcode saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            else
            {
                MessageBox.Show("No barcode image to save.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService.CanGoBack)
            {
                NavigationService.GoBack();
            }
            else
            {
                NavigationService.Navigate(new HomePage());
            }
        }
    }
}