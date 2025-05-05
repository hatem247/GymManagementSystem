using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using ZXing;

namespace GymManagementSystem
{
    public partial class AddClientPage : Page
    {
        public AddClientPage()
        {
            InitializeComponent();
        }

        // When the "Save Client" button is clicked, this method is called
        private void SaveClient_Click(object sender, RoutedEventArgs e)
        {
            // Validate the input fields
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

            ExcelHelper.AddClient(fullName, weight, bundle, subscriptiontype, phoneNumber);
            ExcelHelper.AddIncomeEntry(Name, phoneNumber, bundle + " " + subscriptiontype);
            StatusText.Text = "Client added successfully!";

            // Generate and display the barcode 
            GenerateBarcode(phoneNumber);
            FullNameBox.Clear();
            WeightBox.Clear();
            PhoneNumberBox.Clear();
            BundleBox.SelectedIndex = -1;
        }
        private void GenerateBarcode(string phone)
        {
            // Initialize the barcode writer from ZXing
            var barcodeWriter = new BarcodeWriter
            {
                Format = BarcodeFormat.CODE_128, // You can choose other formats if needed
                Options = new ZXing.Common.EncodingOptions
                {
                    Height = 100, // Adjust the height of the barcode
                    Width = 300   // Adjust the width of the barcode
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