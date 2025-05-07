using System;
using System.Windows;
using System.Windows.Navigation;

namespace GymManagementSystem
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            var navWindow = new NavigationWindow
            {
                Source = new Uri("LoginPage.xaml", UriKind.Relative)
            };
            navWindow.Show();
        }
    }
}