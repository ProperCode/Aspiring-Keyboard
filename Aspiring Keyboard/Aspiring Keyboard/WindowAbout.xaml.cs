using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;

namespace Aspiring_Keyboard
{
    /// <summary>
    /// Interaction logic for WindowAbout.xaml
    /// </summary>
    public partial class WindowAbout : Window
    {
        public WindowAbout()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WA001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Beula_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WindowEULA w = new WindowEULA();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WA002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Lhomepage_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Process.Start("https://" + Lhomepage.Content.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WA003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Lhomepage_MouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
        }

        private void Lhomepage_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = null;
        }

        private void Bchangelog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WindowChangelog w = new WindowChangelog();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WA004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bthird_party_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("License\\Windows Input Simulator License.txt");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WA005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}