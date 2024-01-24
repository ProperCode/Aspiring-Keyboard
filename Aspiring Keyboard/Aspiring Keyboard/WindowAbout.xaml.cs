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
                MessageBox.Show(ex.Message, "Error WC001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Beula_Click(object sender, RoutedEventArgs e)
        {
            WindowEULA w = new WindowEULA();
            w.Owner = Application.Current.MainWindow;
            w.ShowInTaskbar = false;
            w.ShowDialog();
        }

        private void Lhomepage_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            Process.Start("https://" + Lhomepage.Content.ToString());
        }

        private void Lhomepage_MouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
        }

        private void Lhomepage_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = null;
        }
    }
}
