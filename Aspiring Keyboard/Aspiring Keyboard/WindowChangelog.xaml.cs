using System;
using System.Windows;

namespace Aspiring_Keyboard
{
    /// <summary>
    /// Interaction logic for WindowChangelog.xaml
    /// </summary>
    public partial class WindowChangelog : Window
    {
        public WindowChangelog()
        {
            try
            {
                InitializeComponent();

                TB.IsReadOnly = true;

                TB.Text = "All notable changes to Aspiring Keyboard will be documented here."
                + "\n\n[1.2] - July 30 , 2024:"
                + "\n- Fixed settings saving and loading."
                + "\n\n[1.1] - Januray 28, 2024:"
                + "\n- Fixed minor bugs.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WC001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
