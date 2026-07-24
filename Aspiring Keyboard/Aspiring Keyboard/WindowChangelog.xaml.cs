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

                TB.Text = "[1.4] - July 24, 2026:"
                + "\n- Changed default desired figures number to 2500."
                + "\n- Improved Smart Mousegrid."
                + "\n- Removed 2 characters from Mousegrid alphabet for any keyboard layout."
                + "\n- Removed Center Left Click action."
                + "\n- Removed LAlt + RAlt hotkey."
                + "\n\n[1.3] - August 5, 2024:"
                + "\n- Added separate mouse buttons releasing."
                + "\n- Changed LAlt + RAlt combination."
                + "\n- Fixed settings loading."
                + "\n\n[1.2] - July 30, 2024:"
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
