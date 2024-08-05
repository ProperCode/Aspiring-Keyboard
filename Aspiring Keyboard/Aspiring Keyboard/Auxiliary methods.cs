//highest error nr: AM017
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Aspiring_Keyboard
{
    public partial class MainWindow : Window
    {
        //------------------KEYPRESSES DETECTION----------------------
        [DllImport("User32.dll")]
        private static extern short GetAsyncKeyState(System.Windows.Forms.Keys vKey); // Keys enumeration

        [DllImport("User32.dll")]
        private static extern short GetAsyncKeyState(System.Int32 vKey);

        public static bool IsKeyPushedDown(System.Windows.Forms.Keys vKey)
        {
            return 0 != (GetAsyncKeyState((int)vKey) & 0x8000);
        }

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName,
            string lpWindowName);

        System.Windows.Forms.NotifyIcon ni = new System.Windows.Forms.NotifyIcon();

        System.Windows.Forms.ColorDialog colorDialog1 = new System.Windows.Forms.ColorDialog();
        System.Windows.Forms.ColorDialog colorDialog2 = new System.Windows.Forms.ColorDialog();

        bool loading = false;

        public int GetWindowsScaling()
        {
            return (int)(100 * System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width
                / SystemParameters.PrimaryScreenWidth);
        }

        void restore_default_settings()
        {
            CBlshift_action_a.SelectedItem = "Left click";
            CBrshift_action_a.SelectedItem = "Ctrl left click";
            CBlalt_action_a.SelectedItem = "Drag and drop";
            CBralt_action_a.SelectedItem = "Right click";
            CBlctrl_action_a.SelectedItem = "Double left click";
            CBrctrl_action_a.SelectedItem = "Triple left click";

            CBlshift_action_b.SelectedItem = "Left click";
            CBrshift_action_b.SelectedItem = "Ctrl left click";
            CBlalt_action_b.SelectedItem = "Hold left";
            CBralt_action_b.SelectedItem = "Release left";
            CBlctrl_action_b.SelectedItem = "Double left click";
            CBrctrl_action_b.SelectedItem = "Move mouse";

            CBtype.SelectedItem = "Square horizontal precision";
            CBlines.SelectedItem = "None";
            TBdesired_figures_nr.Text = default_desired_figures_nr.ToString();
            
            color_bg_str = default_color_bg_str;
            color_font_str = default_color_font_str;

            int argb = Convert.ToInt32(color_bg_str);

            byte[] values = BitConverter.GetBytes(argb);

            byte a = values[3];
            byte r = values[2];
            byte g = values[1];
            byte b = values[0];

            TBbackground_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));
            
            argb = Convert.ToInt32(color_font_str);

            values = BitConverter.GetBytes(argb);

            a = values[3];
            r = values[2];
            g = values[1];
            b = values[0];

            TBfont_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));
            
            TBfont_size.Text = "12";
            CHBsmart_mousegrid.IsChecked = true;

            CBkeyboard_layout.SelectedIndex = 0; //Any layout

            CHBrun_at_startup.IsChecked = false;
            CHBstart_minimized.IsChecked = false;
            CHBminimize_to_tray.IsChecked = false;
            CHBcheck_for_updates.IsChecked = false;
            
            bool found = false;
            for (int i = 0; i < ss_voices_priority_list.Count && found == false; i++)
            {
                foreach (InstalledVoice iv in installed_voices)
                {
                    if (iv.VoiceInfo.Name.Contains(ss_voices_priority_list[i]))
                    {
                        default_ss_voice = iv.VoiceInfo.Name;
                        found = true;
                        break;
                    }
                }
            }

            if (CBss_voices.Items.Count > 0)
            {
                if (default_ss_voice == "")
                {
                    CBss_voices.SelectedIndex = 0;
                    default_ss_voice = CBss_voices.SelectedItem.ToString();
                }
                else
                {
                    CBss_voices.SelectedItem = default_ss_voice;
                }
            }
            ss_voice = default_ss_voice;

            TBss_volume.Text = default_ss_volume.ToString();
            CHBread_status.IsChecked = true;
        }

        private void CenterWindowOnScreen()
        {
            double screenWidth = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenWidth);
            double screenHeight = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenHeight);
            double windowWidth = this.Width;
            double windowHeight = this.Height;
            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);
        }

        private void Brestore_default_settings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                restore_default_settings();

                save_settings();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void ni_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            try
            {
                Show();
                this.WindowState = WindowState.Normal;
                SetForegroundWindow(new System.Windows.Interop.WindowInteropHelper(this).Handle);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void mi_switch_mode_Click(object sender, EventArgs e)
        {
            try
            {
                fix_very_rare_problem();
                string icon;

                if (green_mode)
                {
                    green_mode = false;
                    icon = icon_blue;

                    if (read_status) ss.SpeakAsync("Blue mode");
                }
                else
                {
                    green_mode = true;
                    icon = icon_green;

                    if (read_status) ss.SpeakAsync("Green mode");
                }

                this.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                {
                    Stream iconStream = System.Windows.Application.GetResourceStream(
                    new Uri(icon)).Stream;

                    // Create a BitmapSource  
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(icon);
                    bitmap.EndInit();

                    this.Icon = bitmap;
                    ni.Icon = new System.Drawing.Icon(iconStream);

                    iconStream.Close();
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM014", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void mi_toggle_Click(object sender, EventArgs e)
        {
            try
            {
                fix_very_rare_problem();
                string icon;

                if (enabled)
                {
                    enabled = false;
                    if (read_status) ss.SpeakAsync("disabled");

                    icon = icon_red;
                }
                else
                {
                    enabled = true;
                    if (read_status) ss.SpeakAsync("enabled");

                    if (green_mode)
                        icon = icon_green;
                    else
                        icon = icon_blue;
                }

                this.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                {
                    Stream iconStream = System.Windows.Application.GetResourceStream(
                    new Uri(icon)).Stream;

                    // Create a BitmapSource  
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(icon);
                    bitmap.EndInit();

                    this.Icon = bitmap;
                    ni.Icon = new System.Drawing.Icon(iconStream);

                    iconStream.Close();
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM015", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void mi_exit_Click(object sender, EventArgs e)
        {
            try
            {
                Window_Closing(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM016", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                ni.Visible = false;
                ni.Dispose();

                release_buttons_and_keys();
                if (smart_grid)
                    save_grids();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            Process.GetCurrentProcess().Kill();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                release_buttons_and_keys();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void update_app_if_necessary()
        {
            try
            {
                string content;
                MyWebClient wc = new MyWebClient();
                content = wc.DownloadString(url_latest_version);

                latest_version = content.Replace("\r\n", "").Trim();
            }
            catch (WebException we)
            {
                latest_version = "unknown";
            }

            bool update_available = false;

            if (latest_version != "unknown" &&
                int.Parse(latest_version.Replace(".", "")) > int.Parse(prog_version.Replace(".", "")))
            {
                update_available = true;
            }

            if ((bool)CHBcheck_for_updates.IsChecked && update_available)
            {
                MessageBoxResult dialogResult = System.Windows.MessageBox.Show("A new program version" +
                    " is available. Do you want to download it now?",
                //    " is available. Do you want to perform an automatic update now?",
                    "New Version Available", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (dialogResult == MessageBoxResult.Yes)
                {
                    //Open download page
                    Process.Start("https://" + url_homepage);
                }
            }
        }

        void TBbackground_color_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                System.Windows.Forms.DialogResult dr = colorDialog1.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    int argb = Convert.ToInt32(colorDialog1.Color.ToArgb().ToString());
                    color_bg_str = argb.ToString();

                    byte[] values = BitConverter.GetBytes(argb);

                    byte a = values[3];
                    byte r = values[2];
                    byte g = values[1];
                    byte b = values[0];

                    TBbackground_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TBfont_color_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                System.Windows.Forms.DialogResult dr = colorDialog2.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    int argb = Convert.ToInt32(colorDialog2.Color.ToArgb().ToString());
                    color_font_str = argb.ToString();

                    byte[] values = BitConverter.GetBytes(argb);

                    byte a = values[3];
                    byte r = values[2];
                    byte g = values[1];
                    byte b = values[0];

                    TBfont_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));
                    color_font = Color.FromArgb(a, r, g, b);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CBtype_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CBtype.SelectedItem.ToString() == "Hexagonal")
            {
                CBlines.SelectedIndex = 0;
                CBlines.IsEnabled = false;
            }
            else
            {
                CBlines.IsEnabled = true;
            }
        }

        private void CBkeyboard_layout_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (loading == false)
            {
                if (CBkeyboard_layout.SelectedItem.ToString() == "US English / US International")
                    create_grid_alphabet_for_US_kl();
                else
                    create_grid_alphabet_for_any_kl();

                grid_symbols_limit = grid_alphabet.Count;
                max_figures_nr = (int)Math.Pow((double)grid_symbols_limit, 2);

                TBdesired_figures_nr.Text = max_figures_nr.ToString();
            }
        }

        private void CHBsmart_mousegrid_Checked(object sender, RoutedEventArgs e)
        { /* needed for proper loading */ }

        private void CHBrun_at_startup_Checked(object sender, RoutedEventArgs e)
        { /* needed for proper loading */ }

        private void CHBstart_minimized_Checked(object sender, RoutedEventArgs e)
        { /* needed for proper loading */ }

        private void CHBminimize_to_tray_Checked(object sender, RoutedEventArgs e)
        { /* needed for proper loading */ }

        private void CHBread_status_Checked(object sender, RoutedEventArgs e)
        { /* needed for proper loading */ }

        private void CHBcheck_for_updates_Checked(object sender, RoutedEventArgs e)
        { /* needed for proper loading */ }

        private void W_StateChanged(object sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Minimized && CHBminimize_to_tray.IsChecked == true)
            {
                this.Hide();
            }
        }

        private void Bsave_settings_Click(object sender, RoutedEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Wait;

            save_settings();

            Mouse.OverrideCursor = null;

            if (auto_updates)
                update_app_if_necessary();
        }

        private void MIexit_Click(object sender, RoutedEventArgs e)
        {
            Window_Closing(null, null);
        }

        private void MIuser_guide_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("User Guide.pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM011", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIuseful_key_combinations_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("Useful Windows Key Combinations.pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM012", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIabout_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string content;
                MyWebClient wc = new MyWebClient();
                content = wc.DownloadString(url_latest_version);

                latest_version = content.Replace("\r\n", "").Trim();
            }
            catch (WebException we)
            {
                latest_version = "unknown";
            }

            try
            {
                WindowAbout w = new WindowAbout();

                w.Lprogram_name.Content = prog_name;
                w.Llatest_version.Content = "Latest version: " + latest_version;
                w.Linstalled_version.Content = "Installed version: " + prog_version;
                w.Lhomepage.Content = url_homepage;
                w.Lcopyright.Content = copyright_text;

                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM013", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIchangelog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WindowChangelog w = new WindowChangelog();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM017", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void set_values()
        {
            if (saving_enabled)
            {
                if (smart_grid && desired_figures_nr >= 5)
                    save_grids();
            }

            if (CBlshift_action_a.SelectedIndex == 0)
                left_shift_action_a = ActionX.left_click;
            else if (CBlshift_action_a.SelectedIndex == 1)
                left_shift_action_a = ActionX.right_click;
            else if (CBlshift_action_a.SelectedIndex == 2)
                left_shift_action_a = ActionX.double_left_click;
            else if (CBlshift_action_a.SelectedIndex == 3)
                left_shift_action_a = ActionX.triple_left_click;
            else if (CBlshift_action_a.SelectedIndex == 4)
                left_shift_action_a = ActionX.center_left_click;
            else if (CBlshift_action_a.SelectedIndex == 5)
                left_shift_action_a = ActionX.ctrl_left_click;
            else if (CBlshift_action_a.SelectedIndex == 6)
                left_shift_action_a = ActionX.move_mouse;
            else if (CBlshift_action_a.SelectedIndex == 7)
                left_shift_action_a = ActionX.drag_and_drop;
            else if (CBlshift_action_a.SelectedIndex == 8)
                left_shift_action_a = ActionX.hold_left;
            else if (CBlshift_action_a.SelectedIndex == 9)
                left_shift_action_a = ActionX.hold_right;
            else if (CBlshift_action_a.SelectedIndex == 10)
                left_shift_action_a = ActionX.release_left;
            else if (CBlshift_action_a.SelectedIndex == 11)
                left_shift_action_a = ActionX.release_right;
            else if (CBlshift_action_a.SelectedIndex == 12)
                left_shift_action_a = ActionX.none;

            if (CBrshift_action_a.SelectedIndex == 0)
                right_shift_action_a = ActionX.left_click;
            else if (CBrshift_action_a.SelectedIndex == 1)
                right_shift_action_a = ActionX.right_click;
            else if (CBrshift_action_a.SelectedIndex == 2)
                right_shift_action_a = ActionX.double_left_click;
            else if (CBrshift_action_a.SelectedIndex == 3)
                right_shift_action_a = ActionX.triple_left_click;
            else if (CBrshift_action_a.SelectedIndex == 4)
                right_shift_action_a = ActionX.center_left_click;
            else if (CBrshift_action_a.SelectedIndex == 5)
                right_shift_action_a = ActionX.ctrl_left_click;
            else if (CBrshift_action_a.SelectedIndex == 6)
                right_shift_action_a = ActionX.move_mouse;
            else if (CBrshift_action_a.SelectedIndex == 7)
                right_shift_action_a = ActionX.drag_and_drop;
            else if (CBrshift_action_a.SelectedIndex == 8)
                right_shift_action_a = ActionX.hold_left;
            else if (CBrshift_action_a.SelectedIndex == 9)
                right_shift_action_a = ActionX.hold_right;
            else if (CBrshift_action_a.SelectedIndex == 10)
                right_shift_action_a = ActionX.release_left;
            else if (CBrshift_action_a.SelectedIndex == 11)
                right_shift_action_a = ActionX.release_right;
            else if (CBrshift_action_a.SelectedIndex == 12)
                right_shift_action_a = ActionX.none;

            if (CBlshift_action_b.SelectedIndex == 0)
                left_shift_action_b = ActionX.left_click;
            else if (CBlshift_action_b.SelectedIndex == 1)
                left_shift_action_b = ActionX.right_click;
            else if (CBlshift_action_b.SelectedIndex == 2)
                left_shift_action_b = ActionX.double_left_click;
            else if (CBlshift_action_b.SelectedIndex == 3)
                left_shift_action_b = ActionX.triple_left_click;
            else if (CBlshift_action_b.SelectedIndex == 4)
                left_shift_action_b = ActionX.center_left_click;
            else if (CBlshift_action_b.SelectedIndex == 5)
                left_shift_action_b = ActionX.ctrl_left_click;
            else if (CBlshift_action_b.SelectedIndex == 6)
                left_shift_action_b = ActionX.move_mouse;
            else if (CBlshift_action_b.SelectedIndex == 7)
                left_shift_action_b = ActionX.drag_and_drop;
            else if (CBlshift_action_b.SelectedIndex == 8)
                left_shift_action_b = ActionX.hold_left;
            else if (CBlshift_action_b.SelectedIndex == 9)
                left_shift_action_b = ActionX.hold_right;
            else if (CBlshift_action_b.SelectedIndex == 10)
                left_shift_action_b = ActionX.release_left;
            else if (CBlshift_action_b.SelectedIndex == 11)
                left_shift_action_b = ActionX.release_right;
            else if (CBlshift_action_b.SelectedIndex == 12)
                left_shift_action_b = ActionX.none;

            if (CBrshift_action_b.SelectedIndex == 0)
                right_shift_action_b = ActionX.left_click;
            else if (CBrshift_action_b.SelectedIndex == 1)
                right_shift_action_b = ActionX.right_click;
            else if (CBrshift_action_b.SelectedIndex == 2)
                right_shift_action_b = ActionX.double_left_click;
            else if (CBrshift_action_b.SelectedIndex == 3)
                right_shift_action_b = ActionX.triple_left_click;
            else if (CBrshift_action_b.SelectedIndex == 4)
                right_shift_action_b = ActionX.center_left_click;
            else if (CBrshift_action_b.SelectedIndex == 5)
                right_shift_action_b = ActionX.ctrl_left_click;
            else if (CBrshift_action_b.SelectedIndex == 6)
                right_shift_action_b = ActionX.move_mouse;
            else if (CBrshift_action_b.SelectedIndex == 7)
                right_shift_action_b = ActionX.drag_and_drop;
            else if (CBrshift_action_b.SelectedIndex == 8)
                right_shift_action_b = ActionX.hold_left;
            else if (CBrshift_action_b.SelectedIndex == 9)
                right_shift_action_b = ActionX.hold_right;
            else if (CBrshift_action_b.SelectedIndex == 10)
                right_shift_action_b = ActionX.release_left;
            else if (CBrshift_action_b.SelectedIndex == 11)
                right_shift_action_b = ActionX.release_right;
            else if (CBrshift_action_b.SelectedIndex == 12)
                right_shift_action_b = ActionX.none;

            if (CBlalt_action_a.SelectedIndex == 0)
                left_alt_action_a = ActionX.left_click;
            else if (CBlalt_action_a.SelectedIndex == 1)
                left_alt_action_a = ActionX.right_click;
            else if (CBlalt_action_a.SelectedIndex == 2)
                left_alt_action_a = ActionX.double_left_click;
            else if (CBlalt_action_a.SelectedIndex == 3)
                left_alt_action_a = ActionX.triple_left_click;
            else if (CBlalt_action_a.SelectedIndex == 4)
                left_alt_action_a = ActionX.center_left_click;
            else if (CBlalt_action_a.SelectedIndex == 5)
                left_alt_action_a = ActionX.ctrl_left_click;
            else if (CBlalt_action_a.SelectedIndex == 6)
                left_alt_action_a = ActionX.move_mouse;
            else if (CBlalt_action_a.SelectedIndex == 7)
                left_alt_action_a = ActionX.drag_and_drop;
            else if (CBlalt_action_a.SelectedIndex == 8)
                left_alt_action_a = ActionX.hold_left;
            else if (CBlalt_action_a.SelectedIndex == 9)
                left_alt_action_a = ActionX.hold_right;
            else if (CBlalt_action_a.SelectedIndex == 10)
                left_alt_action_a = ActionX.release_left;
            else if (CBlalt_action_a.SelectedIndex == 11)
                left_alt_action_a = ActionX.release_right;
            else if (CBlalt_action_a.SelectedIndex == 12)
                left_alt_action_a = ActionX.none;

            if (CBralt_action_a.SelectedIndex == 0)
                right_alt_action_a = ActionX.left_click;
            else if (CBralt_action_a.SelectedIndex == 1)
                right_alt_action_a = ActionX.right_click;
            else if (CBralt_action_a.SelectedIndex == 2)
                right_alt_action_a = ActionX.double_left_click;
            else if (CBralt_action_a.SelectedIndex == 3)
                right_alt_action_a = ActionX.triple_left_click;
            else if (CBralt_action_a.SelectedIndex == 4)
                right_alt_action_a = ActionX.center_left_click;
            else if (CBralt_action_a.SelectedIndex == 5)
                right_alt_action_a = ActionX.ctrl_left_click;
            else if (CBralt_action_a.SelectedIndex == 6)
                right_alt_action_a = ActionX.move_mouse;
            else if (CBralt_action_a.SelectedIndex == 7)
                right_alt_action_a = ActionX.drag_and_drop;
            else if (CBralt_action_a.SelectedIndex == 8)
                right_alt_action_a = ActionX.hold_left;
            else if (CBralt_action_a.SelectedIndex == 9)
                right_alt_action_a = ActionX.hold_right;
            else if (CBralt_action_a.SelectedIndex == 10)
                right_alt_action_a = ActionX.release_left;
            else if (CBralt_action_a.SelectedIndex == 11)
                right_alt_action_a = ActionX.release_right;
            else if (CBralt_action_a.SelectedIndex == 12)
                right_alt_action_a = ActionX.none;

            if (CBlalt_action_b.SelectedIndex == 0)
                left_alt_action_b = ActionX.left_click;
            else if (CBlalt_action_b.SelectedIndex == 1)
                left_alt_action_b = ActionX.right_click;
            else if (CBlalt_action_b.SelectedIndex == 2)
                left_alt_action_b = ActionX.double_left_click;
            else if (CBlalt_action_b.SelectedIndex == 3)
                left_alt_action_b = ActionX.triple_left_click;
            else if (CBlalt_action_b.SelectedIndex == 4)
                left_alt_action_b = ActionX.center_left_click;
            else if (CBlalt_action_b.SelectedIndex == 5)
                left_alt_action_b = ActionX.ctrl_left_click;
            else if (CBlalt_action_b.SelectedIndex == 6)
                left_alt_action_b = ActionX.move_mouse;
            else if (CBlalt_action_b.SelectedIndex == 7)
                left_alt_action_b = ActionX.drag_and_drop;
            else if (CBlalt_action_b.SelectedIndex == 8)
                left_alt_action_b = ActionX.hold_left;
            else if (CBlalt_action_b.SelectedIndex == 9)
                left_alt_action_b = ActionX.hold_right;
            else if (CBlalt_action_b.SelectedIndex == 10)
                left_alt_action_b = ActionX.release_left;
            else if (CBlalt_action_b.SelectedIndex == 11)
                left_alt_action_b = ActionX.release_right;
            else if (CBlalt_action_b.SelectedIndex == 12)
                left_alt_action_b = ActionX.none;

            if (CBralt_action_b.SelectedIndex == 0)
                right_alt_action_b = ActionX.left_click;
            else if (CBralt_action_b.SelectedIndex == 1)
                right_alt_action_b = ActionX.right_click;
            else if (CBralt_action_b.SelectedIndex == 2)
                right_alt_action_b = ActionX.double_left_click;
            else if (CBralt_action_b.SelectedIndex == 3)
                right_alt_action_b = ActionX.triple_left_click;
            else if (CBralt_action_b.SelectedIndex == 4)
                right_alt_action_b = ActionX.center_left_click;
            else if (CBralt_action_b.SelectedIndex == 5)
                right_alt_action_b = ActionX.ctrl_left_click;
            else if (CBralt_action_b.SelectedIndex == 6)
                right_alt_action_b = ActionX.move_mouse;
            else if (CBralt_action_b.SelectedIndex == 7)
                right_alt_action_b = ActionX.drag_and_drop;
            else if (CBralt_action_b.SelectedIndex == 8)
                right_alt_action_b = ActionX.hold_left;
            else if (CBralt_action_b.SelectedIndex == 9)
                right_alt_action_b = ActionX.hold_right;
            else if (CBralt_action_b.SelectedIndex == 10)
                right_alt_action_b = ActionX.release_left;
            else if (CBralt_action_b.SelectedIndex == 11)
                right_alt_action_b = ActionX.release_right;
            else if (CBralt_action_b.SelectedIndex == 12)
                right_alt_action_b = ActionX.none;

            if (CBlctrl_action_a.SelectedIndex == 0)
                left_ctrl_action_a = ActionX.left_click;
            else if (CBlctrl_action_a.SelectedIndex == 1)
                left_ctrl_action_a = ActionX.right_click;
            else if (CBlctrl_action_a.SelectedIndex == 2)
                left_ctrl_action_a = ActionX.double_left_click;
            else if (CBlctrl_action_a.SelectedIndex == 3)
                left_ctrl_action_a = ActionX.triple_left_click;
            else if (CBlctrl_action_a.SelectedIndex == 4)
                left_ctrl_action_a = ActionX.center_left_click;
            else if (CBlctrl_action_a.SelectedIndex == 5)
                left_ctrl_action_a = ActionX.ctrl_left_click;
            else if (CBlctrl_action_a.SelectedIndex == 6)
                left_ctrl_action_a = ActionX.move_mouse;
            else if (CBlctrl_action_a.SelectedIndex == 7)
                left_ctrl_action_a = ActionX.drag_and_drop;
            else if (CBlctrl_action_a.SelectedIndex == 8)
                left_ctrl_action_a = ActionX.hold_left;
            else if (CBlctrl_action_a.SelectedIndex == 9)
                left_ctrl_action_a = ActionX.hold_right;
            else if (CBlctrl_action_a.SelectedIndex == 10)
                left_ctrl_action_a = ActionX.release_left;
            else if (CBlctrl_action_a.SelectedIndex == 11)
                left_ctrl_action_a = ActionX.release_right;
            else if (CBlctrl_action_a.SelectedIndex == 12)
                left_ctrl_action_a = ActionX.none;

            if (CBrctrl_action_a.SelectedIndex == 0)
                right_ctrl_action_a = ActionX.left_click;
            else if (CBrctrl_action_a.SelectedIndex == 1)
                right_ctrl_action_a = ActionX.right_click;
            else if (CBrctrl_action_a.SelectedIndex == 2)
                right_ctrl_action_a = ActionX.double_left_click;
            else if (CBrctrl_action_a.SelectedIndex == 3)
                right_ctrl_action_a = ActionX.triple_left_click;
            else if (CBrctrl_action_a.SelectedIndex == 4)
                right_ctrl_action_a = ActionX.center_left_click;
            else if (CBrctrl_action_a.SelectedIndex == 5)
                right_ctrl_action_a = ActionX.ctrl_left_click;
            else if (CBrctrl_action_a.SelectedIndex == 6)
                right_ctrl_action_a = ActionX.move_mouse;
            else if (CBrctrl_action_a.SelectedIndex == 7)
                right_ctrl_action_a = ActionX.drag_and_drop;
            else if (CBrctrl_action_a.SelectedIndex == 8)
                right_ctrl_action_a = ActionX.hold_left;
            else if (CBrctrl_action_a.SelectedIndex == 9)
                right_ctrl_action_a = ActionX.hold_right;
            else if (CBrctrl_action_a.SelectedIndex == 10)
                right_ctrl_action_a = ActionX.release_left;
            else if (CBrctrl_action_a.SelectedIndex == 11)
                right_ctrl_action_a = ActionX.release_right;
            else if (CBrctrl_action_a.SelectedIndex == 12)
                right_ctrl_action_a = ActionX.none;

            if (CBlctrl_action_b.SelectedIndex == 0)
                left_ctrl_action_b = ActionX.left_click;
            else if (CBlctrl_action_b.SelectedIndex == 1)
                left_ctrl_action_b = ActionX.right_click;
            else if (CBlctrl_action_b.SelectedIndex == 2)
                left_ctrl_action_b = ActionX.double_left_click;
            else if (CBlctrl_action_b.SelectedIndex == 3)
                left_ctrl_action_b = ActionX.triple_left_click;
            else if (CBlctrl_action_b.SelectedIndex == 4)
                left_ctrl_action_b = ActionX.center_left_click;
            else if (CBlctrl_action_b.SelectedIndex == 5)
                left_ctrl_action_b = ActionX.ctrl_left_click;
            else if (CBlctrl_action_b.SelectedIndex == 6)
                left_ctrl_action_b = ActionX.move_mouse;
            else if (CBlctrl_action_b.SelectedIndex == 7)
                left_ctrl_action_b = ActionX.drag_and_drop;
            else if (CBlctrl_action_b.SelectedIndex == 8)
                left_ctrl_action_b = ActionX.hold_left;
            else if (CBlctrl_action_b.SelectedIndex == 9)
                left_ctrl_action_b = ActionX.hold_right;
            else if (CBlctrl_action_b.SelectedIndex == 10)
                left_ctrl_action_b = ActionX.release_left;
            else if (CBlctrl_action_b.SelectedIndex == 11)
                left_ctrl_action_b = ActionX.release_right;
            else if (CBlctrl_action_b.SelectedIndex == 12)
                left_ctrl_action_b = ActionX.none;

            if (CBrctrl_action_b.SelectedIndex == 0)
                right_ctrl_action_b = ActionX.left_click;
            else if (CBrctrl_action_b.SelectedIndex == 1)
                right_ctrl_action_b = ActionX.right_click;
            else if (CBrctrl_action_b.SelectedIndex == 2)
                right_ctrl_action_b = ActionX.double_left_click;
            else if (CBrctrl_action_b.SelectedIndex == 3)
                right_ctrl_action_b = ActionX.triple_left_click;
            else if (CBrctrl_action_b.SelectedIndex == 4)
                right_ctrl_action_b = ActionX.center_left_click;
            else if (CBrctrl_action_b.SelectedIndex == 5)
                right_ctrl_action_b = ActionX.ctrl_left_click;
            else if (CBrctrl_action_b.SelectedIndex == 6)
                right_ctrl_action_b = ActionX.move_mouse;
            else if (CBrctrl_action_b.SelectedIndex == 7)
                right_ctrl_action_b = ActionX.drag_and_drop;
            else if (CBrctrl_action_b.SelectedIndex == 8)
                right_ctrl_action_b = ActionX.hold_left;
            else if (CBrctrl_action_b.SelectedIndex == 9)
                right_ctrl_action_b = ActionX.hold_right;
            else if (CBrctrl_action_b.SelectedIndex == 10)
                right_ctrl_action_b = ActionX.release_left;
            else if (CBrctrl_action_b.SelectedIndex == 11)
                right_ctrl_action_b = ActionX.release_right;
            else if (CBrctrl_action_b.SelectedIndex == 12)
                right_ctrl_action_b = ActionX.none;

            bool grid_size_changed = false;
            GridType prev_grid_type = grid_type;

            if (CBtype.SelectedIndex == 0)
                grid_type = GridType.hexagonal;
            else if (CBtype.SelectedIndex == 1)
                grid_type = GridType.square;
            else if (CBtype.SelectedIndex == 2)
                grid_type = GridType.square_horizontal_precision;
            else if (CBtype.SelectedIndex == 3)
                grid_type = GridType.square_vertical_precision;
            else if (CBtype.SelectedIndex == 4)
                grid_type = GridType.square_combined_precision;

            //if grid size changed whole grid must be generated again
            if (desired_figures_nr != int.Parse(TBdesired_figures_nr.Text)
                || grid_type != prev_grid_type)
            {
                grid_size_changed = true;
            }

            grid_lines = CBlines.SelectedIndex;

            desired_figures_nr = int.Parse(TBdesired_figures_nr.Text);
            font_size = int.Parse(TBfont_size.Text);

            if (saving_enabled == false)
            {
                int argb = Convert.ToInt32(color_bg_str);

                byte[] values = BitConverter.GetBytes(argb);

                byte a = values[3];
                byte r = values[2];
                byte g = values[1];
                byte b = values[0];

                TBbackground_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));
                color_bg = Color.FromArgb(a, r, g, b);

                argb = Convert.ToInt32(color_font_str);

                values = BitConverter.GetBytes(argb);

                a = values[3];
                r = values[2];
                g = values[1];
                b = values[0];

                TBfont_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));
                color_font = Color.FromArgb(a, r, g, b);
            }
            else
            {
                SolidColorBrush scb = (SolidColorBrush)TBbackground_color.Background;
                System.Drawing.Color color =
                    System.Drawing.Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
                color_bg_str = color.ToArgb().ToString();
                color_bg = Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);

                scb = (SolidColorBrush)TBfont_color.Background;
                color =
                    System.Drawing.Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
                color_font_str = color.ToArgb().ToString();
                color_font = Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
            }

            bool prev_smart_grid = smart_grid; 
            smart_grid = (bool)CHBsmart_mousegrid.IsChecked;

            string prev_keyboard_layout = keyboard_layout;
            keyboard_layout = CBkeyboard_layout.SelectedItem.ToString();

            //need a .bat file to start an .exe file for some reasons
            Microsoft.Win32.RegistryKey rkApp = Microsoft.Win32.Registry.CurrentUser
                .OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            if (CHBrun_at_startup.IsChecked == true)
            {
                if (rkApp.GetValue(prog_name) == null)
                {
                    rkApp.SetValue(prog_name,
                        System.Reflection.Assembly.GetExecutingAssembly().Location.Replace(".exe", ".vbs"));
                }
                generate_bat_file();
            }
            else if (rkApp.GetValue(prog_name) != null)
            {
                rkApp.DeleteValue(prog_name, false);
            }

            auto_updates = (bool)CHBcheck_for_updates.IsChecked;

            if (CBss_voices.Items.Count > 0)
                ss_voice = CBss_voices.SelectedItem.ToString();
            if (ss_voice != "")
                ss.SelectVoice(ss_voice);

            ss_volume = int.Parse(TBss_volume.Text);
            ss.Volume = ss_volume;
            read_status = (bool)CHBread_status.IsChecked;

            if ((grid_size_changed || prev_smart_grid != smart_grid
                || prev_keyboard_layout != keyboard_layout)
                && THRmonitor != null) //we don't want this executed when settings are being loaded
            {
                thread_suspend1 = true;
                while (thread_suspended1 == false)
                {
                    Thread.Sleep(10); //wait for thread to finish its jobs
                }

                if (keyboard_layout == "US English / US International")
                    create_grid_alphabet_for_US_kl();
                else
                    create_grid_alphabet_for_any_kl();

                grid_symbols_limit = grid_alphabet.Count;
                max_figures_nr = (int)Math.Pow((double)grid_symbols_limit, 2);

                if (int.Parse(TBdesired_figures_nr.Text) > max_figures_nr)
                    TBdesired_figures_nr.Text = max_figures_nr.ToString();

                Create_Grid();

                if (smart_grid)
                {
                    load_grids();
                }

                thread_suspend1 = false;
            }
            //only needed if grid was already created //we don't want this executed when settings are being loaded
            else if (grid_width != 0)
            {
                if (MW != null)
                {
                    //need to close mousegrid window or there will be 1 more window every time you change
                    //mousegrid settings that require mousegrid regeneration
                    //you can easily see these windows by pressing  windows key + tab
                    MW.Close();
                }
                MW = new MouseGrid(grid_width, grid_height, grid_lines, grid_type, font_family,
                    font_size, color_bg, color_font, rows_nr, cols_nr, figure_width, figure_height,
                    grids[0].elements);
            }

            if ((bool)CHBread_status.IsChecked)
                read_status = true;
            else
                read_status = false;
        }

        void generate_bat_file()
        {
            FileStream fs = null;
            StreamWriter sw = null;
            string file_path = System.IO.Path.Combine(
                System.Reflection.Assembly.GetExecutingAssembly().Location.
                Replace(".exe", ".vbs"));

            try
            {
                fs = new FileStream(file_path, FileMode.Create, FileAccess.Write);
                sw = new StreamWriter(fs);

                //cmd script (black window appearing for a moment is a problem)
                //sw.WriteLine("cd \"" + app_folder_path + "\"");
                //sw.WriteLine("start " 
                //    + System.Diagnostics.Process.GetCurrentProcess().MainModule.ModuleName);

                //vbs script (no window appearing during execution)
                sw.WriteLine("Set objShell = CreateObject(\"Wscript.Shell\")");
                sw.WriteLine("objShell.CurrentDirectory = \"" + app_folder_path + "\"");
                sw.WriteLine("strApp = \"\"\""
                    + System.Diagnostics.Process.GetCurrentProcess().MainModule.ModuleName + "\"\"\"");
                sw.WriteLine("objShell.Run(strApp)");

                sw.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM007", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    if (sw != null)
                        sw.Close();
                    if (fs != null)
                        fs.Close();
                }
                catch (Exception ex2) { }
            }
        }

        void save_settings()
        {
            if (saving_enabled)
            {
                FileStream fs = null;
                StreamWriter sw = null;

                try
                {
                    if (CBlshift_action_a.SelectedIndex == -1)
                        throw new Exception("Left shift action in green mode was not selected.");
                    if (CBrshift_action_a.SelectedIndex == -1)
                        throw new Exception("Right shift action in green mode was not selected.");
                    if (CBlshift_action_b.SelectedIndex == -1)
                        throw new Exception("Left shift action in blue mode was not selected.");
                    if (CBrshift_action_b.SelectedIndex == -1)
                        throw new Exception("Right shift action in blue mode was not selected.");
                    if (CBtype.SelectedIndex == -1)
                        throw new Exception("Mousegrid type was not selected.");
                    if (CBlines.SelectedIndex == -1)
                        throw new Exception("Mousegrid lines were not selected.");

                    int trash;
                    if (int.TryParse(TBdesired_figures_nr.Text, out trash) == false
                        || int.Parse(TBdesired_figures_nr.Text) < 5
                        || int.Parse(TBdesired_figures_nr.Text) > max_figures_nr)
                        throw new Exception("Desired figures number for current keyboard layout must be between 5 and "
                            + max_figures_nr + ".");

                    if (int.TryParse(TBfont_size.Text, out trash) == false
                        || int.Parse(TBfont_size.Text) < 1
                        || int.Parse(TBfont_size.Text) > max_font_size)
                        throw new Exception("Mousegrid font size must be between 1 and "
                            + max_font_size + ".");

                    if (CBss_voices.SelectedIndex == -1 && (bool)CHBread_status.IsChecked)
                        throw new Exception("Speech synthesis voice must be selected when" +
                            " \"Read status\" is checked.");
                                        
                    if (int.TryParse(TBss_volume.Text, out trash) == false
                        || int.Parse(TBss_volume.Text) < 0
                        || int.Parse(TBss_volume.Text) > 100)
                        throw new Exception("Speech synthesis volume" +
                            " must be between 0 and 100");

                    string file_path = System.IO.Path.Combine(saving_folder_path, filename_settings);

                    if (Directory.Exists(saving_folder_path) == false)
                    {
                        Directory.CreateDirectory(saving_folder_path);
                    }

                    fs = new FileStream(file_path, FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs);

                    set_values();

                    sw.WriteLine(CBlshift_action_a.SelectedIndex);
                    sw.WriteLine(CBrshift_action_a.SelectedIndex);
                    sw.WriteLine(CBlshift_action_b.SelectedIndex);
                    sw.WriteLine(CBrshift_action_b.SelectedIndex);

                    sw.WriteLine(CBlalt_action_a.SelectedIndex);
                    sw.WriteLine(CBralt_action_a.SelectedIndex);
                    sw.WriteLine(CBlalt_action_b.SelectedIndex);
                    sw.WriteLine(CBralt_action_b.SelectedIndex);

                    sw.WriteLine(CBlctrl_action_a.SelectedIndex);
                    sw.WriteLine(CBrctrl_action_a.SelectedIndex);
                    sw.WriteLine(CBlctrl_action_b.SelectedIndex);
                    sw.WriteLine(CBrctrl_action_b.SelectedIndex);

                    sw.WriteLine(CBtype.SelectedIndex);
                    sw.WriteLine(CBlines.SelectedIndex);
                    sw.WriteLine(TBdesired_figures_nr.Text);
                    sw.WriteLine(color_bg_str);
                    sw.WriteLine(color_font_str);
                    sw.WriteLine(TBfont_size.Text);
                    sw.WriteLine(CHBsmart_mousegrid.IsChecked);

                    sw.WriteLine(CBkeyboard_layout.SelectedIndex);
                    sw.WriteLine(CHBcheck_for_updates.IsChecked);
                    sw.WriteLine(CHBrun_at_startup.IsChecked);
                    sw.WriteLine(CHBstart_minimized.IsChecked);
                    sw.WriteLine(CHBminimize_to_tray.IsChecked);

                    sw.WriteLine(CBss_voices.SelectedIndex);
                    sw.WriteLine(TBss_volume.Text);
                    sw.WriteLine(CHBread_status.IsChecked);

                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error AM008", MessageBoxButton.OK, MessageBoxImage.Error);

                    try
                    {
                        if (sw != null)
                            sw.Close();
                        if (fs != null)
                            fs.Close();
                    }
                    catch (Exception ex2) { }
                }
            }
        }

        void load_settings()
        {
            FileStream fs = null;
            StreamReader sr = null;
            string file_path = System.IO.Path.Combine(saving_folder_path, filename_settings);

            try
            {
                loading = true;

                if (File.Exists(file_path))
                {
                    fs = new FileStream(file_path, FileMode.Open, FileAccess.Read);
                    sr = new StreamReader(fs);

                    CBlshift_action_a.SelectedIndex = int.Parse(sr.ReadLine());
                    CBrshift_action_a.SelectedIndex = int.Parse(sr.ReadLine());
                    CBlshift_action_b.SelectedIndex = int.Parse(sr.ReadLine());
                    CBrshift_action_b.SelectedIndex = int.Parse(sr.ReadLine());

                    CBlalt_action_a.SelectedIndex = int.Parse(sr.ReadLine());
                    CBralt_action_a.SelectedIndex = int.Parse(sr.ReadLine());
                    CBlalt_action_b.SelectedIndex = int.Parse(sr.ReadLine());
                    CBralt_action_b.SelectedIndex = int.Parse(sr.ReadLine());

                    CBlctrl_action_a.SelectedIndex = int.Parse(sr.ReadLine());
                    CBrctrl_action_a.SelectedIndex = int.Parse(sr.ReadLine());
                    CBlctrl_action_b.SelectedIndex = int.Parse(sr.ReadLine());
                    CBrctrl_action_b.SelectedIndex = int.Parse(sr.ReadLine());

                    CBtype.SelectedIndex = int.Parse(sr.ReadLine());
                    CBlines.SelectedIndex = int.Parse(sr.ReadLine());
                    TBdesired_figures_nr.Text = sr.ReadLine();
                    color_bg_str = sr.ReadLine();
                    color_font_str = sr.ReadLine();
                    TBfont_size.Text = sr.ReadLine();
                    CHBsmart_mousegrid.IsChecked = smart_grid = bool.Parse(sr.ReadLine());

                    CBkeyboard_layout.SelectedIndex = int.Parse(sr.ReadLine());

                    if (CBkeyboard_layout.SelectedItem.ToString() == "US English / US International")
                        create_grid_alphabet_for_US_kl();
                    else
                        create_grid_alphabet_for_any_kl();

                    grid_symbols_limit = grid_alphabet.Count;
                    max_figures_nr = (int)Math.Pow((double)grid_symbols_limit, 2);

                    CHBcheck_for_updates.IsChecked = bool.Parse(sr.ReadLine());
                    CHBrun_at_startup.IsChecked = bool.Parse(sr.ReadLine());
                    CHBstart_minimized.IsChecked = bool.Parse(sr.ReadLine());
                    CHBminimize_to_tray.IsChecked = bool.Parse(sr.ReadLine());

                    CBss_voices.SelectedIndex = int.Parse(sr.ReadLine());
                    TBss_volume.Text = sr.ReadLine();
                    CHBread_status.IsChecked = bool.Parse(sr.ReadLine());

                    //Checkboxes Checked and Unchecked events work only after form is loaded
                    //so they have to be called manually in order to load save data properly
                    CHBcheck_for_updates_Checked(new object(), new RoutedEventArgs());
                    CHBsmart_mousegrid_Checked(new object(), new RoutedEventArgs());
                    CHBrun_at_startup_Checked(new object(), new RoutedEventArgs());
                    CHBstart_minimized_Checked(new object(), new RoutedEventArgs());
                    CHBminimize_to_tray_Checked(new object(), new RoutedEventArgs());

                    CHBread_status_Checked(new object(), new RoutedEventArgs());

                    set_values();

                    sr.Close();
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                loading_error = true;
                MessageBox.Show(ex.Message, "Error AM009", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    if (sr != null)
                        sr.Close();
                    if (fs != null)
                        fs.Close();
                }
                catch (Exception ex2) { }
            }
            finally
            {
                loading = false;
            }
        }

        void fix_wrong_loaded_values()
        {
            try
            {
                if (ss_volume > 100 || ss_volume < 0)
                {
                    ss_volume = 100;
                    TBss_volume.Text = ss_volume.ToString();
                }
                if (desired_figures_nr > max_figures_nr)
                {
                    desired_figures_nr = max_figures_nr;
                    TBdesired_figures_nr.Text = desired_figures_nr.ToString();
                }
                if (int.Parse(color_bg_str) < -16777216
                            || int.Parse(color_bg_str) > -1)
                {
                    color_bg_str = "-1973791";
                    color_bg = Color.FromRgb(225, 225, 225); //bg color
                }
                if (int.Parse(color_font_str) < -16777216
                    || int.Parse(color_font_str) > -1)
                {
                    color_font_str = "-16777216";
                    color_font = Color.FromRgb(0, 0, 0); //font color
                }
                if (font_size < 1 || font_size > max_font_size)
                {
                    font_size = 12;
                    TBfont_size.Text = font_size.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM010", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private class MyWebClient : WebClient
        {
            protected override WebRequest GetWebRequest(Uri uri)
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                WebRequest w = base.GetWebRequest(uri);
                w.Timeout = 3000;
                return w;
            }
        }
    }

    public static class StringExtensions
    {
        public static string FirstCharToUpper(this string input)
        {
            switch (input)
            {
                case null: throw new ArgumentNullException(nameof(input));
                case "": throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input));
                default: return input[0].ToString().ToUpper() + input.Substring(1);
            }
        }
    }
}