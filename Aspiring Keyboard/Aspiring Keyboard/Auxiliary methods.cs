//highest error nr: AM010
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

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
            CBlalt_action_b.SelectedItem = "Drag and drop";
            CBralt_action_b.SelectedItem = "Right click";
            CBlctrl_action_b.SelectedItem = "Double left click";
            CBrctrl_action_b.SelectedItem = "Move mouse";

            CBtype.SelectedItem = "Square combined precision";
            CBlines.SelectedItem = "None";
            TBdesired_figures_nr.Text = max_figures_nr.ToString();
            TBbackground_color.Text = "-1973791";
            TBfont_color.Text = "-16777216";
            TBfont_size.Text = "12";
            CHBsmart_mousegrid.IsChecked = true;

            CHBrun_at_startup.IsChecked = false;
            CHBstart_minimized.IsChecked = false;
            CHBminimize_to_tray.IsChecked = false;
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

        private void TBbackground_color_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                System.Windows.Forms.DialogResult dr = colorDialog1.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    TBbackground_color.Text = colorDialog1.Color.ToArgb().ToString();
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
                    TBfont_color.Text = colorDialog2.Color.ToArgb().ToString();
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
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("User Guide.pdf");
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Process.Start("Mousegrid Alphabet.pdf");
        }

        void set_values()
        {
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
            else if (CBlshift_action_a.SelectedIndex == 0)
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
            else if (CBrshift_action_a.SelectedIndex == 0)
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
            else if (CBlshift_action_b.SelectedIndex == 0)
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
            else if (CBrshift_action_b.SelectedIndex == 0)
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
            else if (CBlalt_action_a.SelectedIndex == 0)
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
            else if (CBralt_action_a.SelectedIndex == 0)
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
            else if (CBlalt_action_b.SelectedIndex == 0)
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
            else if (CBralt_action_b.SelectedIndex == 0)
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
            else if (CBlctrl_action_a.SelectedIndex == 0)
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
            else if (CBrctrl_action_a.SelectedIndex == 0)
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
            else if (CBlctrl_action_b.SelectedIndex == 0)
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
            else if (CBrctrl_action_b.SelectedIndex == 0)
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

            int argb = Convert.ToInt32(TBbackground_color.Text);

            byte[] values = BitConverter.GetBytes(argb);

            byte a = values[3];
            byte r = values[2];
            byte g = values[1];
            byte b = values[0];

            color1 = Color.FromArgb(a, r, g, b);

            argb = Convert.ToInt32(TBfont_color.Text);

            values = BitConverter.GetBytes(argb);

            a = values[3];
            r = values[2];
            g = values[1];
            b = values[0];

            color2 = Color.FromArgb(a, r, g, b);

            bool prev_smart_grid = smart_grid; 
            smart_grid = (bool)CHBsmart_mousegrid.IsChecked;

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

            if ((grid_size_changed || prev_smart_grid != smart_grid) && THRmonitor != null) //we don't want this executed when settings are being loaded
            {
                thread_suspend1 = true;
                while (thread_suspended1 == false)
                {
                    Thread.Sleep(10); //wait for thread to finish its jobs
                }

                Create_Grid();

                if (smart_grid && prev_smart_grid == false)
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
                    //you can easily see this windows by pressing  windows key + tab
                    MW.Close();
                }
                MW = new MouseGrid(grid_width, grid_height, grid_lines, grid_type, font_family,
                    font_size, color1, color2, rows_nr, cols_nr, figure_width, figure_height,
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

                    if (int.Parse(TBdesired_figures_nr.Text) < 5
                        || int.Parse(TBdesired_figures_nr.Text) > max_figures_nr)
                        throw new Exception("Desired figures number must be between 5 and "
                            + max_figures_nr + ".");

                    if (int.Parse(TBbackground_color.Text) < -16777216
                        || int.Parse(TBbackground_color.Text) > -1)
                        throw new Exception("Mousegrid size must be between -16777216 and -1.");

                    if (int.Parse(TBfont_color.Text) < -16777216
                        || int.Parse(TBfont_color.Text) > -1)
                        throw new Exception("Mousegrid size must be between -16777216 and -1.");

                    if (int.Parse(TBfont_size.Text) < 1)
                        throw new Exception("Mousegrid font size must be greater than 0.");

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
                    sw.WriteLine(TBbackground_color.Text);
                    sw.WriteLine(TBfont_color.Text);
                    sw.WriteLine(TBfont_size.Text);
                    sw.WriteLine(CHBsmart_mousegrid.IsChecked);

                    sw.WriteLine(CHBrun_at_startup.IsChecked);
                    sw.WriteLine(CHBstart_minimized.IsChecked);
                    sw.WriteLine(CHBminimize_to_tray.IsChecked);
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
                    TBbackground_color.Text = sr.ReadLine();
                    TBfont_color.Text = sr.ReadLine();
                    TBfont_size.Text = sr.ReadLine();
                    CHBsmart_mousegrid.IsChecked = smart_grid = bool.Parse(sr.ReadLine());

                    CHBrun_at_startup.IsChecked = bool.Parse(sr.ReadLine());
                    CHBstart_minimized.IsChecked = bool.Parse(sr.ReadLine());
                    CHBminimize_to_tray.IsChecked = bool.Parse(sr.ReadLine());
                    CHBread_status.IsChecked = bool.Parse(sr.ReadLine());

                    //Checkboxes Checked and Unchecked events work only after form is loaded
                    //so they have to be called manually in order to load save data properly
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
        }

        void fix_wrong_loaded_values()
        {
            try
            {
                if (desired_figures_nr > max_figures_nr)
                {
                    desired_figures_nr = max_figures_nr;
                    TBdesired_figures_nr.Text = desired_figures_nr.ToString();
                }
                if (int.Parse(TBbackground_color.Text) < -16777216
                            || int.Parse(TBbackground_color.Text) > -1)
                {
                    TBbackground_color.Text = "-1973791";
                    color1 = Color.FromRgb(225, 225, 225); //bg color
                }

                if (int.Parse(TBfont_color.Text) < -16777216
                    || int.Parse(TBfont_color.Text) > -1)
                {
                    TBfont_color.Text = "-16777216";
                    color2 = Color.FromRgb(0, 0, 0); //font color
                }
                if (font_size < 1)
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