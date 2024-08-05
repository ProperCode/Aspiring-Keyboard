//highest error nr: MW005
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Windows.Threading;
using WindowsInput;
using WindowsInput.Native;
using System.Speech.Synthesis;

//resolved difficult issue: something sometimes messed up mousegrid and it would not appear

namespace Aspiring_Keyboard
{
    public partial class MainWindow : Window
    {
        const bool check_if_already_running = true;

        const string prog_name = "Aspiring Keyboard";
        const string prog_version = "1.3";
        const string url_latest_version = "https://raw.githubusercontent.com/ProperCode/Aspiring-Keyboard/main/other/latest_version.txt";
        const string url_homepage = "github.com/ProperCode/Aspiring-Keyboard";
        string latest_version = "";
        const string copyright_text = "Copyright © 2024 Mikołaj Magowski. All rights reserved.";
        const string filename_settings = "settings_ak.txt";
        const string grids_foldername = "grids";
        const bool resized_grid = true;
        const bool movable_grid = true;

        const int default_desired_figures_nr = 2704;
        const string default_color_bg_str = "-1973791"; //light grey
        const string default_color_font_str = "-16777216"; //black
        string default_ss_voice = "";
        const int default_ss_volume = 100;

        bool enabled = true;

        const string icon_red = "pack://application:,,,/Aspiring Keyboard;component/images/ak red.ico";
        const string icon_green = "pack://application:,,,/Aspiring Keyboard;component/images/ak green.ico";
        const string icon_blue = "pack://application:,,,/Aspiring Keyboard;component/images/ak blue.ico";

        const string folder_name_grid_hexagonal = "hexagonal";
        const string folder_name_grid_square = "square";
        const string folder_name_grid_horizontal = "square_horizontal";
        const string folder_name_grid_vertical = "square_vertical";
        const string folder_name_grid_combined = "square_combined";

        string users_directory_path
            = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        //full path is necessary if run at startup is used (running at startup uses different current
        //directory
        string app_folder_path = System.IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location);
        string saving_folder_path;

        bool saving_enabled = false;
        bool loading_error = false;

        Process prc;

        InputSimulator sim = new InputSimulator();

        Thread THRmonitor;
        //bool thread_abort1 = false; //redeclaring and starting stopped thread doesn't work (weird)
        bool thread_suspend1 = false;
        bool thread_suspended1 = false;

        List<Grid_Symbol> grid_alphabet = new List<Grid_Symbol>();

        int grid_symbols_limit = 1;//50-58 is recommended
        int max_figures_nr = 2704;
        int desired_figures_nr = 2704;
        GridType grid_type = GridType.hexagonal;
        bool smart_grid = true;
        double offset = 0.25;
        double np_offset = 0.25; //offset by percentage of figure height (normal precision)
        double hp_offset = 0.025; //high precision offset
        int offset_x = 0; //grid offset
        int offset_y = 0;
        int grid_lines = 0; //0-2 (0 - no lines, 1 - dotted lines, 2 - normal lines)

        Color color_font;// = Color.FromRgb(0, 0, 0); //font color
        //Color color_bg = Color.FromRgb(255, 255, 255);
        Color color_bg;// = Color.FromRgb(225, 225, 225); //bg color

        string color_bg_str; //light grey
        string color_font_str; //black        

        const int max_font_size = 400;
        int font_size = 12;
        //bool auto_grid_font_size = true; //bad idea
        FontFamily font_family = new FontFamily("Verdana");
        //FontFamily font_family = new FontFamily("Tahoma");
        //FontFamily font_family = new FontFamily("Microsoft Sans Serif");
        //FontFamily font_family = new FontFamily("Calibri");//12,10
        //FontFamily font_family = new FontFamily("Arial");
        //FontFamily font_family = new FontFamily("Times New Roman");
        //FontFamily font_family = new FontFamily("Courier New");

        int max_figures;

        int screen_width, screen_height;

        List<int> rows = new List<int>();
        List<int> cols = new List<int>();

        int rows_nr, cols_nr;

        int figures;
        double figure_width, figure_height;
        double grid_width, grid_height;

        MouseGrid MW;
        bool grid_visible = false;
        ActionX last_action = ActionX.none;

        List<Process_grid> grids = new List<Process_grid>();
        int grid_ind = 0;
        int prev_installed_apps_count = -1;

        string[] keyboard_layouts = { "Any", "US English / US International" };
        string keyboard_layout;

        //Used by speech synthesis:
        List<string> ss_voices_priority_list = new List<string>() { "Zira", "Susan", "Hazel", "Linda",
            "Catherine", "Heera", "Sean" };
        //Zira - US, Susan/Hazel - UK, Linda - Canada, Catherine - Australia, Heera - India, Sean - Ireland
        IReadOnlyCollection<InstalledVoice> installed_voices;

        string ss_voice;
        int ss_volume;
        bool read_status = true;
        bool auto_updates = false;

        SpeechSynthesizer ss = new SpeechSynthesizer();

        bool green_mode = true;        

        ActionX left_shift_action_a = ActionX.none;
        ActionX right_shift_action_a = ActionX.none;
        ActionX left_alt_action_a = ActionX.none;
        ActionX right_alt_action_a = ActionX.none;
        ActionX left_ctrl_action_a = ActionX.none;
        ActionX right_ctrl_action_a = ActionX.none;

        ActionX left_shift_action_b = ActionX.none;
        ActionX right_shift_action_b = ActionX.none;
        ActionX left_alt_action_b = ActionX.none;
        ActionX right_alt_action_b = ActionX.none;
        ActionX left_ctrl_action_b = ActionX.none;
        ActionX right_ctrl_action_b = ActionX.none;

        ActionX action = ActionX.none;

        bool repeat_action_indefinitely = false;

        System.Windows.Forms.MenuItem mi_switch_mode;
        System.Windows.Forms.MenuItem mi_toggle;
        System.Windows.Forms.MenuItem mi_exit;

        public MainWindow()
        {
            try
            {
                saving_enabled = false;
                //saving_folder_path = users_directory_path + "\\" + prog_name;
                saving_folder_path = app_folder_path;

                prc = Process.GetCurrentProcess();
                prc.PriorityClass = ProcessPriorityClass.High;

                Stream iconStream = System.Windows.Application.GetResourceStream(
                    new Uri(icon_green)).Stream;

                // Create a BitmapSource  
                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(icon_green);
                bitmap.EndInit();

                ni.Icon = new System.Drawing.Icon(iconStream);
                ni.MouseClick += new System.Windows.Forms.MouseEventHandler(ni_MouseClick);
                ni.Visible = true;

                iconStream.Close();

                System.Windows.Forms.ContextMenu cm = new System.Windows.Forms.ContextMenu();
                mi_switch_mode = new System.Windows.Forms.MenuItem();
                mi_toggle = new System.Windows.Forms.MenuItem();
                mi_exit = new System.Windows.Forms.MenuItem();

                // Initialize contextMenu1
                cm.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { mi_switch_mode });
                cm.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { mi_toggle });
                cm.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { mi_exit });

                mi_switch_mode.Index = 0;
                mi_switch_mode.Text = "Switch mode";
                mi_switch_mode.Click += new System.EventHandler(mi_switch_mode_Click);

                mi_toggle.Index = 1;
                mi_toggle.Text = "Toggle";
                mi_toggle.Click += new System.EventHandler(mi_toggle_Click);

                mi_exit.Index = 2;
                mi_exit.Text = "Exit";
                mi_exit.Click += new System.EventHandler(mi_exit_Click);

                ni.ContextMenu = cm;

                if (check_if_already_running)
                {
                    is_program_already_running();
                }

                InitializeComponent();

                this.Title = prog_name;
                //Lprogram_name.Content = prog_name;
                //Linstalled_version.Content = "Installed version: " + prog_version;
                //Lhomepage.Content = "Homepage: " + Middle_Man.url_homepage;
                //Lcopyright.Content = copyright_text;

                TBcontrol_keys.Text = "Pause Break - turn on/off"
                    + "\r\nScroll Lock - change mode"
                    + "\r\nCaps Lock - left click without losing focus"
                    + "\r\nInsert - repeat last action"
                    + "\r\nLAlt + RAlt - release left mouse button"
                    + "\r\nLShift + RShift - browser forward button"
                    + "\r\nNum Lock - browser back button"
                    + "\r\nLShift + Caps Lock - left shift action in infinity mode"
                    + "\r\nRShift + Caps Lock - right shift action in infinity mode"
                    + "\r\nCaps Lock + Up/Down arrow - scroll up/down"
                    + "\r\nWhen mousegrid is visible:"
                    + "\r\n- Arrow keys - move mousegrid"
                    + "\r\n- Caps Lock  - change mousegrid movement mode"
                    + "\r\n- Escape - cancel action and hide mousegrid";

                installed_voices = ss.GetInstalledVoices();

                foreach (string kl in keyboard_layouts)
                {
                    CBkeyboard_layout.Items.Add(kl);
                }

                foreach (InstalledVoice iv in installed_voices)
                {
                    CBss_voices.Items.Add(iv.VoiceInfo.Name);
                }

                foreach (GridType type in (GridType[])Enum.GetValues(typeof(GridType)))
                {
                    CBtype.Items.Add(type.ToString().Replace("_", " ").FirstCharToUpper());
                }

                foreach (ActionX action in (ActionX[])Enum.GetValues(typeof(ActionX)))
                {
                    CBlshift_action_a.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBrshift_action_a.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlalt_action_a.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBralt_action_a.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlctrl_action_a.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBrctrl_action_a.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlshift_action_b.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBrshift_action_b.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlalt_action_b.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBralt_action_b.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlctrl_action_b.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                    CBrctrl_action_b.Items.Add(action.ToString().Replace("_", " ").FirstCharToUpper());
                }

                CBlines.Items.Add("None");
                CBlines.Items.Add("Dotted");
                CBlines.Items.Add("Solid");                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW001", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            try
            {
                restore_default_settings();

                set_values();

                load_settings();

                fix_wrong_loaded_values();

                saving_enabled = true;

                if (loading_error)
                {
                    MessageBox.Show("Loading error was detected. All settings" +
                        "will be restored to default and saved.", "Warning",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    restore_default_settings();
                    save_settings(); //save settings so loading error won't happen again (default values
                                     //will take place of unread values)
                }

                CenterWindowOnScreen();

                if ((bool)CHBstart_minimized.IsChecked == true)
                {
                    WindowState = WindowState.Minimized;

                    if (CHBminimize_to_tray.IsChecked == true)
                    {
                        this.Hide();
                    }
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
                    load_grids();

                green_mode = true;
                
                THRmonitor = new Thread(new ThreadStart(monitor));
                THRmonitor.Start();

                if (read_status) ss.SpeakAsync("Green mode enabled");

                CBlshift_action_a.Focus();

                if(auto_updates)
                    update_app_if_necessary();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void monitor()
        {
            string icon;
            int loop_nr = 0;

            while (true)
            {
                try
                {
                    if (IsKeyPushedDown(System.Windows.Forms.Keys.Pause))
                    {
                        fix_very_rare_problem();

                        while (IsKeyPushedDown(System.Windows.Forms.Keys.Pause))
                        {
                            Thread.Sleep(10);
                        }

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
                    else if (enabled)
                    {
                        if (IsKeyPushedDown(System.Windows.Forms.Keys.Scroll))
                        {
                            fix_very_rare_problem();

                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Scroll))
                            {
                                Thread.Sleep(10);
                            }

                            sim.Keyboard.KeyPress(VirtualKeyCode.SCROLL);

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
                        else if (IsKeyPushedDown(System.Windows.Forms.Keys.NumLock))
                        {
                            fix_very_rare_problem();

                            while (IsKeyPushedDown(System.Windows.Forms.Keys.NumLock))
                            {
                                Thread.Sleep(10);
                            }

                            sim.Keyboard.KeyPress(VirtualKeyCode.NUMLOCK);

                            sim.Keyboard.KeyPress(VirtualKeyCode.BROWSER_BACK);
                            //sim.Mouse.VerticalScroll(-40);

                            if (read_status) ss.SpeakAsync("Back"); 
                        }
                        else if (last_action != ActionX.none &&
                            IsKeyPushedDown(System.Windows.Forms.Keys.Insert))
                        {
                            fix_very_rare_problem();

                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Insert))
                            {
                                Thread.Sleep(10);
                            }

                            sim.Keyboard.KeyPress(VirtualKeyCode.INSERT);

                            adv_mouse(true);
                        }
                        else if (grid_visible == false)
                        {
                            bool dont_lose_focus = false;
                            bool cancel = false;

                            if (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey))
                            {
                                fix_very_rare_problem();

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey))
                                {
                                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                        typeof(VirtualKeyCode)))
                                    {
                                        if (vkc != VirtualKeyCode.LSHIFT
                                            && vkc != VirtualKeyCode.SHIFT
                                            && vkc != VirtualKeyCode.CAPITAL
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            cancel = true;
                                            //MessageBox.Show(vkc.ToString());
                                            break;
                                        }
                                    }

                                    if (IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey))
                                    {
                                        //close_that();

                                        //if (read_status) ss.SpeakAsync("Closing.");

                                        sim.Keyboard.KeyPress(VirtualKeyCode.BROWSER_FORWARD);

                                        if (read_status) ss.SpeakAsync("Forward");

                                        while (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey)
                                            || IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey))
                                        {
                                            Thread.Sleep(10);
                                        }
                                    }
                                    else if (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                    {
                                        repeat_action_indefinitely = true;
                                        break;
                                    }
                                }

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                {
                                    Thread.Sleep(10);
                                }

                                if (repeat_action_indefinitely)
                                {
                                    Thread.Sleep(10);

                                    sim.Keyboard.KeyPress(VirtualKeyCode.CAPITAL);
                                }

                                Thread.Sleep(10);
                                
                                if (cancel == false)
                                {
                                    if(green_mode)
                                        action = left_shift_action_a;
                                    else
                                        action = left_shift_action_b;
                                }
                            }
                            else if (IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey))
                            {
                                fix_very_rare_problem();

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey))
                                {
                                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                        typeof(VirtualKeyCode)))
                                    {
                                        if (vkc != VirtualKeyCode.RSHIFT
                                            && vkc != VirtualKeyCode.SHIFT
                                            && vkc != VirtualKeyCode.CAPITAL
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            cancel = true;
                                            //MessageBox.Show(vkc.ToString());
                                        }
                                    }

                                    if (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey))
                                    {
                                        //close_that();

                                        //if (read_status) ss.SpeakAsync("Closing.");

                                        sim.Keyboard.KeyPress(VirtualKeyCode.BROWSER_FORWARD);

                                        if (read_status) ss.SpeakAsync("Forward");

                                        while (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey)
                                            || IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey))
                                        {
                                            Thread.Sleep(10);
                                        }
                                    }
                                    else if (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                    {
                                        repeat_action_indefinitely = true;
                                        break;
                                    }

                                    Thread.Sleep(10);
                                }

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                {
                                    Thread.Sleep(10);
                                }

                                if (repeat_action_indefinitely)
                                {
                                    Thread.Sleep(10);

                                    sim.Keyboard.KeyPress(VirtualKeyCode.CAPITAL);
                                }

                                if (cancel == false)
                                {
                                    if (green_mode)
                                        action = right_shift_action_a;
                                    else
                                        action = right_shift_action_b;
                                }
                            }
                            else if (IsKeyPushedDown(System.Windows.Forms.Keys.LMenu))
                            {
                                fix_very_rare_problem();

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LMenu))
                                {
                                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                        typeof(VirtualKeyCode)))
                                    {
                                        if (vkc != VirtualKeyCode.LMENU
                                            && vkc != VirtualKeyCode.MENU
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            cancel = true;
                                            //MessageBox.Show(vkc.ToString());
                                            break;
                                        }
                                    }

                                    if (IsKeyPushedDown(System.Windows.Forms.Keys.RMenu))
                                    {
                                        release_LMB();

                                        if (read_status) ss.SpeakAsync("Left released.");

                                        while (IsKeyPushedDown(System.Windows.Forms.Keys.LMenu)
                                            || IsKeyPushedDown(System.Windows.Forms.Keys.RMenu))
                                        {
                                            Thread.Sleep(10);
                                        }

                                        break;
                                    }

                                    Thread.Sleep(10);
                                }

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LMenu)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.RMenu))
                                {
                                    Thread.Sleep(10);
                                }

                                if (cancel == false)
                                {
                                    if (green_mode)
                                        action = left_alt_action_a;
                                    else
                                        action = left_alt_action_b;
                                }
                            }
                            else if (IsKeyPushedDown(System.Windows.Forms.Keys.RMenu))
                            {
                                fix_very_rare_problem();

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.RMenu))
                                {
                                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                        typeof(VirtualKeyCode)))
                                    {
                                        if (vkc != VirtualKeyCode.RMENU
                                            && vkc != VirtualKeyCode.MENU
                                            && vkc != VirtualKeyCode.CONTROL //right alt is seen as control
                                            && vkc != VirtualKeyCode.LCONTROL //and Lcontrol
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            cancel = true;
                                            //MessageBox.Show(vkc.ToString());
                                            break;
                                        }
                                    }

                                    if (IsKeyPushedDown(System.Windows.Forms.Keys.LMenu))
                                    {
                                        release_LMB();

                                        if (read_status) ss.SpeakAsync("Left released.");

                                        while (IsKeyPushedDown(System.Windows.Forms.Keys.LMenu)
                                            || IsKeyPushedDown(System.Windows.Forms.Keys.RMenu))
                                        {
                                            Thread.Sleep(10);
                                        }

                                        break;
                                    }

                                    Thread.Sleep(10);
                                }

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LMenu)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.RMenu))
                                {
                                    Thread.Sleep(10);
                                }

                                if (cancel == false)
                                {
                                    if (green_mode)
                                        action = right_alt_action_a;
                                    else
                                        action = right_alt_action_b;
                                }
                            }
                            else if (IsKeyPushedDown(System.Windows.Forms.Keys.LControlKey))
                            {
                                fix_very_rare_problem();

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LControlKey))
                                {
                                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                        typeof(VirtualKeyCode)))
                                    {
                                        if (vkc != VirtualKeyCode.LCONTROL
                                            && vkc != VirtualKeyCode.CONTROL
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            cancel = true;
                                            //MessageBox.Show(vkc.ToString());
                                            break;
                                        }
                                    }
                                    Thread.Sleep(10);
                                }

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LControlKey))
                                {
                                    Thread.Sleep(10);
                                }

                                if (cancel == false)
                                {
                                    if (green_mode)
                                        action = left_ctrl_action_a;
                                    else
                                        action = left_ctrl_action_b;
                                }
                            }
                            else if (IsKeyPushedDown(System.Windows.Forms.Keys.RControlKey))
                            {
                                fix_very_rare_problem();

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.RControlKey))
                                {
                                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                        typeof(VirtualKeyCode)))
                                    {
                                        if (vkc != VirtualKeyCode.RCONTROL
                                            && vkc != VirtualKeyCode.CONTROL
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            cancel = true;
                                            //MessageBox.Show(vkc.ToString());
                                            break;
                                        }
                                    }
                                    Thread.Sleep(10);
                                }

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.RControlKey))
                                {
                                    Thread.Sleep(10);
                                }

                                if (cancel == false)
                                {
                                    if (green_mode)
                                        action = right_ctrl_action_a;
                                    else
                                        action = right_ctrl_action_b;
                                }
                            }
                            else if (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                            {
                                fix_very_rare_problem();

                                action = ActionX.left_click;

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                {
                                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                        typeof(VirtualKeyCode)))
                                    {
                                        if (vkc == VirtualKeyCode.UP
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            sim.Mouse.VerticalScroll(4);
                                            action = ActionX.none;
                                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Up))
                                                Thread.Sleep(10);
                                        }
                                        else if (vkc == VirtualKeyCode.DOWN
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            sim.Mouse.VerticalScroll(-4);
                                            action = ActionX.none;
                                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Down))
                                                Thread.Sleep(10);
                                        }
                                        /* PROBLEMATIC
                                        else if (vkc == VirtualKeyCode.LEFT
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            sim.Mouse.HorizontalScroll(-300);
                                            action = ActionX.scroll_left;
                                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Left))
                                                Thread.Sleep(10);
                                        }
                                        else if (vkc == VirtualKeyCode.RIGHT
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            sim.Mouse.HorizontalScroll(300);
                                            action = ActionX.scroll_right;
                                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Right))
                                                Thread.Sleep(10);
                                        }
                                        */
                                    }

                                    Thread.Sleep(10);
                                }

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.Up)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.Down))
                                {
                                    Thread.Sleep(10);
                                }

                                if (action == ActionX.left_click)
                                {
                                    dont_lose_focus = true;
                                }
                                
                                Thread.Sleep(10);

                                sim.Keyboard.KeyPress(VirtualKeyCode.CAPITAL);

                                if (read_status)
                                {
                                    if (action == ActionX.left_click)
                                    {
                                        ss.SpeakAsync("focus preserved");
                                    }
                                }

                                if (action != ActionX.left_click)
                                    action = ActionX.none;
                            }

                            do
                            {
                                if (action == ActionX.center_left_click)
                                {
                                    if (read_status) ss.SpeakAsync("center");

                                    LMBClick((int)(screen_width / 2), (int)(screen_height / 2));

                                    action = ActionX.none;
                                }
                                else if (action == ActionX.release_left)
                                {
                                    if (read_status) ss.SpeakAsync("left released");

                                    left_up();

                                    action = ActionX.none;
                                }
                                else if (action == ActionX.release_right)
                                {
                                    if (read_status) ss.SpeakAsync("right released");

                                    right_up();

                                    action = ActionX.none;
                                }
                                else if (action != ActionX.none && cancel == false)
                                {
                                    adv_mouse(false, dont_lose_focus);

                                    //sometimes RMB may stay pressed and mess things up
                                    release_buttons_and_keys();
                                }
                            }
                            while (repeat_action_indefinitely);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error MW003", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                thread_suspended1 = true;

                while (thread_suspend1)
                {
                    Thread.Sleep(10);
                }

                thread_suspended1 = false;

                Thread.Sleep(40);
                loop_nr++;
            }
        }

        void fix_very_rare_problem()
        {
            //this solves very rare mouse button and keys releasing problem (very weird issue)
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON))
            {
                left_up();
            }
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON))
            {
                right_up();
            }
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.BROWSER_BACK))
            {
                sim.Keyboard.KeyUp(VirtualKeyCode.BROWSER_BACK);
            }
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.BROWSER_FORWARD))
            {
                sim.Keyboard.KeyUp(VirtualKeyCode.BROWSER_FORWARD);
            }
        }

        void is_program_already_running()
        {
            Process[] arr = Process.GetProcesses();
            string[] a;
            int i = 0;

            foreach (Process p in arr)
            {
                if (p.ProcessName == prog_name)
                {
                    i++;
                }
            }

            if (i > 1)
            {
                MessageBox.Show(prog_name + " is already running.",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Process.GetCurrentProcess().Kill();
            }
        }

        void Create_Grid(bool reset_smart_grid = true)
        {
            //Mathematically incorrect, but returns desired value (thanks to Floor):
            int count = (int)Math.Floor(Math.Sqrt((double)desired_figures_nr));

            string s;
            VirtualKeyCode v;

            if (count > grid_symbols_limit)
                count = grid_symbols_limit;

            max_figures = count * count;

            //returns scaled width and height if windows scaling is enabled in system settings
            screen_width = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenWidth);
            screen_height = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenHeight);

            //returns real width and height if windows scaling is enabled in system settings
            //screen_width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
            //screen_height = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;

            rows = new List<int>();
            cols = new List<int>();

            int a, b;

            //i - square/hexagon width
            for (int i = 1; i < 2000; i++)
            {
                a = (int)Math.Round((double)(screen_width / i));
                b = (int)Math.Round((double)(screen_height / i));

                if (grid_type == GridType.hexagonal)
                {
                    if ((a * b + (a - 1) * (b - 1)) <= max_figures)
                    {
                        cols.Add(a);
                        rows.Add(b);
                    }
                }
                else
                {
                    if (a * b <= max_figures)
                    {
                        cols.Add(a);
                        rows.Add(b);
                    }
                }
            }

            rows_nr = rows[0];
            cols_nr = cols[0];

            if (grid_type == GridType.hexagonal)
                figures = cols_nr * rows_nr + (cols_nr - 1) * (rows_nr - 1);
            else
                figures = cols_nr * rows_nr;

            figure_width = screen_width / (double)cols_nr;
            figure_height = screen_height / (double)rows_nr;

            if (resized_grid == true)
            {
                grid_width = screen_width;
                grid_height = screen_height;
            }
            else
            {
                grid_width = figure_width * (double)cols_nr;
                grid_height = figure_height * (double)rows_nr;
            }

            //bad idea:
            //if (auto_grid_font_size)
            //{
            //    if (grid_type == GridType.hexagonal)
            //        font_size = (int)(Math.Floor(figure_width * 14 / 35.555555555555557));
            //    else
            //        font_size = (int)(Math.Floor(figure_width * 14 / 25.098039215686274));
            //}

            if (reset_smart_grid)
            {
                grids = new List<Process_grid>();

                grids.Add(new Process_grid("Default Process_grid ind=0"));

                int ind = 0;
                int count2;

                count2 = (int)Math.Ceiling(Math.Sqrt(figures));
                if (count2 < count)
                    count = count2;

                for (int i = 0; i < count && ind < figures; i++)
                {
                    for (int j = 0; j < count && ind < figures; j++)
                    {
                        grids[0].elements.Add(
                            new Grid_element(grid_alphabet[i].symbol + grid_alphabet[j].symbol));

                        ind++;
                    }
                }

                grids[0].count = grids[0].elements.Count;
                grid_ind = 0;
            }

            if (MW != null)
            {
                //need to close mousegrid window or there will be 1 more window every time you change
                //mousegrid settings that require mousegrid regeneration
                //you can easily see this windows by pressing  windows key + tab
                MW.Close();
            }
            MW = new MouseGrid(grid_width, grid_height, grid_lines, grid_type, font_family, font_size, 
                color_bg, color_font, rows_nr, cols_nr, figure_width, figure_height, grids[0].elements);
            //MW.regenerate_grid_symbols();
        }

        int drag_x = 0, drag_y = 0;
        int last_x = 0, last_y = 0;
        int execution_nr = 0;

        private void Bhomepage_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://github.com/ProperCode/Aspiring-Keyboard");
        }

        void adv_mouse(bool repeat_last_action = false, bool dont_lose_focus = false)
        {
            if (repeat_last_action == false)
            {
                IntPtr handle2 = GetForegroundWindow();
                last_action = action;

                int d = 1;
                if (action == ActionX.drag_and_drop)
                    d = 0;

                for (; d < 2; d++)
                {
                    if (smart_grid)
                    {
                        IntPtr handle = GetForegroundWindow();
                        string process_name = "";
                        grid_ind = -1;

                        Process[] arr = Process.GetProcesses();

                        foreach (Process p in arr)
                        {
                            if (p.MainWindowHandle == handle)
                            {
                                process_name = p.ProcessName;
                                break;
                            }
                        }

                        for (int k = 1; k < grids.Count && grid_ind == -1; k++)
                        {
                            if (grids[k].process_name == process_name)
                                grid_ind = k;
                        }

                        if (grid_ind == -1)
                        {
                            grids.Add(new Process_grid(process_name));
                            grid_ind = grids.Count - 1;

                            for (int k = 0; k < grids[0].elements.Count; k++)
                            {
                                grids[grid_ind].elements.Add(new Grid_element(
                                    grids[0].elements[k].symbols));
                            }

                            grids[grid_ind].count = grids[grid_ind].elements.Count;
                        }

                        //start_time();
                        MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                        {
                            MW.elements = grids[grid_ind].elements;
                            MW.regenerate_grid_symbols();//1-211 ms
                        }));
                        //stop_time();
                    }
                    else grid_ind = 0;

                    if (movable_grid)
                    {
                        offset_x = offset_y = 0;

                        MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                        {
                            MW.Top = offset_y;
                            MW.Left = offset_x;
                        }));
                    }

                    //start_time();
                    if (smart_grid)
                    {
                        MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                        {
                            MW.Topmost = true;
                            MW.Hide();//for Topmost
                            MW.Opacity = 1;
                            
                            MW.Show();

                            if (dont_lose_focus == false)
                            {
                                MW.Activate(); //important (gives focus to window)
                            }
                        }));
                    }
                    else
                    {
                        MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                        {
                            MW.Show();

                            if (dont_lose_focus == false)
                            {
                                MW.Activate(); //important (gives focus to window)
                            }
                        }));
                    }

                    //stop_time();
                    grid_visible = true;

                    if (execution_nr == 0)
                    {
                        string status = "";

                        if (repeat_action_indefinitely)
                            status = "indefinite ";

                        if (action == ActionX.drag_and_drop && d == 0)
                        {
                            status += "drag";
                        }
                        else if (action == ActionX.double_left_click)
                        {
                            status += "double";
                        }
                        else if (action == ActionX.triple_left_click)
                        {
                            status += "triple";
                        }
                        else if (action == ActionX.hold_left)
                        {
                            status += "hold left";
                        }
                        else if (action == ActionX.hold_right)
                        {
                            status += "hold right";
                        }
                        else if (action == ActionX.move_mouse)
                        {
                            status += "move mouse";
                        }
                        else if (action == ActionX.right_click)
                        {
                            status += "right click";
                        }

                        if (read_status) ss.SpeakAsync(status);
                    }

                    Thread.Sleep(50);

                    //This 2nd activate is sometimes needed:
                    if (dont_lose_focus == false)
                        MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                    {
                        MW.Activate(); //important (gives focus to window)
                    }));

                    string figure_str = "";
                    bool shift_pressed = false;
                    bool cancel = false;
                    offset = np_offset;

                    while (figure_str.Length < 2 && cancel == false)
                    {
                        if (IsKeyPushedDown(System.Windows.Forms.Keys.Escape))
                        {
                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Escape))
                            {
                                Thread.Sleep(10);
                            }

                            cancel = true;
                            repeat_action_indefinitely = false;
                            d = 2;
                        }
                        else if (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                        {
                            while (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                            {
                                Thread.Sleep(10);
                            }

                            Thread.Sleep(10);

                            sim.Keyboard.KeyPress(VirtualKeyCode.CAPITAL);

                            if (offset == np_offset) //if normal precision
                            {
                                offset = hp_offset;

                                if (read_status) ss.SpeakAsync("High precision");
                            }
                            else //if high precision
                            {
                                offset = np_offset;

                                if (read_status) ss.SpeakAsync("Normal precision");
                            }
                        }
                        else if (IsKeyPushedDown(System.Windows.Forms.Keys.Up))
                        {
                            offset_y -= (int)Math.Round(offset * figure_height);

                            MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                MW.Top = offset_y;
                                MW.Left = offset_x;
                            }));
                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Up))
                            {
                                Thread.Sleep(10);
                            }
                        }
                        else if (IsKeyPushedDown(System.Windows.Forms.Keys.Down))
                        {
                            offset_y += (int)Math.Round(offset * figure_height);

                            MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                MW.Top = offset_y;
                                MW.Left = offset_x;
                            }));
                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Down))
                            {
                                Thread.Sleep(10);
                            }
                        }
                        else if (IsKeyPushedDown(System.Windows.Forms.Keys.Left))
                        {
                            offset_x -= (int)Math.Round(offset * figure_width);

                            MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                MW.Top = offset_y;
                                MW.Left = offset_x;
                            }));
                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Left))
                            {
                                Thread.Sleep(10);
                            }
                        }
                        else if (IsKeyPushedDown(System.Windows.Forms.Keys.Right))
                        {
                            offset_x += (int)Math.Round(offset * figure_width);

                            MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                MW.Top = offset_y;
                                MW.Left = offset_x;
                            }));
                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Right))
                            {
                                Thread.Sleep(10);
                            }
                        }
                        else
                        {
                            if (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey)
                                || IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey))
                                shift_pressed = true;
                            else
                                shift_pressed = false;

                            //if (IsKeyPushedDown(System.Windows.Forms.Keys.Divide))
                            //{
                            //    //figure_str += "/";
                            //    //while (IsKeyPushedDown(System.Windows.Forms.Keys.Divide))
                            //    //{
                            //    //    Thread.Sleep(10);
                            //    //}
                            //}
                            //else
                            {
                                foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                    typeof(VirtualKeyCode)))
                                {
                                    if (sim.InputDeviceState.IsKeyDown(vkc)
                                        && vkc != VirtualKeyCode.LSHIFT
                                        && vkc != VirtualKeyCode.RSHIFT)
                                    {
                                        foreach (Grid_Symbol gs in grid_alphabet)
                                        {
                                            if (vkc == gs.vkc && ((gs.shift && shift_pressed)
                                                || (gs.shift == false && shift_pressed == false)))
                                            {
                                                //MessageBox.Show(gs.vkc.ToString());
                                                //Debug only section start----------------------
                                                //MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                                //{
                                                //    if (smart_grid)
                                                //        MW.Opacity = 0;
                                                //    else
                                                //        MW.Hide();
                                                //}));
                                                //Thread.Sleep(200);
                                                //Debug only section end------------------------
                                                figure_str += gs.symbol;
                                                while (sim.InputDeviceState.IsKeyDown(vkc))
                                                {
                                                    Thread.Sleep(10);
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        Thread.Sleep(40);
                    }

                    while (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey)
                        || IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey))
                    {
                        Thread.Sleep(10);
                    }

                    MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                    {
                        if (smart_grid)
                        {
                            //Mousegrid must be also hidden, otherwise using ALT + F4 could close it
                            //disabling topmost should be enough to solve this issue (if not use ShoveToBackground
                            //as well
                            MW.Topmost = false;

                            MW.Opacity = 0;
                            //MW.Hide(); //don't hide it or old mousegrid is going to be shown first (old smart data)
                        }
                        else
                            MW.Hide();
                    }));
                    grid_visible = false;

                    Thread.Sleep(50); //needed because MW.Activate() is used

                    SetForegroundWindow(handle2);

                    Thread.Sleep(50); //needed because of setforeground

                    bool f = false;
                    int x, y;

                    for (int i = 0; i < grids[grid_ind].elements.Count && f == false && cancel == false; i++)
                    {
                        if (figure_str == grids[grid_ind].elements[i].symbols)
                        {
                            f = true;

                            if (smart_grid)
                            {
                                int ind = -1;

                                int score = get_symbols_score(grids[grid_ind].elements[i].symbols[0],
                                    grids[grid_ind].elements[i].symbols[1]);

                                ind = find_index_with_lower_count(grid_ind,
                                    grids[grid_ind].elements[i].count, score);

                                if (ind != -1)
                                {
                                    string symbols = grids[grid_ind].elements[ind].symbols;
                                    uint count = grids[grid_ind].elements[ind].count;

                                    grids[grid_ind].elements[ind].count = grids[grid_ind].elements[i].count;
                                    grids[grid_ind].elements[ind].symbols = grids[grid_ind].elements[i].symbols;

                                    grids[grid_ind].elements[i].count = count;
                                    grids[grid_ind].elements[i].symbols = symbols;
                                }

                                grids[grid_ind].elements[i].count++;
                            }

                            //delayed Hide, because Mousegrid must be shown when figures content is updated
                            //or it would update when shown next time (figure content change would be noticed
                            //by user)
                            if (smart_grid)
                            {
                                MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    MW.Hide();
                                }));
                            }

                            if (grid_type == GridType.hexagonal)
                            {
                                int two_rows = (int)(i / (cols_nr * 2 - 1));
                                int m = i % (cols_nr * 2 - 1);

                                if (m < cols_nr)//first
                                {
                                    x = (int)Math.Round((double)(screen_width * (m + 0.5) / cols_nr));
                                    y = (int)Math.Round((double)(screen_height * (two_rows + 0.5) / rows_nr));
                                }
                                else
                                {
                                    x = (int)Math.Round((double)(screen_width * (m - cols_nr + 1) / cols_nr));
                                    y = (int)Math.Round((double)(screen_height * (two_rows + 1) / rows_nr));
                                }
                            }
                            else
                            {
                                double offest_v = 0.5;
                                double offest_h = 0.5;

                                if (grid_type == MainWindow.GridType.square_horizontal_precision
                                    || grid_type == MainWindow.GridType.square_combined_precision)
                                {
                                    int col_nr = i % cols_nr;
                                    if (col_nr % 3 == 0)
                                        offest_v = 0.25;
                                    else if (col_nr % 3 == 1)
                                        offest_v = 0.5;
                                    else
                                        offest_v = 0.75;
                                }
                                if (grid_type == MainWindow.GridType.square_vertical_precision
                                    || grid_type == MainWindow.GridType.square_combined_precision)
                                {
                                    int row_nr = (int)(i / cols_nr);
                                    if (row_nr % 3 == 0)
                                        offest_h = 0.25;
                                    else if (row_nr % 3 == 1)
                                        offest_h = 0.5;
                                    else
                                        offest_h = 0.75;
                                }

                                x = (int)Math.Round((double)(screen_width *
                                    (i % cols_nr + offest_h) / cols_nr));
                                y = (int)Math.Round((double)(screen_height *
                                    ((int)(i / cols_nr) + offest_v) / rows_nr));
                            }

                            x += offset_x;
                            y += offset_y;

                            if (x < 0) x = 0;
                            else if (x > screen_width - 1) x = screen_width - 1;

                            if (y < 0) y = 0;
                            else if (y > screen_height - 1) y = screen_height - 1;

                            last_x = x;
                            last_y = y;

                            if (action == ActionX.left_click)
                            {
                                LMBClick(x, y);
                            }
                            else if (action == ActionX.right_click)
                            {
                                RMBClick(x, y);
                            }
                            else if (action == ActionX.double_left_click)
                            {
                                DLMBClick(x, y);
                            }
                            else if (action == ActionX.triple_left_click)
                            {
                                LMBClick(x, y);
                                LMBClick(x, y);
                                LMBClick(x, y);
                            }
                            else if (action == ActionX.ctrl_left_click)
                            {
                                sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                                LMBClick(x, y);
                                sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                            }
                            else if (action == ActionX.move_mouse)
                            {
                                real_mouse_move(x, y);
                            }
                            else if (action == ActionX.hold_left)
                            {
                                if (read_status) ss.SpeakAsync("holding left");
                                LMBHold(x, y);
                            }
                            else if (action == ActionX.hold_right)
                            {
                                if (read_status) ss.SpeakAsync("holding right");
                                RMBHold(x, y);
                            }
                            else if (action == ActionX.drag_and_drop)
                            {
                                if (d == 0)
                                {
                                    drag_x = x;
                                    drag_y = y;

                                    if (read_status) ss.SpeakAsync("drop");
                                }
                                else
                                {
                                    move_mouse(drag_x, drag_y);
                                    LMBHold(drag_x, drag_y);

                                    freeze_mouse(100);
                                    real_mouse_move(x, y);
                                    freeze_mouse(100);
                                    left_up();
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                int x = last_x;
                int y = last_y;
                action = last_action;

                if (action == ActionX.left_click)
                {
                    LMBClick(x, y);
                }
                else if (action == ActionX.right_click)
                {
                    RMBClick(x, y);
                }
                else if (action == ActionX.double_left_click)
                {
                    DLMBClick(x, y);
                }
                else if (action == ActionX.triple_left_click)
                {
                    LMBClick(x, y);
                    LMBClick(x, y);
                    LMBClick(x, y);
                }
                else if (action == ActionX.ctrl_left_click)
                {
                    sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                    LMBClick(x, y);
                    sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                }
                else if (action == ActionX.move_mouse)
                {
                    real_mouse_move(x, y);
                }
                else if (action == ActionX.hold_left)
                {
                    if (read_status) ss.SpeakAsync("holding left");
                    LMBHold(x, y);
                }
                else if (action == ActionX.hold_right)
                {
                    if (read_status) ss.SpeakAsync("holding right");
                    RMBHold(x, y);
                }
                else if (action == ActionX.drag_and_drop)
                {
                    move_mouse(drag_x, drag_y);
                    LMBHold(drag_x, drag_y);

                    freeze_mouse(100);
                    real_mouse_move(x, y);
                    freeze_mouse(100);
                    left_up();
                }
            }

            if (repeat_action_indefinitely == false)
            {
                action = ActionX.none;
                execution_nr = 0;
            }
            else
            {
                execution_nr++;
            }
        }

        int get_symbols_score(char s1, char s2)
        {
            int score;
            int ind1, ind2;

            ind1 = get_index_by_symbol(s1);
            ind2 = get_index_by_symbol(s2);

            if (keyboard_layout == "Any")
            {
                if (ind1 <= 25)
                    score = ind1;
                else
                    score = ind1 * 10;

                if (ind2 <= 25)
                    score += ind2;
                else
                    score += ind2 * 10;
            }
            else
            {
                if (ind1 <= 24)
                    score = ind1;
                else
                    score = ind1 * 10;

                if (ind2 <= 24)
                    score += ind2;
                else
                    score += ind2 * 10;
            }

            return score;
        }

        //find index with lower count and best score (for that count)
        int find_index_with_lower_count(int grid_ind, uint count, int best_score)
        {
            int ind = -1;
            int score;
            int ind1, ind2;

            for (int i = 0; i < grids[grid_ind].elements.Count; i++)
            {
                if (grids[grid_ind].elements[i].count <= count)
                {
                    score = get_symbols_score(grids[grid_ind].elements[i].symbols[0],
                        grids[grid_ind].elements[i].symbols[1]);

                    if (score < best_score)
                    {
                        best_score = score;
                        ind = i;
                    }
                }
            }

            return ind;
        }

        int get_index_by_symbol(char c)
        {
            string s = c.ToString();

            for (int i = 0; i < grid_alphabet.Count; i++)
            {
                if (s == grid_alphabet[i].symbol)
                    return i;
            }

            return -1;
        }

        private void save_grids()
        {
            FileStream fs = null;
            StreamWriter sw = null;

            string type_name = get_grid_folder_name_by_type(grid_type);
            string kl_name = "Any"; //kl = keyboard layout

            if (keyboard_layout.Substring(0, 2) == "US")
                kl_name = "US";

            string folder_path = Path.Combine(new string[] { System.IO.Path.Combine(saving_folder_path, 
                grids_foldername), kl_name, type_name, desired_figures_nr.ToString()});

            try
            {
                if (Directory.Exists(folder_path) == false)
                {
                    Directory.CreateDirectory(folder_path);
                }

                for (int i = 1; i < grids.Count; i++)
                {
                    fs = new FileStream(System.IO.Path.Combine(new string[] {
                        folder_path, grids[i].process_name + ".txt" }),
                        FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs);

                    sw.WriteLine(grids[i].process_name);
                    sw.WriteLine(grids[i].count);

                    foreach (Grid_element element in grids[i].elements)
                    {
                        sw.WriteLine(element.symbols);
                        sw.WriteLine(element.count);
                    }

                    sw.Close();
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW004", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex2) { }
            }
        }

        private void load_grids()
        {
            FileStream fs = null;
            StreamReader sr = null;

            string type_name = get_grid_folder_name_by_type(grid_type);
            string kl_name = "Any"; //kl = keyboard layout

            if (keyboard_layout.Substring(0, 2) == "US")
                kl_name = "US";

            string folder_path = Path.Combine(new string[] { System.IO.Path.Combine(saving_folder_path,
                grids_foldername), kl_name, type_name, desired_figures_nr.ToString()});            

            try
            {
                if (Directory.Exists(folder_path))
                {
                    string[] files = Directory.GetFiles(folder_path);
                    bool cancel = false;

                    foreach (string file in files)
                    {
                        cancel = false;

                        fs = new FileStream(file, FileMode.Open, FileAccess.Read);
                        sr = new StreamReader(fs);

                        if (sr.EndOfStream)
                        {
                            sr.Close();
                            fs.Close();
                            continue;
                        }
                        Process_grid g = new Process_grid(sr.ReadLine());

                        if (sr.EndOfStream)
                        {
                            sr.Close();
                            fs.Close();
                            continue;
                        }
                        int count;
                        if (int.TryParse(sr.ReadLine(), out count) == false)
                        {
                            sr.Close();
                            fs.Close();
                            continue;
                        }
                        if (count != grids[0].count)
                        {
                            sr.Close();
                            fs.Close();
                            continue;
                        }
                        g.count = count;

                        g.elements = new List<Grid_element>();
                        string s, w;
                        uint count2;

                        while (sr.EndOfStream == false)
                        {
                            s = sr.ReadLine();

                            if (sr.EndOfStream)
                            {
                                cancel = true;
                                break;
                            }

                            if (uint.TryParse(sr.ReadLine(), out count2) == false)
                            {
                                cancel = true;
                                break;
                            }

                            Grid_element element = new Grid_element(s);
                            element.count = count2;
                            g.elements.Add(element);
                        }

                        if (cancel == false)
                            grids.Add(g);

                        sr.Close();
                        fs.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW005", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    sr.Close();
                    fs.Close();
                }
                catch (Exception ex2) { }
            }
        }        

        string get_grid_folder_name_by_type(GridType gt)
        {
            if (gt == GridType.hexagonal)
                return folder_name_grid_hexagonal;
            else if (gt == GridType.square)
                return folder_name_grid_square;
            else if (gt == GridType.square_combined_precision)
                return folder_name_grid_combined;
            else if (gt == GridType.square_horizontal_precision)
                return folder_name_grid_horizontal;
            else if (gt == GridType.square_vertical_precision)
                return folder_name_grid_vertical;
            else return null;
        }
    }
}