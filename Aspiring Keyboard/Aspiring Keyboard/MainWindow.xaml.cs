﻿//highest error nr: MW005
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

//resolved difficult issue: something sometimes messes up mousegrid and it doesn't appear

namespace Aspiring_Keyboard
{
    public partial class MainWindow : Window
    {
        const bool check_if_already_running = true;

        const string prog_name = "Aspiring Keyboard";
        const string prog_version = "1.0-Alpha.1";
        const string copyright_text = "Copyright © 2023 Mikołaj Magowski. All rights reserved.";
        const string filename_settings = "settings_ak.txt";
        const string grids_foldername = "grids";
        const bool resized_grid = true;
        const bool movable_grid = true;
        
        bool enabled = true;

        const string icon_red = "pack://application:,,,/Aspiring Keyboard;component/images/ak red.ico";
        const string icon_green = "pack://application:,,,/Aspiring Keyboard;component/images/ak green.ico";
        const string icon_blue = "pack://application:,,,/Aspiring Keyboard;component/images/ak blue.ico";

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
        int max_figures_nr = 1;
        int desired_figures_nr = 1;
        GridType grid_type = GridType.hexagonal;
        bool smart_grid = true;
        double offset = 0.25;
        double np_offset = 0.25; //offset by percentage of figure height (normal precision)
        double hp_offset = 0.025; //high precision offset
        int offset_x = 0; //grid offset
        int offset_y = 0;
        int grid_lines = 0; //0-2 (0 - no lines, 1 - dotted lines, 2 - normal lines)
        
        Color color2 = Color.FromRgb(0, 0, 0); //font color
        //Color color1 = Color.FromRgb(255, 255, 255);
        Color color1 = Color.FromRgb(225, 225, 225); //bg color

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
        Command last_command = Command.none;

        List<Process_grid> grids = new List<Process_grid>();
        int grid_ind = 0;
        int prev_installed_apps_count = -1;

        SpeechSynthesizer ss = new SpeechSynthesizer();

        bool green_mode = true;
        bool read_status = true;

        Command left_shift_command_a = Command.none;
        Command right_shift_command_a = Command.none;
        Command left_alt_command_a = Command.none;
        Command right_alt_command_a = Command.none;
        Command left_ctrl_command_a = Command.none;
        Command right_ctrl_command_a = Command.none;

        Command left_shift_command_b = Command.none;
        Command right_shift_command_b = Command.none;
        Command left_alt_command_b = Command.none;
        Command right_alt_command_b = Command.none;
        Command left_ctrl_command_b = Command.none;
        Command right_ctrl_command_b = Command.none;

        Command command = Command.none;

        bool repeat_command_indefinitely = false;
        
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

                if (check_if_already_running)
                {
                    is_program_already_running();
                }

                InitializeComponent();

                this.Title = prog_name + " v." + prog_version;
                //Lprogram_name.Content = prog_name;
                //Linstalled_version.Content = "Installed version: " + prog_version;
                //Lhomepage.Content = "Homepage: " + Middle_Man.url_homepage;
                //Lcopyright.Content = copyright_text;

                TBcontrol_keys.Text = "Pause Break - turn on/off"
                    + "\r\nScroll Lock - change mode"
                    + "\r\nLAlt + RAlt - release both mouse buttons"
                    + "\r\nInsert - repeat last command"
                    + "\r\nCaps Lock - left click without losing focus"
                    + "\r\nLShift + RShift - browser forward button"
                    + "\r\nNum Lock - browser back button"
                    + "\r\nLShift + Caps Lock - left shift command in infinity mode"
                    + "\r\nRShift + Caps Lock - right shift command in infinity mode"
                    + "\r\nCaps Lock + Up/Down arrow - scroll up/down"
                    + "\r\nWhen mousegrid is visible:"
                    + "\r\n- Arrow keys - move mousegrid"
                    + "\r\n- Caps Lock  - change mousegrid movement mode"
                    + "\r\n- Escape - cancel command and hide mousegrid";

                foreach (GridType type in (GridType[])Enum.GetValues(typeof(GridType)))
                {
                    CBtype.Items.Add(type.ToString().Replace("_", " ").FirstCharToUpper());
                }

                foreach (Command command in (Command[])Enum.GetValues(typeof(Command)))
                {
                    CBlshift_command_a.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBrshift_command_a.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlalt_command_a.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBralt_command_a.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlctrl_command_a.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBrctrl_command_a.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlshift_command_b.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBrshift_command_b.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlalt_command_b.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBralt_command_b.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBlctrl_command_b.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                    CBrctrl_command_b.Items.Add(command.ToString().Replace("_", " ").FirstCharToUpper());
                }

                CBlines.Items.Add("None");
                CBlines.Items.Add("Dotted");
                CBlines.Items.Add("Solid");

                create_normal_grid_alphabet();
                grid_symbols_limit = grid_alphabet.Count;

                max_figures_nr = (int)Math.Pow((double)grid_symbols_limit, 2);
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

                Create_Grid();

                if (smart_grid)
                    load_grids();

                green_mode = true;
                
                THRmonitor = new Thread(new ThreadStart(monitor));
                THRmonitor.Start();

                if (read_status) ss.SpeakAsync("Green mode enabled");

                CBlshift_command_a.Focus();
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
                        else if (last_command != Command.none &&
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

                                        break;
                                    }
                                    else if (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                    {
                                        repeat_command_indefinitely = true;
                                        break;
                                    }
                                }

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.LShiftKey)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.RShiftKey)
                                    || IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                {
                                    Thread.Sleep(10);
                                }

                                if (repeat_command_indefinitely)
                                {
                                    Thread.Sleep(10);

                                    sim.Keyboard.KeyPress(VirtualKeyCode.CAPITAL);
                                }

                                Thread.Sleep(10);
                                
                                if (cancel == false)
                                {
                                    if(green_mode)
                                        command = left_shift_command_a;
                                    else
                                        command = left_shift_command_b;
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

                                        break;
                                    }
                                    else if (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                    {
                                        repeat_command_indefinitely = true;
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

                                if (repeat_command_indefinitely)
                                {
                                    Thread.Sleep(10);

                                    sim.Keyboard.KeyPress(VirtualKeyCode.CAPITAL);
                                }

                                if (cancel == false)
                                {
                                    if (green_mode)
                                        command = right_shift_command_a;
                                    else
                                        command = right_shift_command_b;
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
                                        release_buttons();

                                        if (read_status) ss.SpeakAsync("Mouse buttons released.");
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
                                        command = left_alt_command_a;
                                    else
                                        command = left_alt_command_b;
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
                                        release_buttons();

                                        if (read_status) ss.SpeakAsync("Mouse buttons released.");
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
                                        command = right_alt_command_a;
                                    else
                                        command = right_alt_command_b;
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
                                        command = left_ctrl_command_a;
                                    else
                                        command = left_ctrl_command_b;
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
                                        command = right_ctrl_command_a;
                                    else
                                        command = right_ctrl_command_b;
                                }
                            }
                            else if (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                            {
                                fix_very_rare_problem();

                                command = Command.left_click;

                                while (IsKeyPushedDown(System.Windows.Forms.Keys.CapsLock))
                                {
                                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(
                                        typeof(VirtualKeyCode)))
                                    {
                                        if (vkc == VirtualKeyCode.UP
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            sim.Mouse.VerticalScroll(4);
                                            command = Command.scroll_up;
                                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Up))
                                                Thread.Sleep(10);
                                        }
                                        else if (vkc == VirtualKeyCode.DOWN
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            sim.Mouse.VerticalScroll(-4);
                                            command = Command.scroll_down;
                                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Down))
                                                Thread.Sleep(10);
                                        }
                                        /* PROBLEMATIC
                                        else if (vkc == VirtualKeyCode.LEFT
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            sim.Mouse.HorizontalScroll(-300);
                                            command = Command.scroll_left;
                                            while (IsKeyPushedDown(System.Windows.Forms.Keys.Left))
                                                Thread.Sleep(10);
                                        }
                                        else if (vkc == VirtualKeyCode.RIGHT
                                            && sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            sim.Mouse.HorizontalScroll(300);
                                            command = Command.scroll_right;
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

                                if (command == Command.left_click)
                                {
                                    dont_lose_focus = true;
                                }
                                
                                Thread.Sleep(10);

                                sim.Keyboard.KeyPress(VirtualKeyCode.CAPITAL);

                                if (read_status)
                                {
                                    if (command == Command.left_click)
                                    {
                                        ss.SpeakAsync("focus preserved");
                                    }
                                }

                                if (command != Command.left_click)
                                    command = Command.none;
                            }

                            do
                            {
                                if (command != Command.none && cancel == false)
                                {
                                    adv_mouse(false, dont_lose_focus);

                                    //sometimes RMB may stay pressed and mess things up
                                    release_buttons_and_keys();
                                }
                            }
                            while (repeat_command_indefinitely);
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

        void close_that()
        {
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(VK_F4, 0, KEYEVENTF_KEYDOWN, 0);
            Thread.Sleep(75);
            keybd_event(VK_F4, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);

            // admin req
            /*IntPtr handle = GetForegroundWindow();

            Process[] arr = Process.GetProcesses();
            Process process = null;

            try
            {
                foreach (Process p in arr)
                {
                    IntPtr h = p.MainWindowHandle;

                    if (h == handle)
                    {
                        process = p;
                        process.CloseMainWindow();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    if (process != null)
                        process.Kill(); //kill nie zawsze zadziała - może wystąpić odmowa dostępu
                }
                catch (Exception ex2) { }
            }
            */
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
            MW = new MouseGrid(grid_width, grid_height, grid_lines, grid_type, font_family, font_size, color1,
                color2, rows_nr, cols_nr, figure_width, figure_height, grids[0].elements);
            //MW.regenerate_grid_symbols();
        }

        int drag_x = 0, drag_y = 0;
        int last_x = 0, last_y = 0;
        int execution_nr = 0;

        void adv_mouse(bool repeat_last_action = false, bool dont_lose_focus = false)
        {
            if (repeat_last_action == false)
            {
                IntPtr handle2 = GetForegroundWindow();
                last_command = command;

                int d = 1;
                if (command == Command.drag_and_drop)
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

                        if (repeat_command_indefinitely)
                            status = "indefinite ";

                        if (command == Command.drag_and_drop && d == 0)
                        {
                            status += "drag";
                        }
                        else if (command == Command.double_left_click)
                        {
                            status += "double";
                        }
                        else if (command == Command.triple_left_click)
                        {
                            status += "triple";
                        }
                        else if (command == Command.center_left_click)
                        {
                            status += "center";
                        }
                        else if (command == Command.hold_left)
                        {
                            status += "hold left";
                        }
                        else if (command == Command.hold_right)
                        {
                            status += "hold right";
                        }
                        else if (command == Command.move_mouse)
                        {
                            status += "move mouse";
                        }
                        else if (command == Command.right_click)
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
                            repeat_command_indefinitely = false;
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
                                grids[grid_ind].elements[i].count++;

                                int ind = -1;

                                int score = get_symbsols_score(grids[grid_ind].elements[i].symbols[0],
                                    grids[grid_ind].elements[i].symbols[1]);

                                ind = find_lowest_index_with_lower_count(grid_ind,
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

                            if (command == Command.left_click)
                            {
                                LMBClick(x, y);
                            }
                            else if (command == Command.right_click)
                            {
                                RMBClick(x, y);
                            }
                            else if (command == Command.double_left_click)
                            {
                                DLMBClick(x, y);
                            }
                            else if (command == Command.triple_left_click)
                            {
                                LMBClick(x, y);
                                LMBClick(x, y);
                                LMBClick(x, y);
                            }
                            else if (command == Command.center_left_click)
                            {
                                LMBClick((int)(screen_width / 2), (int)(screen_height / 2));
                            }
                            else if (command == Command.ctrl_left_click)
                            {
                                sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                                LMBClick(x, y);
                                sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                            }
                            else if (command == Command.move_mouse)
                            {
                                real_mouse_move(x, y);
                            }
                            else if (command == Command.hold_left)
                            {
                                if (read_status) ss.SpeakAsync("holding left");
                                LMBHold(x, y);
                            }
                            else if (command == Command.hold_right)
                            {
                                if (read_status) ss.SpeakAsync("holding right");
                                RMBHold(x, y);
                            }
                            else if (command == Command.drag_and_drop)
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
                command = last_command;

                if (command == Command.left_click)
                {
                    LMBClick(x, y);
                }
                else if (command == Command.right_click)
                {
                    RMBClick(x, y);
                }
                else if (command == Command.double_left_click)
                {
                    DLMBClick(x, y);
                }
                else if (command == Command.triple_left_click)
                {
                    LMBClick(x, y);
                    LMBClick(x, y);
                    LMBClick(x, y);
                }
                else if (command == Command.center_left_click)
                {
                    LMBClick((int)(screen_width / 2), (int)(screen_height / 2));
                }
                else if (command == Command.ctrl_left_click)
                {
                    sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                    LMBClick(x, y);
                    sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                }
                else if (command == Command.move_mouse)
                {
                    real_mouse_move(x, y);
                }
                else if (command == Command.hold_left)
                {
                    if (read_status) ss.SpeakAsync("holding left");
                    LMBHold(x, y);
                }
                else if (command == Command.hold_right)
                {
                    if (read_status) ss.SpeakAsync("holding right");
                    RMBHold(x, y);
                }
                else if (command == Command.drag_and_drop)
                {
                    move_mouse(drag_x, drag_y);
                    LMBHold(drag_x, drag_y);

                    freeze_mouse(100);
                    real_mouse_move(x, y);
                    freeze_mouse(100);
                    left_up();
                }
            }

            if (repeat_command_indefinitely == false)
            {
                command = Command.none;
                execution_nr = 0;
            }
            else
            {
                execution_nr++;
            }
        }

        int get_symbsols_score(char s1, char s2)
        {
            int score;
            int ind1, ind2;

            ind1 = get_index_by_symbol(s1);
            ind2 = get_index_by_symbol(s2);

            if (ind1 <= 26)
                score = ind1;
            else
                score = ind1 * 10;

            if (ind2 <= 26)
                score += ind2;
            else
                score += ind2 * 10;

            return score;
        }

        int find_lowest_index_with_lower_count(int grid_ind, uint count, int best_score)
        {
            int ind = -1;
            int score;
            int ind1, ind2;

            for (int i = 0; i < grids[grid_ind].elements.Count; i++)
            {
                if (grids[grid_ind].elements[i].count < count)
                {
                    score = get_symbsols_score(grids[grid_ind].elements[i].symbols[0],
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
            string folder_path = System.IO.Path.Combine(saving_folder_path, grids_foldername);

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
            string folder_path = System.IO.Path.Combine(saving_folder_path, grids_foldername);

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
    }
}