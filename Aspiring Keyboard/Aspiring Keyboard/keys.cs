using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using WindowsInput;
using WindowsInput.Native;

namespace Aspiring_Keyboard
{
    public partial class MainWindow : Window
    {
        Thread THRkeymaster;

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        private const byte VK_MENU = 0x12;
        private const byte VK_F4 = 0x73;
        private const byte VK_OEM_PLUS = 0xBB; //this presses =, not plus !!!
        //private const byte  = ;
        private const int KEYEVENTF_KEYUP = 0x2;
        private const int KEYEVENTF_KEYDOWN = 0x0;

        void key_press(VirtualKeyCode vkc, bool async, int down_ms = 75)
        {
            if (async)
            {
                THRkeymaster = new Thread(() => key_press(vkc, down_ms));
                THRkeymaster.Start();
            }
            else
                key_press(vkc, down_ms);
        }

        void key_press(VirtualKeyCode vkc, int down_ms = 75)
        {
            //left alt in WindowsInput library is bugged (keyup doesn't work)
            if (vkc == VirtualKeyCode.LMENU)
            {
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                Thread.Sleep(down_ms);
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            }
            //Plus in WindowsInput library is bugged
            else if (vkc == VirtualKeyCode.OEM_PLUS)
            {
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYUP, 0);
                Thread.Sleep(down_ms);
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYDOWN, 0);
            }
            else
            {
                sim.Keyboard.KeyDown(vkc);
                Thread.Sleep(down_ms);
                sim.Keyboard.KeyUp(vkc);
            }
        }

        void key_down(VirtualKeyCode vkc)
        {
            //left alt in WindowsInput library is bugged (keyup doesn't work)
            if (vkc == VirtualKeyCode.LMENU)
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            //Plus in WindowsInput library is bugged
            else if (vkc == VirtualKeyCode.OEM_PLUS)
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYDOWN, 0);
            else
                sim.Keyboard.KeyDown(vkc);
        }

        void key_up(VirtualKeyCode vkc)
        {
            //left alt in WindowsInput library is bugged (keyup doesn't work)
            if (vkc == VirtualKeyCode.LMENU)
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
            //Plus in WindowsInput library is bugged
            else if (vkc == VirtualKeyCode.OEM_PLUS)
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYUP, 0);
            else
                sim.Keyboard.KeyUp(vkc);
        }

        void release_buttons_and_keys()
        {
            foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(typeof(VirtualKeyCode)))
            {
                if (sim.InputDeviceState.IsKeyDown(vkc))
                    sim.Keyboard.KeyUp(vkc);
            }

            //left alt in WindowsInput library is bugged (keyup doesn't work)
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LMENU))
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);

            //Plus in WindowsInput library is bugged
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.OEM_PLUS))
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYUP, 0);
        }

        void release_buttons()
        {
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON))
            {
                left_up();
            }
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON))
            {
                right_up();
            }
        }
    }
}