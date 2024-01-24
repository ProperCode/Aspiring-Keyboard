using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using WindowsInput.Native;

namespace Aspiring_Keyboard
{
    public partial class MainWindow : Window
    {
        Thread THRkeymaster;

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
            sim.Keyboard.KeyDown(vkc);
            Thread.Sleep(down_ms);
            sim.Keyboard.KeyUp(vkc);
        }

        void key_down(VirtualKeyCode vkc)
        {
            sim.Keyboard.KeyDown(vkc);
        }

        void key_up(VirtualKeyCode vkc)
        {
            sim.Keyboard.KeyUp(vkc);
        }

        void release_buttons_and_keys()
        {
            foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(typeof(VirtualKeyCode)))
            {
                if (sim.InputDeviceState.IsKeyDown(vkc))
                    sim.Keyboard.KeyUp(vkc);
            }
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