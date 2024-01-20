using System.Collections.Generic;
using System.Windows;
using WindowsInput.Native;

namespace Aspiring_Keyboard
{
    public class Grid_element
	{
		public string symbols;
		public uint count = 0;

		public Grid_element(string s)
		{
			symbols = s;
		}
	}

	public partial class MainWindow : Window
    {
		public class Process_grid
		{
			public string process_name;
			public List<Grid_element> elements = new List<Grid_element>();
			public int count;

			public Process_grid(string Process_name)
			{
				process_name = Process_name;
			}
		}

		public class Grid_Symbol
		{
			public string symbol;
			public VirtualKeyCode vkc;
			public bool shift;

			public Grid_Symbol(string s, VirtualKeyCode v, bool Shift = false)
			{
				symbol = s;
				vkc = v;
				shift = Shift;
			}
		}
		
		public enum ActionX
		{
			left_click,
			right_click,
			double_left_click,
			triple_left_click,
			center_left_click,
			ctrl_left_click,
			move_mouse,
			drag_and_drop,
			hold_left,
			hold_right,
			scroll_up,
            scroll_down,
            scroll_left,
            scroll_right,
            none
		}

		public enum GridType
        {
			hexagonal,
			square,
			square_horizontal_precision,
			square_vertical_precision,
			square_combined_precision
		}

		void create_normal_grid_alphabet()
		{
			grid_alphabet = new List<Grid_Symbol>();

			grid_alphabet.Add(new Grid_Symbol("a", VirtualKeyCode.VK_A));
			grid_alphabet.Add(new Grid_Symbol("b", VirtualKeyCode.VK_B));
			grid_alphabet.Add(new Grid_Symbol("c", VirtualKeyCode.VK_C));
			grid_alphabet.Add(new Grid_Symbol("d", VirtualKeyCode.VK_D));
			grid_alphabet.Add(new Grid_Symbol("e", VirtualKeyCode.VK_E));
			grid_alphabet.Add(new Grid_Symbol("f", VirtualKeyCode.VK_F));
			grid_alphabet.Add(new Grid_Symbol("g", VirtualKeyCode.VK_G));
			grid_alphabet.Add(new Grid_Symbol("h", VirtualKeyCode.VK_H));
			grid_alphabet.Add(new Grid_Symbol("i", VirtualKeyCode.VK_I));
			grid_alphabet.Add(new Grid_Symbol("j", VirtualKeyCode.VK_J));
			grid_alphabet.Add(new Grid_Symbol("k", VirtualKeyCode.VK_K));
			grid_alphabet.Add(new Grid_Symbol("l", VirtualKeyCode.VK_L));
			grid_alphabet.Add(new Grid_Symbol("n", VirtualKeyCode.VK_N));
			grid_alphabet.Add(new Grid_Symbol("o", VirtualKeyCode.VK_O));
			grid_alphabet.Add(new Grid_Symbol("p", VirtualKeyCode.VK_P));
			grid_alphabet.Add(new Grid_Symbol("q", VirtualKeyCode.VK_Q));
			grid_alphabet.Add(new Grid_Symbol("r", VirtualKeyCode.VK_R));
			grid_alphabet.Add(new Grid_Symbol("s", VirtualKeyCode.VK_S));
			grid_alphabet.Add(new Grid_Symbol("t", VirtualKeyCode.VK_T));
			grid_alphabet.Add(new Grid_Symbol("u", VirtualKeyCode.VK_U));
			grid_alphabet.Add(new Grid_Symbol("v", VirtualKeyCode.VK_V));
			grid_alphabet.Add(new Grid_Symbol("x", VirtualKeyCode.VK_X));
			grid_alphabet.Add(new Grid_Symbol("y", VirtualKeyCode.VK_Y));
			grid_alphabet.Add(new Grid_Symbol("z", VirtualKeyCode.VK_Z));
			grid_alphabet.Add(new Grid_Symbol(";", VirtualKeyCode.OEM_1)); //not supported by all keyboard layouts
            grid_alphabet.Add(new Grid_Symbol("/", VirtualKeyCode.OEM_2));
			grid_alphabet.Add(new Grid_Symbol("[", VirtualKeyCode.OEM_4)); //not supported by all keyboard layouts
            grid_alphabet.Add(new Grid_Symbol("8", VirtualKeyCode.VK_8));
            grid_alphabet.Add(new Grid_Symbol("5", VirtualKeyCode.VK_5));
            grid_alphabet.Add(new Grid_Symbol("-", VirtualKeyCode.OEM_MINUS));
            grid_alphabet.Add(new Grid_Symbol("A", VirtualKeyCode.VK_A, true));
			grid_alphabet.Add(new Grid_Symbol("B", VirtualKeyCode.VK_B, true));
			grid_alphabet.Add(new Grid_Symbol("C", VirtualKeyCode.VK_C, true));
			grid_alphabet.Add(new Grid_Symbol("D", VirtualKeyCode.VK_D, true));
			grid_alphabet.Add(new Grid_Symbol("E", VirtualKeyCode.VK_E, true));
			grid_alphabet.Add(new Grid_Symbol("F", VirtualKeyCode.VK_F, true));
			grid_alphabet.Add(new Grid_Symbol("G", VirtualKeyCode.VK_G, true));
			grid_alphabet.Add(new Grid_Symbol("H", VirtualKeyCode.VK_H, true));
			grid_alphabet.Add(new Grid_Symbol("I", VirtualKeyCode.VK_I, true));
			grid_alphabet.Add(new Grid_Symbol("J", VirtualKeyCode.VK_J, true));
			grid_alphabet.Add(new Grid_Symbol("K", VirtualKeyCode.VK_K, true));
			grid_alphabet.Add(new Grid_Symbol("L", VirtualKeyCode.VK_L, true));
			grid_alphabet.Add(new Grid_Symbol("N", VirtualKeyCode.VK_N, true));
			grid_alphabet.Add(new Grid_Symbol("O", VirtualKeyCode.VK_O, true));
			grid_alphabet.Add(new Grid_Symbol("P", VirtualKeyCode.VK_P, true));
			grid_alphabet.Add(new Grid_Symbol("Q", VirtualKeyCode.VK_Q, true));
			grid_alphabet.Add(new Grid_Symbol("R", VirtualKeyCode.VK_R, true));
			grid_alphabet.Add(new Grid_Symbol("S", VirtualKeyCode.VK_S, true));
			grid_alphabet.Add(new Grid_Symbol("T", VirtualKeyCode.VK_T, true));
			grid_alphabet.Add(new Grid_Symbol("U", VirtualKeyCode.VK_U, true));
			grid_alphabet.Add(new Grid_Symbol("V", VirtualKeyCode.VK_V, true));
			grid_alphabet.Add(new Grid_Symbol("X", VirtualKeyCode.VK_X, true));
			grid_alphabet.Add(new Grid_Symbol("Y", VirtualKeyCode.VK_Y, true));
			grid_alphabet.Add(new Grid_Symbol("Z", VirtualKeyCode.VK_Z, true));
			grid_alphabet.Add(new Grid_Symbol("?", VirtualKeyCode.OEM_2, true)); //not supported by all keyboard layouts
			grid_alphabet.Add(new Grid_Symbol("\"", VirtualKeyCode.OEM_7, true));  //not supported by all keyboard layouts
            grid_alphabet.Add(new Grid_Symbol("*", VirtualKeyCode.VK_8, true)); // not supported by all keyboard layouts
            grid_alphabet.Add(new Grid_Symbol("_", VirtualKeyCode.OEM_MINUS, true)); //not supported by all keyboard layouts
			//grid_alphabet.Add(new Grid_Symbol("\"", VirtualKeyCode.OEM_3, true));

            //grid_alphabet.Add(new Grid_Symbol("1", VirtualKeyCode.VK_1));
            //grid_alphabet.Add(new Grid_Symbol("2", VirtualKeyCode.VK_2));
            //grid_alphabet.Add(new Grid_Symbol("3", VirtualKeyCode.VK_3));
            //grid_alphabet.Add(new Grid_Symbol("4", VirtualKeyCode.VK_4));
            //grid_alphabet.Add(new Grid_Symbol("5", VirtualKeyCode.VK_5));
            //grid_alphabet.Add(new Grid_Symbol("6", VirtualKeyCode.VK_6));
            //grid_alphabet.Add(new Grid_Symbol("7", VirtualKeyCode.VK_7));
            //grid_alphabet.Add(new Grid_Symbol("8", VirtualKeyCode.VK_8));
            //grid_alphabet.Add(new Grid_Symbol("9", VirtualKeyCode.VK_9));
            //grid_alphabet.Add(new Grid_Symbol("0", VirtualKeyCode.VK_0));
        }

		System.Diagnostics.Stopwatch watch;

		void start_time()
        {
			watch = System.Diagnostics.Stopwatch.StartNew();
		}

		void stop_time()
		{
			watch.Stop();
			var elapsedMs = watch.ElapsedMilliseconds;

			//TB.Dispatcher.Invoke(DispatcherPriority.Normal,
			//	new Action(() => { TB.Text = elapsedMs.ToString(); })); 
		}
	}
}
