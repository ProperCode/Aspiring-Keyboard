using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using GregsStack.InputSimulatorStandard.Native;
using System.Threading;

namespace GregsStack.InputSimulatorStandard
{
    using Native;

    /// <inheritdoc />
    /// <summary>
    /// An implementation of <see cref="IInputDeviceStateAdapter" /> for Windows by calling the native <see cref="NativeMethods.GetKeyState(ushort)" /> and <see cref="NativeMethods.GetAsyncKeyState(ushort)" /> methods.
    /// </summary>
    public class WindowsInputDeviceStateAdapter : IInputDeviceStateAdapter
    {
        /// <inheritdoc />
        /// <summary>
        /// Determines whether the specified key is up or down by calling the GetKeyState function. (See: http://msdn.microsoft.com/en-us/library/ms646301(VS.85).aspx)
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode" /> for the key.</param>
        /// <returns><c>true</c> if the key is down; otherwise, <c>false</c>.</returns>
        /// <remarks>
        /// The key status returned from this function changes as a thread reads key messages from its message queue. The status does not reflect the interrupt-level state associated with the hardware. Use the GetAsyncKeyState function to retrieve that information. 
        /// An application calls GetKeyState in response to a keyboard-input message. This function retrieves the state of the key when the input message was generated. 
        /// To retrieve state information for all the virtual keys, use the GetKeyboardState function. 
        /// An application can use the virtual-key code constants VK_SHIFT, VK_CONTROL, and VK_MENU as values for Bthe nVirtKey parameter. This gives the status of the SHIFT, CTRL, or ALT keys without distinguishing between left and right. An application can also use the following virtual-key code constants as values for nVirtKey to distinguish between the left and right instances of those keys. 
        /// VK_LSHIFT
        /// VK_RSHIFT
        /// VK_LCONTROL
        /// VK_RCONTROL
        /// VK_LMENU
        /// VK_RMENU
        /// These left- and right-distinguishing constants are available to an application only through the GetKeyboardState, SetKeyboardState, GetAsyncKeyState, GetKeyState, and MapVirtualKey functions. 
        /// </remarks>
        public bool IsKeyDown(VirtualKeyCode keyCode)
        {
            var result = NativeMethods.GetKeyState((ushort)keyCode);
            return (result < 0);
        }

        /// <inheritdoc />
        /// <summary>
        /// Determines whether the specified key is up or down by calling the <see cref="NativeMethods.GetKeyState(ushort)" /> function. (See: http://msdn.microsoft.com/en-us/library/ms646301(VS.85).aspx)
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode" /> for the key.</param>
        /// <returns><c>true</c> if the key is up; otherwise, <c>false</c>.</returns>
        /// <remarks>
        /// The key status returned from this function changes as a thread reads key messages from its message queue. The status does not reflect the interrupt-level state associated with the hardware. Use the GetAsyncKeyState function to retrieve that information. 
        /// An application calls GetKeyState in response to a keyboard-input message. This function retrieves the state of the key when the input message was generated. 
        /// To retrieve state information for all the virtual keys, use the GetKeyboardState function. 
        /// An application can use the virtual-key code constants VK_SHIFT, VK_CONTROL, and VK_MENU as values for Bthe nVirtKey parameter. This gives the status of the SHIFT, CTRL, or ALT keys without distinguishing between left and right. An application can also use the following virtual-key code constants as values for nVirtKey to distinguish between the left and right instances of those keys. 
        /// VK_LSHIFT
        /// VK_RSHIFT
        /// VK_LCONTROL
        /// VK_RCONTROL
        /// VK_LMENU
        /// VK_RMENU
        /// These left- and right-distinguishing constants are available to an application only through the GetKeyboardState, SetKeyboardState, GetAsyncKeyState, GetKeyState, and MapVirtualKey functions. 
        /// </remarks>
        public bool IsKeyUp(VirtualKeyCode keyCode) => !this.IsKeyDown(keyCode);

        /// <inheritdoc />
        /// <summary>
        /// Determines whether the physical key is up or down at the time the function is called regardless of whether the application thread has read the keyboard event from the message pump by calling the <see cref="NativeMethods.GetAsyncKeyState(ushort)" /> function. (See: http://msdn.microsoft.com/en-us/library/ms646293(VS.85).aspx)
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode" /> for the key.</param>
        /// <returns><c>true</c> if the key is down; otherwise, <c>false</c>.</returns>
        /// <remarks>
        /// The GetAsyncKeyState function works with mouse buttons. However, it checks on the state of the physical mouse buttons, not on the logical mouse buttons that the physical buttons are mapped to. For example, the call GetAsyncKeyState(VK_LBUTTON) always returns the state of the left physical mouse button, regardless of whether it is mapped to the left or right logical mouse button. You can determine the system's current mapping of physical mouse buttons to logical mouse buttons by calling 
        /// Copy CodeGetSystemMetrics(SM_SWAPBUTTON) which returns TRUE if the mouse buttons have been swapped.
        /// Although the least significant bit of the return value indicates whether the key has been pressed since the last query, due to the preemptive multitasking nature of Windows, another application can call GetAsyncKeyState and receive the "recently pressed" bit instead of your application. The behavior of the least significant bit of the return value is retained strictly for compatibility with 16-bit Windows applications (which are non-preemptive) and should not be relied upon.
        /// You can use the virtual-key code constants VK_SHIFT, VK_CONTROL, and VK_MENU as values for the vKey parameter. This gives the state of the SHIFT, CTRL, or ALT keys without distinguishing between left and right. 
        /// Windows NT/2000/XP: You can use the following virtual-key code constants as values for vKey to distinguish between the left and right instances of those keys. 
        /// Code Meaning 
        /// VK_LSHIFT Left-shift key. 
        /// VK_RSHIFT Right-shift key. 
        /// VK_LCONTROL Left-control key. 
        /// VK_RCONTROL Right-control key. 
        /// VK_LMENU Left-menu key. 
        /// VK_RMENU Right-menu key. 
        /// These left- and right-distinguishing constants are only available when you call the GetKeyboardState, SetKeyboardState, GetAsyncKeyState, GetKeyState, and MapVirtualKey functions. 
        /// </remarks>
        public bool IsHardwareKeyDown(VirtualKeyCode keyCode)
        {
            var result = NativeMethods.GetAsyncKeyState((ushort)keyCode);
            return result < 0;
        }

        /// <inheritdoc />
        /// <summary>
        /// Determines whether the physical key is up or down at the time the function is called regardless of whether the application thread has read the keyboard event from the message pump by calling the <see cref="NativeMethods.GetAsyncKeyState(ushort)" /> function. (See: http://msdn.microsoft.com/en-us/library/ms646293(VS.85).aspx)
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode" /> for the key.</param>
        /// <returns><c>true</c> if the key is up; otherwise, <c>false</c>.</returns>
        /// <remarks>
        /// The GetAsyncKeyState function works with mouse buttons. However, it checks on the state of the physical mouse buttons, not on the logical mouse buttons that the physical buttons are mapped to. For example, the call GetAsyncKeyState(VK_LBUTTON) always returns the state of the left physical mouse button, regardless of whether it is mapped to the left or right logical mouse button. You can determine the system's current mapping of physical mouse buttons to logical mouse buttons by calling 
        /// Copy CodeGetSystemMetrics(SM_SWAPBUTTON) which returns TRUE if the mouse buttons have been swapped.
        /// Although the least significant bit of the return value indicates whether the key has been pressed since the last query, due to the preemptive multitasking nature of Windows, another application can call GetAsyncKeyState and receive the "recently pressed" bit instead of your application. The behavior of the least significant bit of the return value is retained strictly for compatibility with 16-bit Windows applications (which are non-preemptive) and should not be relied upon.
        /// You can use the virtual-key code constants VK_SHIFT, VK_CONTROL, and VK_MENU as values for the vKey parameter. This gives the state of the SHIFT, CTRL, or ALT keys without distinguishing between left and right. 
        /// Windows NT/2000/XP: You can use the following virtual-key code constants as values for vKey to distinguish between the left and right instances of those keys. 
        /// Code Meaning 
        /// VK_LSHIFT Left-shift key. 
        /// VK_RSHIFT Right-shift key. 
        /// VK_LCONTROL Left-control key. 
        /// VK_RCONTROL Right-control key. 
        /// VK_LMENU Left-menu key. 
        /// VK_RMENU Right-menu key. 
        /// These left- and right-distinguishing constants are only available when you call the GetKeyboardState, SetKeyboardState, GetAsyncKeyState, GetKeyState, and MapVirtualKey functions. 
        /// </remarks>
        public bool IsHardwareKeyUp(VirtualKeyCode keyCode) => !this.IsHardwareKeyDown(keyCode);

        /// <inheritdoc />
        /// <summary>
        /// Determines whether the toggling key is toggled on (in-effect) or not by calling the <see cref="NativeMethods.GetKeyState(ushort)" /> function.  (See: http://msdn.microsoft.com/en-us/library/ms646301(VS.85).aspx)
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode" /> for the key.</param>
        /// <returns><c>true</c> if the toggling key is toggled on (in-effect); otherwise, <c>false</c>.</returns>
        /// <remarks>
        /// The key status returned from this function changes as a thread reads key messages from its message queue. The status does not reflect the interrupt-level state associated with the hardware. Use the GetAsyncKeyState function to retrieve that information. 
        /// An application calls GetKeyState in response to a keyboard-input message. This function retrieves the state of the key when the input message was generated. 
        /// To retrieve state information for all the virtual keys, use the GetKeyboardState function. 
        /// An application can use the virtual-key code constants VK_SHIFT, VK_CONTROL, and VK_MENU as values for the nVirtKey parameter. This gives the status of the SHIFT, CTRL, or ALT keys without distinguishing between left and right. An application can also use the following virtual-key code constants as values for nVirtKey to distinguish between the left and right instances of those keys. 
        /// VK_LSHIFT
        /// VK_RSHIFT
        /// VK_LCONTROL
        /// VK_RCONTROL
        /// VK_LMENU
        /// VK_RMENU
        /// These left- and right-distinguishing constants are available to an application only through the GetKeyboardState, SetKeyboardState, GetAsyncKeyState, GetKeyState, and MapVirtualKey functions. 
        /// </remarks>
        public bool IsTogglingKeyInEffect(VirtualKeyCode keyCode)
        {
            var result = NativeMethods.GetKeyState((ushort)keyCode);
            return (result & 0x01) == 0x01;
        }
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using System;
    using System.Runtime.InteropServices;

    using Native;

    /// <inheritdoc />
    /// <summary>
    /// Implements the <see cref="IInputMessageDispatcher" /> by calling <see cref="NativeMethods.SendInput(uint,Input[],int)" />.
    /// </summary>
    internal class WindowsInputMessageDispatcher : IInputMessageDispatcher
    {
        /// <inheritdoc />
        /// <summary>
        /// Dispatches the specified list of <see cref="Input" /> messages in their specified order by issuing a single called to <see cref="NativeMethods.SendInput(uint,Input[],int)" />.
        /// </summary>
        /// <param name="inputs">The list of <see cref="Input" /> messages to be dispatched.</param>
        /// <exception cref="ArgumentException">If the <paramref name="inputs" /> array is empty.</exception>
        /// <exception cref="ArgumentNullException">If the <paramref name="inputs" /> array is null.</exception>
        /// <exception cref="Exception">If the any of the commands in the <paramref name="inputs" /> array could not be sent successfully.</exception>
        public void DispatchInput(Input[] inputs)
        {
            if (inputs == null)
            {
                throw new ArgumentNullException(nameof(inputs));
            }

            if (inputs.Length == 0)
            {
                throw new ArgumentException("The input array was empty", nameof(inputs));
            }

            var successful = NativeMethods.SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(Input)));
            if (successful != inputs.Length)
            {
                throw new Exception("Some simulated input commands were not sent successfully. The most common reason for this happening are the security features of Windows including User Interface Privacy Isolation (UIPI). Your application can only send commands to applications of the same or lower elevation. Similarly certain commands are restricted to Accessibility/UIAutomation applications. Refer to the project home page and the code samples for more information.");
            }
        }
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using Native;

    /// <summary>
    /// The contract for a service that interprets the state of input devices.
    /// </summary>
    public interface IInputDeviceStateAdapter
    {
        /// <summary>
        /// Determines whether the specified key is up or down.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        /// <returns><c>true</c> if the key is down; otherwise, <c>false</c>.</returns>
        bool IsKeyDown(VirtualKeyCode keyCode);

        /// <summary>
        /// Determines whether the specified key is up or down.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        /// <returns><c>true</c> if the key is up; otherwise, <c>false</c>.</returns>
        bool IsKeyUp(VirtualKeyCode keyCode);

        /// <summary>
        /// Determines whether the physical key is up or down at the time the function is called regardless of whether the application thread has read the keyboard event from the message pump.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        /// <returns><c>true</c> if the key is down; otherwise, <c>false</c>.</returns>
        bool IsHardwareKeyDown(VirtualKeyCode keyCode);

        /// <summary>
        /// Determines whether the physical key is up or down at the time the function is called regardless of whether the application thread has read the keyboard event from the message pump.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        /// <returns><c>true</c> if the key is up; otherwise, <c>false</c>.</returns>
        bool IsHardwareKeyUp(VirtualKeyCode keyCode);

        /// <summary>
        /// Determines whether the toggling key is toggled on (in-effect) or not.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        /// <returns><c>true</c> if the toggling key is toggled on (in-effect); otherwise, <c>false</c>.</returns>
        bool IsTogglingKeyInEffect(VirtualKeyCode keyCode);
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using System;

    using Native;

    /// <summary>
    /// The contract for a service that dispatches <see cref="Input"/> messages to the appropriate destination.
    /// </summary>
    internal interface IInputMessageDispatcher
    {
        /// <summary>
        /// Dispatches the specified list of <see cref="Input"/> messages in their specified order.
        /// </summary>
        /// <param name="inputs">The list of <see cref="Input"/> messages to be dispatched.</param>
        /// <exception cref="ArgumentException">If the <paramref name="inputs"/> array is empty.</exception>
        /// <exception cref="ArgumentNullException">If the <paramref name="inputs"/> array is null.</exception>
        /// <exception cref="Exception">If the any of the commands in the <paramref name="inputs"/> array could not be sent successfully.</exception>
        void DispatchInput(Input[] inputs);
    }
}

namespace GregsStack.InputSimulatorStandard
{
    /// <summary>
    /// The contract for a service that simulates Keyboard and Mouse input and Hardware Input Device state detection for the Windows Platform.
    /// </summary>
    public interface IInputSimulator
    {
        /// <summary>
        /// Gets the <see cref="IKeyboardSimulator"/> instance for simulating Keyboard input.
        /// </summary>
        /// <value>The <see cref="IKeyboardSimulator"/> instance.</value>
        IKeyboardSimulator Keyboard { get; }

        /// <summary>
        /// Gets the <see cref="IMouseSimulator"/> instance for simulating Mouse input.
        /// </summary>
        /// <value>The <see cref="IMouseSimulator"/> instance.</value>
        IMouseSimulator Mouse { get; }

        /// <summary>
        /// Gets the <see cref="IInputDeviceStateAdapter"/> instance for determining the state of the various input devices.
        /// </summary>
        /// <value>The <see cref="IInputDeviceStateAdapter"/> instance.</value>
        IInputDeviceStateAdapter InputDeviceState { get; }
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using System;
    using System.Collections.Generic;

    using Native;

    /// <summary>
    /// The service contract for a keyboard simulator for the Windows platform.
    /// </summary>
    public interface IKeyboardSimulator
    {
        /// <summary>
        /// Simulates the key down gesture for the specified key.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        IKeyboardSimulator KeyDown(VirtualKeyCode keyCode);

        /// <summary>
        /// Simulates the key press gesture for the specified key.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        IKeyboardSimulator KeyPress(VirtualKeyCode keyCode);

        /// <summary>
        /// Simulates a key press for each of the specified key codes in the order they are specified.
        /// </summary>
        /// <param name="keyCodes"></param>
        IKeyboardSimulator KeyPress(params VirtualKeyCode[] keyCodes);

        /// <summary>
        /// Simulates the key up gesture for the specified key.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        IKeyboardSimulator KeyUp(VirtualKeyCode keyCode);

        /// <summary>
        /// Simulates a modified keystroke where there are multiple modifiers and multiple keys like CTRL-ALT-K-C where CTRL and ALT are the modifierKeys and K and C are the keys.
        /// The flow is Modifiers KeyDown in order, Keys Press in order, Modifiers KeyUp in reverse order.
        /// </summary>
        /// <param name="modifierKeyCodes">The list of <see cref="VirtualKeyCode"/>s for the modifier keys.</param>
        /// <param name="keyCodes">The list of <see cref="VirtualKeyCode"/>s for the keys to simulate.</param>
        IKeyboardSimulator ModifiedKeyStroke(IEnumerable<VirtualKeyCode> modifierKeyCodes, IEnumerable<VirtualKeyCode> keyCodes);

        /// <summary>
        /// Simulates a modified keystroke where there are multiple modifiers and one key like CTRL-ALT-C where CTRL and ALT are the modifierKeys and C is the key.
        /// The flow is Modifiers KeyDown in order, Key Press, Modifiers KeyUp in reverse order.
        /// </summary>
        /// <param name="modifierKeyCodes">The list of <see cref="VirtualKeyCode"/>s for the modifier keys.</param>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        IKeyboardSimulator ModifiedKeyStroke(IEnumerable<VirtualKeyCode> modifierKeyCodes, VirtualKeyCode keyCode);

        /// <summary>
        /// Simulates a modified keystroke where there is one modifier and multiple keys like CTRL-K-C where CTRL is the modifierKey and K and C are the keys.
        /// The flow is Modifier KeyDown, Keys Press in order, Modifier KeyUp.
        /// </summary>
        /// <param name="modifierKey">The <see cref="VirtualKeyCode"/> for the modifier key.</param>
        /// <param name="keyCodes">The list of <see cref="VirtualKeyCode"/>s for the keys to simulate.</param>
        IKeyboardSimulator ModifiedKeyStroke(VirtualKeyCode modifierKey, IEnumerable<VirtualKeyCode> keyCodes);

        /// <summary>
        /// Simulates a simple modified keystroke like CTRL-C where CTRL is the modifierKey and C is the key.
        /// The flow is Modifier KeyDown, Key Press, Modifier KeyUp.
        /// </summary>
        /// <param name="modifierKeyCode">The <see cref="VirtualKeyCode"/> for the  modifier key.</param>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/> for the key.</param>
        IKeyboardSimulator ModifiedKeyStroke(VirtualKeyCode modifierKeyCode, VirtualKeyCode keyCode);

        /// <summary>
        /// Simulates uninterrupted text entry via the keyboard.
        /// </summary>
        /// <param name="text">The text to be simulated.</param>
        IKeyboardSimulator TextEntry(string text);

        /// <summary>
        /// Simulates a single character text entry via the keyboard.
        /// </summary>
        /// <param name="character">The unicode character to be simulated.</param>
        IKeyboardSimulator TextEntry(char character);

        /// <summary>
        /// Sleeps the executing thread to create a pause between simulated inputs.
        /// </summary>
        /// <param name="millisecondsTimeout">The number of milliseconds to wait.</param>
        IKeyboardSimulator Sleep(int millisecondsTimeout);

        /// <summary>
        /// Sleeps the executing thread to create a pause between simulated inputs.
        /// </summary>
        /// <param name="timeout">The time to wait.</param>
        IKeyboardSimulator Sleep(TimeSpan timeout);
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using System;

    /// <summary>
    /// The service contract for a mouse simulator for the Windows platform.
    /// </summary>
    public interface IMouseSimulator
    {
        /// <summary>
        /// Gets or sets the amount of mouse wheel scrolling per click. The default value for this property is 120 and different values may cause some applications to interpret the scrolling differently than expected.
        /// </summary>
        int MouseWheelClickSize { get; set; }

        /// <summary>
        /// Retrieves the position of the mouse cursor, in screen coordinates.
        /// </summary>
        System.Drawing.Point Position { get; }

        /// <summary>
        /// Simulates mouse movement by the specified distance measured as a delta from the current mouse location in pixels.
        /// </summary>
        /// <param name="pixelDeltaX">The distance in pixels to move the mouse horizontally.</param>
        /// <param name="pixelDeltaY">The distance in pixels to move the mouse vertically.</param>
        IMouseSimulator MoveMouseBy(int pixelDeltaX, int pixelDeltaY);

        /// <summary>
        /// Simulates mouse movement to the specified location on the primary display device.
        /// </summary>
        /// <param name="absoluteX">The destination's absolute X-coordinate on the primary display device where 0 is the extreme left hand side of the display device and 65535 is the extreme right hand side of the display device.</param>
        /// <param name="absoluteY">The destination's absolute Y-coordinate on the primary display device where 0 is the top of the display device and 65535 is the bottom of the display device.</param>
        IMouseSimulator MoveMouseTo(double absoluteX, double absoluteY);

        /// <summary>
        /// Simulates mouse movement to the specified location on the Virtual Desktop which includes all active displays.
        /// </summary>
        /// <param name="absoluteX">The destination's absolute X-coordinate on the virtual desktop where 0 is the left hand side of the virtual desktop and 65535 is the extreme right hand side of the virtual desktop.</param>
        /// <param name="absoluteY">The destination's absolute Y-coordinate on the virtual desktop where 0 is the top of the virtual desktop and 65535 is the bottom of the virtual desktop.</param>
        IMouseSimulator MoveMouseToPositionOnVirtualDesktop(double absoluteX, double absoluteY);

        /// <summary>
        /// Simulates a mouse left button down gesture.
        /// </summary>
        IMouseSimulator LeftButtonDown();

        /// <summary>
        /// Simulates a mouse left button up gesture.
        /// </summary>
        IMouseSimulator LeftButtonUp();

        /// <summary>
        /// Simulates a mouse left button click gesture.
        /// </summary>
        IMouseSimulator LeftButtonClick();

        /// <summary>
        /// Simulates a mouse left button double-click gesture.
        /// </summary>
        IMouseSimulator LeftButtonDoubleClick();

        /// <summary>
        /// Simulates a mouse middle button down gesture.
        /// </summary>
        IMouseSimulator MiddleButtonDown();

        /// <summary>
        /// Simulates a mouse middle button up gesture.
        /// </summary>
        IMouseSimulator MiddleButtonUp();

        /// <summary>
        /// Simulates a mouse middle button click gesture.
        /// </summary>
        IMouseSimulator MiddleButtonClick();

        /// <summary>
        /// Simulates a mouse middle button double-click gesture.
        /// </summary>
        IMouseSimulator MiddleButtonDoubleClick();

        /// <summary>
        /// Simulates a mouse right button down gesture.
        /// </summary>
        IMouseSimulator RightButtonDown();

        /// <summary>
        /// Simulates a mouse right button up gesture.
        /// </summary>
        IMouseSimulator RightButtonUp();

        /// <summary>
        /// Simulates a mouse right button click gesture.
        /// </summary>
        IMouseSimulator RightButtonClick();

        /// <summary>
        /// Simulates a mouse right button double-click gesture.
        /// </summary>
        IMouseSimulator RightButtonDoubleClick();

        /// <summary>
        /// Simulates a mouse X button down gesture.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        IMouseSimulator XButtonDown(int buttonId);

        /// <summary>
        /// Simulates a mouse X button up gesture.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        IMouseSimulator XButtonUp(int buttonId);

        /// <summary>
        /// Simulates a mouse X button click gesture.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        IMouseSimulator XButtonClick(int buttonId);

        /// <summary>
        /// Simulates a mouse X button double-click gesture.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        IMouseSimulator XButtonDoubleClick(int buttonId);

        /// <summary>
        /// Simulates mouse vertical wheel scroll gesture.
        /// </summary>
        /// <param name="scrollAmountInClicks">The amount to scroll in clicks. A positive value indicates that the wheel was rotated forward, away from the user; a negative value indicates that the wheel was rotated backward, toward the user.</param>
        IMouseSimulator VerticalScroll(int scrollAmountInClicks);

        /// <summary>
        /// Simulates a mouse horizontal wheel scroll gesture. Supported by Windows Vista and later.
        /// </summary>
        /// <param name="scrollAmountInClicks">The amount to scroll in clicks. A positive value indicates that the wheel was rotated to the right; a negative value indicates that the wheel was rotated to the left.</param>
        IMouseSimulator HorizontalScroll(int scrollAmountInClicks);

        /// <summary>
        /// Sleeps the executing thread to create a pause between simulated inputs.
        /// </summary>
        /// <param name="millisecondsTimeout">The number of milliseconds to wait.</param>
        IMouseSimulator Sleep(int millisecondsTimeout);

        /// <summary>
        /// Sleeps the executing thread to create a pause between simulated inputs.
        /// </summary>
        /// <param name="timeout">The time to wait.</param>
        IMouseSimulator Sleep(TimeSpan timeout);
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using System;
    using System.Collections;
    using System.Collections.Generic;

    using Native;

    /// <inheritdoc />
    /// <summary>
    /// A helper class for building a list of <see cref="Input" /> messages ready to be sent to the native Windows API.
    /// </summary>
    internal class InputBuilder : IEnumerable<Input>
    {
        /// <summary>
        /// The public list of <see cref="Input"/> messages being built by this instance.
        /// </summary>
        private readonly List<Input> inputList;

        private static readonly List<VirtualKeyCode> ExtendedKeys = new List<VirtualKeyCode>
        {
            VirtualKeyCode.NUMPAD_RETURN,
            VirtualKeyCode.MENU,
            VirtualKeyCode.RMENU,
            VirtualKeyCode.CONTROL,
            VirtualKeyCode.RCONTROL,
            VirtualKeyCode.INSERT,
            VirtualKeyCode.DELETE,
            VirtualKeyCode.HOME,
            VirtualKeyCode.END,
            VirtualKeyCode.PRIOR,
            VirtualKeyCode.NEXT,
            VirtualKeyCode.RIGHT,
            VirtualKeyCode.UP,
            VirtualKeyCode.LEFT,
            VirtualKeyCode.DOWN,
            VirtualKeyCode.NUMLOCK,
            VirtualKeyCode.CANCEL,
            VirtualKeyCode.SNAPSHOT,
            VirtualKeyCode.DIVIDE
        };

        /// <summary>
        /// Initializes a new instance of the <see cref="InputBuilder"/> class.
        /// </summary>
        public InputBuilder()
        {
            this.inputList = new List<Input>();
        }

        /// <summary>
        /// Returns the list of <see cref="Input"/> messages as a <see cref="Array"/> of <see cref="Input"/> messages.
        /// </summary>
        /// <returns>The <see cref="System.Array"/> of <see cref="Input"/> messages.</returns>
        public Input[] ToArray() => this.inputList.ToArray();

        /// <inheritdoc />
        /// <summary>
        /// Returns an enumerator that iterates through the list of <see cref="Input" /> messages.
        /// </summary>
        /// <returns>
        /// A <see cref="IEnumerator" /> that can be used to iterate through the list of <see cref="Input" /> messages.
        /// </returns>
        /// <filterpriority>1</filterpriority>
        public IEnumerator<Input> GetEnumerator() => this.inputList.GetEnumerator();

        /// <inheritdoc />
        /// <summary>
        /// Returns an enumerator that iterates through the list of <see cref="Input" /> messages.
        /// </summary>
        /// <returns>
        /// An <see cref="IEnumerator" /> object that can be used to iterate through the list of <see cref="Input" /> messages.
        /// </returns>
        /// <filterpriority>2</filterpriority>
        IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

        /// <summary>
        /// Gets the <see cref="Input"/> at the specified position.
        /// </summary>
        /// <value>The <see cref="Input"/> message at the specified position.</value>
        public Input this[int position] => this.inputList[position];

        /// <summary>
        /// Determines if the <see cref="VirtualKeyCode"/> is an ExtendedKey
        /// </summary>
        /// <param name="keyCode">The key code.</param>
        /// <returns>true if the key code is an extended key; otherwise, false.</returns>
        /// <remarks>
        /// The extended keys consist of the ALT and CTRL keys on the right-hand side of the keyboard; the INS, DEL, HOME, END, PAGE UP, PAGE DOWN, and arrow keys in the clusters to the left of the numeric keypad; the NUM LOCK key; the BREAK (CTRL+PAUSE) key; the PRINT SCRN key; and the divide (/) and ENTER keys in the numeric keypad.
        ///
        /// See http://msdn.microsoft.com/en-us/library/ms646267(v=vs.85).aspx Section "Extended-Key Flag"
        /// </remarks>
        public static bool IsExtendedKey(VirtualKeyCode keyCode) => ExtendedKeys.Contains(keyCode);

        /// <summary>
        /// Adds a key down to the list of <see cref="Input"/> messages.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/>.</param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddKeyDown(VirtualKeyCode keyCode)
        {
            var code = (ushort)((int)keyCode & 0xFFFF);
            var down =
                new Input
                {
                    Type = (uint)InputType.Keyboard,
                    Data =
                            {
                                Keyboard =
                                    new KeyboardInput
                                        {
                                            KeyCode = (ushort) code ,
                                            Scan = (ushort)(NativeMethods.MapVirtualKey((uint)code, 0) & 0xFFU),
                                            Flags = IsExtendedKey(keyCode) ? (uint) KeyboardFlag.ExtendedKey : 0,
                                            Time = 0,
                                            ExtraInfo = IntPtr.Zero
                                        }
                            }
                };

            this.inputList.Add(down);
            return this;
        }

        /// <summary>
        /// Adds a key up to the list of <see cref="Input"/> messages.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/>.</param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddKeyUp(VirtualKeyCode keyCode)
        {
            var code = (ushort)((int)keyCode & 0xFFFF);
            var up =
                new Input
                {
                    Type = (uint)InputType.Keyboard,
                    Data =
                            {
                                Keyboard =
                                    new KeyboardInput
                                        {
                                            KeyCode = (ushort) code,
                                            Scan = (ushort)(NativeMethods.MapVirtualKey((uint)code, 0) & 0xFFU),
                                            Flags = (uint) (IsExtendedKey(keyCode)
                                                                  ? KeyboardFlag.KeyUp | KeyboardFlag.ExtendedKey
                                                                  : KeyboardFlag.KeyUp),
                                            Time = 0,
                                            ExtraInfo = IntPtr.Zero
                                        }
                            }
                };

            this.inputList.Add(up);
            return this;
        }

        /// <summary>
        /// Adds a key press to the list of <see cref="Input"/> messages which is equivalent to a key down followed by a key up.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode"/>.</param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddKeyPress(VirtualKeyCode keyCode)
        {
            this.AddKeyDown(keyCode);
            this.AddKeyUp(keyCode);
            return this;
        }

        /// <summary>
        /// Adds the character to the list of <see cref="Input"/> messages.
        /// </summary>
        /// <param name="character">The <see cref="System.Char"/> to be added to the list of <see cref="Input"/> messages.</param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddCharacter(char character)
        {
            ushort scanCode = character;

            var down = new Input
            {
                Type = (uint)InputType.Keyboard,
                Data =
                                   {
                                       Keyboard =
                                           new KeyboardInput
                                               {
                                                   KeyCode = 0,
                                                   Scan = scanCode,
                                                   Flags = (uint)KeyboardFlag.Unicode,
                                                   Time = 0,
                                                   ExtraInfo = IntPtr.Zero
                                               }
                                   }
            };

            var up = new Input
            {
                Type = (uint)InputType.Keyboard,
                Data =
                                 {
                                     Keyboard =
                                         new KeyboardInput
                                             {
                                                 KeyCode = 0,
                                                 Scan = scanCode,
                                                 Flags =
                                                     (uint)(KeyboardFlag.KeyUp | KeyboardFlag.Unicode),
                                                 Time = 0,
                                                 ExtraInfo = IntPtr.Zero
                                             }
                                 }
            };

            // Handle extended keys:
            // If the scan code is preceded by a prefix byte that has the value 0xE0 (224),
            // we need to include the KEYEVENTF_EXTENDEDKEY flag in the Flags property.
            if ((scanCode & 0xFF00) == 0xE000)
            {
                down.Data.Keyboard.Flags |= (uint)KeyboardFlag.ExtendedKey;
                up.Data.Keyboard.Flags |= (uint)KeyboardFlag.ExtendedKey;
            }

            this.inputList.Add(down);
            this.inputList.Add(up);
            return this;
        }

        /// <summary>
        /// Adds all of the characters in the specified <see cref="IEnumerable{T}"/> of <see cref="char"/>.
        /// </summary>
        /// <param name="characters">The characters to add.</param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddCharacters(IEnumerable<char> characters)
        {
            foreach (var character in characters)
            {
                this.AddCharacter(character);
            }

            return this;
        }

        /// <summary>
        /// Adds the characters in the specified <see cref="string"/>.
        /// </summary>
        /// <param name="characters">The string of <see cref="char"/> to add.</param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddCharacters(string characters) => this.AddCharacters(characters.ToCharArray());

        /// <summary>
        /// Moves the mouse relative to its current position.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddRelativeMouseMovement(int x, int y)
        {
            var movement = new Input { Type = (uint)InputType.Mouse };
            movement.Data.Mouse.Flags = (uint)MouseFlag.Move;
            movement.Data.Mouse.X = x;
            movement.Data.Mouse.Y = y;

            this.inputList.Add(movement);

            return this;
        }

        /// <summary>
        /// Move the mouse to an absolute position.
        /// </summary>
        /// <param name="absoluteX"></param>
        /// <param name="absoluteY"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddAbsoluteMouseMovement(int absoluteX, int absoluteY)
        {
            var movement = new Input { Type = (uint)InputType.Mouse };
            movement.Data.Mouse.Flags = (uint)(MouseFlag.Move | MouseFlag.Absolute);
            movement.Data.Mouse.X = (int)((absoluteX * 65536) / NativeMethods.GetSystemMetrics((uint)SystemMetric.SM_CXSCREEN));
            movement.Data.Mouse.Y = (int)((absoluteY * 65536) / NativeMethods.GetSystemMetrics((uint)SystemMetric.SM_CYSCREEN));

            this.inputList.Add(movement);

            return this;
        }

        /// <summary>
        /// Move the mouse to the absolute position on the virtual desktop.
        /// </summary>
        /// <param name="absoluteX"></param>
        /// <param name="absoluteY"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddAbsoluteMouseMovementOnVirtualDesktop(int absoluteX, int absoluteY)
        {
            var movement = new Input { Type = (uint)InputType.Mouse };
            movement.Data.Mouse.Flags = (uint)(MouseFlag.Move | MouseFlag.Absolute | MouseFlag.VirtualDesk);
            movement.Data.Mouse.X = absoluteX;
            movement.Data.Mouse.Y = absoluteY;

            this.inputList.Add(movement);

            return this;
        }

        /// <summary>
        /// Adds a mouse button down for the specified button.
        /// </summary>
        /// <param name="button"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseButtonDown(MouseButton button)
        {
            var buttonDown = new Input { Type = (uint)InputType.Mouse };
            buttonDown.Data.Mouse.Flags = (uint)ToMouseButtonDownFlag(button);

            this.inputList.Add(buttonDown);

            return this;
        }

        /// <summary>
        /// Adds a mouse button down for the specified button.
        /// </summary>
        /// <param name="xButtonId"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseXButtonDown(int xButtonId)
        {
            var buttonDown = new Input { Type = (uint)InputType.Mouse };
            buttonDown.Data.Mouse.Flags = (uint)MouseFlag.XDown;
            buttonDown.Data.Mouse.MouseData = (uint)xButtonId;

            this.inputList.Add(buttonDown);

            return this;
        }

        /// <summary>
        /// Adds a mouse button up for the specified button.
        /// </summary>
        /// <param name="button"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseButtonUp(MouseButton button)
        {
            var buttonUp = new Input { Type = (uint)InputType.Mouse };
            buttonUp.Data.Mouse.Flags = (uint)ToMouseButtonUpFlag(button);

            this.inputList.Add(buttonUp);

            return this;
        }

        /// <summary>
        /// Adds a mouse button up for the specified button.
        /// </summary>
        /// <param name="xButtonId"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseXButtonUp(int xButtonId)
        {
            var buttonUp = new Input { Type = (uint)InputType.Mouse };
            buttonUp.Data.Mouse.Flags = (uint)MouseFlag.XUp;
            buttonUp.Data.Mouse.MouseData = (uint)xButtonId;

            this.inputList.Add(buttonUp);

            return this;
        }

        /// <summary>
        /// Adds a single click of the specified button.
        /// </summary>
        /// <param name="button"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseButtonClick(MouseButton button) => this.AddMouseButtonDown(button).AddMouseButtonUp(button);

        /// <summary>
        /// Adds a single click of the specified button.
        /// </summary>
        /// <param name="xButtonId"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseXButtonClick(int xButtonId) => this.AddMouseXButtonDown(xButtonId).AddMouseXButtonUp(xButtonId);

        /// <summary>
        /// Adds a double click of the specified button.
        /// </summary>
        /// <param name="button"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseButtonDoubleClick(MouseButton button) => this.AddMouseButtonClick(button).AddMouseButtonClick(button);

        /// <summary>
        /// Adds a double click of the specified button.
        /// </summary>
        /// <param name="xButtonId"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseXButtonDoubleClick(int xButtonId) => this.AddMouseXButtonClick(xButtonId).AddMouseXButtonClick(xButtonId);

        /// <summary>
        /// Scroll the vertical mouse wheel by the specified amount.
        /// </summary>
        /// <param name="scrollAmount"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseVerticalWheelScroll(int scrollAmount)
        {
            var scroll = new Input { Type = (uint)InputType.Mouse };
            scroll.Data.Mouse.Flags = (uint)MouseFlag.VerticalWheel;
            scroll.Data.Mouse.MouseData = (uint)scrollAmount;

            this.inputList.Add(scroll);

            return this;
        }

        /// <summary>
        /// Scroll the horizontal mouse wheel by the specified amount.
        /// </summary>
        /// <param name="scrollAmount"></param>
        /// <returns>This <see cref="InputBuilder"/> instance.</returns>
        public InputBuilder AddMouseHorizontalWheelScroll(int scrollAmount)
        {
            var scroll = new Input { Type = (uint)InputType.Mouse };
            scroll.Data.Mouse.Flags = (uint)MouseFlag.HorizontalWheel;
            scroll.Data.Mouse.MouseData = (uint)scrollAmount;

            this.inputList.Add(scroll);

            return this;
        }

        private static MouseFlag ToMouseButtonDownFlag(MouseButton button)
        {
            switch (button)
            {
                case MouseButton.LeftButton:
                    return MouseFlag.LeftDown;

                case MouseButton.MiddleButton:
                    return MouseFlag.MiddleDown;

                case MouseButton.RightButton:
                    return MouseFlag.RightDown;

                default:
                    return MouseFlag.LeftDown;
            }
        }

        private static MouseFlag ToMouseButtonUpFlag(MouseButton button)
        {
            switch (button)
            {
                case MouseButton.LeftButton:
                    return MouseFlag.LeftUp;

                case MouseButton.MiddleButton:
                    return MouseFlag.MiddleUp;

                case MouseButton.RightButton:
                    return MouseFlag.RightUp;

                default:
                    return MouseFlag.LeftUp;
            }
        }
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using System;

    /// <inheritdoc />
    /// <summary>
    /// Implements the <see cref="IInputSimulator" /> interface to simulate Keyboard and Mouse input and provide the state of those input devices.
    /// </summary>
    public class InputSimulator : IInputSimulator
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InputSimulator"/> class using the default <see cref="KeyboardSimulator"/>, <see cref="MouseSimulator"/> and <see cref="WindowsInputDeviceStateAdapter"/> instances.
        /// </summary>
        public InputSimulator()
            : this(new KeyboardSimulator(), new MouseSimulator(), new WindowsInputDeviceStateAdapter())
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="InputSimulator"/> class using the specified <see cref="IKeyboardSimulator"/>, <see cref="IMouseSimulator"/> and <see cref="IInputDeviceStateAdapter"/> instances.
        /// </summary>
        /// <param name="keyboardSimulator">The <see cref="IKeyboardSimulator"/> instance to use for simulating keyboard input.</param>
        /// <param name="mouseSimulator">The <see cref="IMouseSimulator"/> instance to use for simulating mouse input.</param>
        /// <param name="inputDeviceStateAdapter">The <see cref="IInputDeviceStateAdapter"/> instance to use for interpreting the state of input devices.</param>
        public InputSimulator(IKeyboardSimulator keyboardSimulator, IMouseSimulator mouseSimulator, IInputDeviceStateAdapter inputDeviceStateAdapter)
        {
            this.Keyboard = keyboardSimulator ?? throw new ArgumentNullException(nameof(keyboardSimulator));
            this.Mouse = mouseSimulator ?? throw new ArgumentNullException(nameof(mouseSimulator));
            this.InputDeviceState = inputDeviceStateAdapter ?? throw new ArgumentNullException(nameof(inputDeviceStateAdapter));
        }

        /// <inheritdoc />
        /// <summary>
        /// Gets the <see cref="IKeyboardSimulator" /> instance for simulating Keyboard input.
        /// </summary>
        /// <value>The <see cref="IKeyboardSimulator" /> instance.</value>
        public IKeyboardSimulator Keyboard { get; }

        /// <inheritdoc />
        /// <summary>
        /// Gets the <see cref="IMouseSimulator" /> instance for simulating Mouse input.
        /// </summary>
        /// <value>The <see cref="IMouseSimulator" /> instance.</value>
        public IMouseSimulator Mouse { get; }

        /// <inheritdoc />
        /// <summary>
        /// Gets the <see cref="IInputDeviceStateAdapter" /> instance for determining the state of the various input devices.
        /// </summary>
        /// <value>The <see cref="IInputDeviceStateAdapter" /> instance.</value>
        public IInputDeviceStateAdapter InputDeviceState { get; }
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using System;
    using System.Collections.Generic;
    using System.Threading;

    using Native;

    /// <inheritdoc />
    /// <summary>
    /// Implements the <see cref="IKeyboardSimulator" /> interface by calling the an <see cref="IInputMessageDispatcher" /> to simulate Keyboard gestures.
    /// </summary>
    public class KeyboardSimulator : IKeyboardSimulator
    {
        /// <summary>
        /// The instance of the <see cref="IInputMessageDispatcher"/> to use for dispatching <see cref="Input"/> messages.
        /// </summary>
        private readonly IInputMessageDispatcher messageDispatcher;

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the <see cref="T:InputSimulatorStandard.KeyboardSimulator" /> class using an instance of a <see cref="T:InputSimulatorStandard.WindowsInputMessageDispatcher" /> for dispatching <see cref="T:InputSimulatorStandard.Native.Input" /> messages.
        /// </summary>
        public KeyboardSimulator()
            : this(new WindowsInputMessageDispatcher())
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="KeyboardSimulator"/> class using the specified <see cref="IInputMessageDispatcher"/> for dispatching <see cref="Input"/> messages.
        /// </summary>
        /// <param name="messageDispatcher">The <see cref="IInputMessageDispatcher"/> to use for dispatching <see cref="Input"/> messages.</param>
        /// <exception cref="InvalidOperationException">If null is passed as the <paramref name="messageDispatcher"/>.</exception>
        internal KeyboardSimulator(IInputMessageDispatcher messageDispatcher)
        {
            this.messageDispatcher = messageDispatcher ?? throw new InvalidOperationException(
                                         string.Format("The {0} cannot operate with a null {1}. Please provide a valid {1} instance to use for dispatching {2} messages.",
                                             typeof(KeyboardSimulator).Name, typeof(IInputMessageDispatcher).Name, typeof(Input).Name));
        }

        /// <inheritdoc />
        /// <summary>
        /// Calls the Win32 SendInput method to simulate a KeyDown.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode" /> to press</param>
        public IKeyboardSimulator KeyDown(VirtualKeyCode keyCode)
        {
            var inputList = new InputBuilder().AddKeyDown(keyCode).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Calls the Win32 SendInput method to simulate a KeyUp.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode" /> to lift up</param>
        public IKeyboardSimulator KeyUp(VirtualKeyCode keyCode)
        {
            var inputList = new InputBuilder().AddKeyUp(keyCode).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Calls the Win32 SendInput method with a KeyDown and KeyUp message in the same input sequence in order to simulate a Key PRESS.
        /// </summary>
        /// <param name="keyCode">The <see cref="VirtualKeyCode" /> to press</param>
        public IKeyboardSimulator KeyPress(VirtualKeyCode keyCode)
        {
            var inputList = new InputBuilder().AddKeyPress(keyCode).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a key press for each of the specified key codes in the order they are specified.
        /// </summary>
        /// <param name="keyCodes"></param>
        public IKeyboardSimulator KeyPress(params VirtualKeyCode[] keyCodes)
        {
            var builder = new InputBuilder();
            this.KeysPress(builder, keyCodes);
            this.SendSimulatedInput(builder.ToArray());
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a simple modified keystroke like CTRL-C where CTRL is the modifierKey and C is the key.
        /// The flow is Modifier KeyDown, Key Press, Modifier KeyUp.
        /// </summary>
        /// <param name="modifierKeyCode">The modifier key</param>
        /// <param name="keyCode">The key to simulate</param>
        public IKeyboardSimulator ModifiedKeyStroke(VirtualKeyCode modifierKeyCode, VirtualKeyCode keyCode)
        {
            this.ModifiedKeyStroke(new[] { modifierKeyCode }, new[] { keyCode });
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a modified keystroke where there are multiple modifiers and one key like CTRL-ALT-C where CTRL and ALT are the modifierKeys and C is the key.
        /// The flow is Modifiers KeyDown in order, Key Press, Modifiers KeyUp in reverse order.
        /// </summary>
        /// <param name="modifierKeyCodes">The list of modifier keys</param>
        /// <param name="keyCode">The key to simulate</param>
        public IKeyboardSimulator ModifiedKeyStroke(IEnumerable<VirtualKeyCode> modifierKeyCodes, VirtualKeyCode keyCode)
        {
            this.ModifiedKeyStroke(modifierKeyCodes, new[] { keyCode });
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a modified keystroke where there is one modifier and multiple keys like CTRL-K-C where CTRL is the modifierKey and K and C are the keys.
        /// The flow is Modifier KeyDown, Keys Press in order, Modifier KeyUp.
        /// </summary>
        /// <param name="modifierKey">The modifier key</param>
        /// <param name="keyCodes">The list of keys to simulate</param>
        public IKeyboardSimulator ModifiedKeyStroke(VirtualKeyCode modifierKey, IEnumerable<VirtualKeyCode> keyCodes)
        {
            this.ModifiedKeyStroke(new[] { modifierKey }, keyCodes);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a modified keystroke where there are multiple modifiers and multiple keys like CTRL-ALT-K-C where CTRL and ALT are the modifierKeys and K and C are the keys.
        /// The flow is Modifiers KeyDown in order, Keys Press in order, Modifiers KeyUp in reverse order.
        /// </summary>
        /// <param name="modifierKeyCodes">The list of modifier keys</param>
        /// <param name="keyCodes">The list of keys to simulate</param>
        public IKeyboardSimulator ModifiedKeyStroke(IEnumerable<VirtualKeyCode> modifierKeyCodes, IEnumerable<VirtualKeyCode> keyCodes)
        {
            var builder = new InputBuilder();
            this.ModifiersDown(builder, modifierKeyCodes);
            this.KeysPress(builder, keyCodes);
            this.ModifiersUp(builder, modifierKeyCodes);

            this.SendSimulatedInput(builder.ToArray());
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Calls the Win32 SendInput method with a stream of KeyDown and KeyUp messages in order to simulate uninterrupted text entry via the keyboard.
        /// </summary>
        /// <param name="text">The text to be simulated.</param>
        public IKeyboardSimulator TextEntry(string text)
        {
            if (text.Length > uint.MaxValue / 2)
            {
                throw new ArgumentException($"The text parameter is too long. It must be less than {uint.MaxValue / 2} characters.", nameof(text));
            }

            var inputList = new InputBuilder().AddCharacters(text).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a single character text entry via the keyboard.
        /// </summary>
        /// <param name="character">The unicode character to be simulated.</param>
        public IKeyboardSimulator TextEntry(char character)
        {
            var inputList = new InputBuilder().AddCharacter(character).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Sleeps the executing thread to create a pause between simulated inputs.
        /// </summary>
        /// <param name="millisecondsTimeout">The number of milliseconds to wait.</param>
        public IKeyboardSimulator Sleep(int millisecondsTimeout)
        {
            return this.Sleep(TimeSpan.FromMilliseconds(millisecondsTimeout));
        }

        /// <inheritdoc />
        /// <summary>
        /// Sleeps the executing thread to create a pause between simulated inputs.
        /// </summary>
        /// <param name="timeout">The time to wait.</param>
        public IKeyboardSimulator Sleep(TimeSpan timeout)
        {
            Thread.Sleep(timeout);
            return this;
        }

        private void ModifiersDown(InputBuilder builder, IEnumerable<VirtualKeyCode> modifierKeyCodes)
        {
            if (modifierKeyCodes == null)
            {
                return;
            }

            foreach (var key in modifierKeyCodes)
            {
                builder.AddKeyDown(key);
            }
        }

        private void ModifiersUp(InputBuilder builder, IEnumerable<VirtualKeyCode> modifierKeyCodes)
        {
            if (modifierKeyCodes == null)
            {
                return;
            }

            // Key up in reverse (I miss LINQ)
            var stack = new Stack<VirtualKeyCode>(modifierKeyCodes);
            while (stack.Count > 0)
            {
                builder.AddKeyUp(stack.Pop());
            }
        }

        private void KeysPress(InputBuilder builder, IEnumerable<VirtualKeyCode> keyCodes)
        {
            if (keyCodes == null)
            {
                return;
            }

            foreach (var key in keyCodes)
            {
                builder.AddKeyPress(key);
            }
        }

        /// <summary>
        /// Sends the list of <see cref="Input"/> messages using the <see cref="IInputMessageDispatcher"/> instance.
        /// </summary>
        /// <param name="inputList">The <see cref="Array"/> of <see cref="Input"/> messages to send.</param>
        private void SendSimulatedInput(Input[] inputList)
        {
            this.messageDispatcher.DispatchInput(inputList);
        }
    }
}

namespace GregsStack.InputSimulatorStandard
{
    /// <summary>
    /// The mouse button
    /// </summary>
    public enum MouseButton
    {
        /// <summary>
        /// Left mouse button
        /// </summary>
        LeftButton,

        /// <summary>
        /// Middle mouse button
        /// </summary>
        MiddleButton,

        /// <summary>
        /// Right mouse button
        /// </summary>
        RightButton,
    }
}

namespace GregsStack.InputSimulatorStandard
{
    using System;
    using System.Threading;

    using Native;

    /// <inheritdoc />
    /// <summary>
    /// Implements the <see cref="IMouseSimulator" /> interface by calling the an <see cref="IInputMessageDispatcher" /> to simulate Mouse gestures.
    /// </summary>
    public class MouseSimulator : IMouseSimulator
    {
        /// <summary>
        /// The instance of the <see cref="IInputMessageDispatcher"/> to use for dispatching <see cref="Input"/> messages.
        /// </summary>
        private readonly IInputMessageDispatcher messageDispatcher;

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the <see cref="T:InputSimulatorStandard.MouseSimulator" /> class using an instance of a <see cref="T:InputSimulatorStandard.WindowsInputMessageDispatcher" /> for dispatching <see cref="T:InputSimulatorStandard.Native.Input" /> messages.
        /// </summary>
        public MouseSimulator()
            : this(new WindowsInputMessageDispatcher())
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MouseSimulator"/> class using the specified <see cref="IInputMessageDispatcher"/> for dispatching <see cref="Input"/> messages.
        /// </summary>
        /// <param name="messageDispatcher">The <see cref="IInputMessageDispatcher"/> to use for dispatching <see cref="Input"/> messages.</param>
        /// <exception cref="InvalidOperationException">If null is passed as the <paramref name="messageDispatcher"/>.</exception>
        internal MouseSimulator(IInputMessageDispatcher messageDispatcher)
        {
            this.messageDispatcher = messageDispatcher ?? throw new InvalidOperationException(
                                         string.Format("The {0} cannot operate with a null {1}. Please provide a valid {1} instance to use for dispatching {2} messages.",
                                             typeof(MouseSimulator).Name, typeof(IInputMessageDispatcher).Name, typeof(Input).Name));
        }

        /// <inheritdoc />
        /// <summary>
        /// Gets or sets the amount of mouse wheel scrolling per click. The default value for this property is 120 and different values may cause some applications to interpret the scrolling differently than expected.
        /// </summary>
        public int MouseWheelClickSize { get; set; } = 120;

        /// <inheritdoc />
        /// <summary>
        /// Retrieves the position of the mouse cursor, in screen coordinates.
        /// </summary>
        public System.Drawing.Point Position
        {
            get
            {
                NativeMethods.GetCursorPos(out Point lpPoint);
                return lpPoint;
            }
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates mouse movement by the specified distance measured as a delta from the current mouse location in pixels.
        /// </summary>
        /// <param name="pixelDeltaX">The distance in pixels to move the mouse horizontally.</param>
        /// <param name="pixelDeltaY">The distance in pixels to move the mouse vertically.</param>
        public IMouseSimulator MoveMouseBy(int pixelDeltaX, int pixelDeltaY)
        {
            var inputList = new InputBuilder().AddRelativeMouseMovement(pixelDeltaX, pixelDeltaY).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates mouse movement to the specified location on the primary display device.
        /// </summary>
        /// <param name="absoluteX">The destination's absolute X-coordinate on the primary display device where 0 is the extreme left hand side of the display device and 65535 is the extreme right hand side of the display device.</param>
        /// <param name="absoluteY">The destination's absolute Y-coordinate on the primary display device where 0 is the top of the display device and 65535 is the bottom of the display device.</param>
        public IMouseSimulator MoveMouseTo(double absoluteX, double absoluteY)
        {
            var inputList = new InputBuilder().AddAbsoluteMouseMovement((int)Math.Truncate(absoluteX), (int)Math.Truncate(absoluteY)).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates mouse movement to the specified location on the Virtual Desktop which includes all active displays.
        /// </summary>
        /// <param name="absoluteX">The destination's absolute X-coordinate on the virtual desktop where 0 is the left hand side of the virtual desktop and 65535 is the extreme right hand side of the virtual desktop.</param>
        /// <param name="absoluteY">The destination's absolute Y-coordinate on the virtual desktop where 0 is the top of the virtual desktop and 65535 is the bottom of the virtual desktop.</param>
        public IMouseSimulator MoveMouseToPositionOnVirtualDesktop(double absoluteX, double absoluteY)
        {
            var inputList = new InputBuilder().AddAbsoluteMouseMovementOnVirtualDesktop((int)Math.Truncate(absoluteX), (int)Math.Truncate(absoluteY)).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse left button down gesture.
        /// </summary>
        public IMouseSimulator LeftButtonDown() => this.ButtonDown(MouseButton.LeftButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse left button up gesture.
        /// </summary>
        public IMouseSimulator LeftButtonUp() => this.ButtonUp(MouseButton.LeftButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse left-click gesture.
        /// </summary>
        public IMouseSimulator LeftButtonClick() => this.ButtonClick(MouseButton.LeftButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse left button double-click gesture.
        /// </summary>
        public IMouseSimulator LeftButtonDoubleClick() => this.ButtonDoubleClick(MouseButton.LeftButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse middle button down gesture.
        /// </summary>
        public IMouseSimulator MiddleButtonDown() => this.ButtonDown(MouseButton.MiddleButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse middle button up gesture.
        /// </summary>
        public IMouseSimulator MiddleButtonUp() => this.ButtonUp(MouseButton.MiddleButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse middle button click gesture.
        /// </summary>
        public IMouseSimulator MiddleButtonClick() => this.ButtonClick(MouseButton.MiddleButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse middle button double click gesture.
        /// </summary>
        public IMouseSimulator MiddleButtonDoubleClick() => this.ButtonDoubleClick(MouseButton.MiddleButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse right button down gesture.
        /// </summary>
        public IMouseSimulator RightButtonDown() => this.ButtonDown(MouseButton.RightButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse right button up gesture.
        /// </summary>
        public IMouseSimulator RightButtonUp() => this.ButtonUp(MouseButton.RightButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse right button click gesture.
        /// </summary>
        public IMouseSimulator RightButtonClick() => this.ButtonClick(MouseButton.RightButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse right button double-click gesture.
        /// </summary>
        public IMouseSimulator RightButtonDoubleClick() => this.ButtonDoubleClick(MouseButton.RightButton);

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse X button down gesture.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        public IMouseSimulator XButtonDown(int buttonId)
        {
            var inputList = new InputBuilder().AddMouseXButtonDown(buttonId).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse X button up gesture.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        public IMouseSimulator XButtonUp(int buttonId)
        {
            var inputList = new InputBuilder().AddMouseXButtonUp(buttonId).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse X button click gesture.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        public IMouseSimulator XButtonClick(int buttonId)
        {
            var inputList = new InputBuilder().AddMouseXButtonClick(buttonId).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse X button double-click gesture.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        public IMouseSimulator XButtonDoubleClick(int buttonId)
        {
            var inputList = new InputBuilder().AddMouseXButtonDoubleClick(buttonId).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates mouse vertical wheel scroll gesture.
        /// </summary>
        /// <param name="scrollAmountInClicks">The amount to scroll in clicks. A positive value indicates that the wheel was rotated forward, away from the user; a negative value indicates that the wheel was rotated backward, toward the user.</param>
        public IMouseSimulator VerticalScroll(int scrollAmountInClicks)
        {
            var inputList = new InputBuilder().AddMouseVerticalWheelScroll(scrollAmountInClicks * this.MouseWheelClickSize).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Simulates a mouse horizontal wheel scroll gesture. Supported by Windows Vista and later.
        /// </summary>
        /// <param name="scrollAmountInClicks">The amount to scroll in clicks. A positive value indicates that the wheel was rotated to the right; a negative value indicates that the wheel was rotated to the left.</param>
        public IMouseSimulator HorizontalScroll(int scrollAmountInClicks)
        {
            var inputList = new InputBuilder().AddMouseHorizontalWheelScroll(scrollAmountInClicks * this.MouseWheelClickSize).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        /// <inheritdoc />
        /// <summary>
        /// Sleeps the executing thread to create a pause between simulated inputs.
        /// </summary>
        /// <param name="millisecondsTimeout">The number of milliseconds to wait.</param>
        public IMouseSimulator Sleep(int millisecondsTimeout)
        {
            return this.Sleep(TimeSpan.FromMilliseconds(millisecondsTimeout));
        }

        /// <inheritdoc />
        /// <summary>
        /// Sleeps the executing thread to create a pause between simulated inputs.
        /// </summary>
        /// <param name="timeout">The time to wait.</param>
        public IMouseSimulator Sleep(TimeSpan timeout)
        {
            Thread.Sleep(timeout);
            return this;
        }

        /// <summary>
        /// Sends the list of <see cref="Input"/> messages using the <see cref="IInputMessageDispatcher"/> instance.
        /// </summary>
        /// <param name="inputList">The <see cref="System.Array"/> of <see cref="Input"/> messages to send.</param>
        private void SendSimulatedInput(Input[] inputList)
        {
            this.messageDispatcher.DispatchInput(inputList);
        }

        private IMouseSimulator ButtonDown(MouseButton mouseButton)
        {
            var inputList = new InputBuilder().AddMouseButtonDown(mouseButton).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        private IMouseSimulator ButtonUp(MouseButton mouseButton)
        {
            var inputList = new InputBuilder().AddMouseButtonUp(mouseButton).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        private IMouseSimulator ButtonClick(MouseButton mouseButton)
        {
            var inputList = new InputBuilder().AddMouseButtonClick(mouseButton).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }

        private IMouseSimulator ButtonDoubleClick(MouseButton mouseButton)
        {
            var inputList = new InputBuilder().AddMouseButtonDoubleClick(mouseButton).ToArray();
            this.SendSimulatedInput(inputList);
            return this;
        }
    }
}