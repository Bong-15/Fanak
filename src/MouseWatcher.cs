using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Threading;

namespace Fanak
{
    internal static class Native
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        public delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc proc, IntPtr hMod, uint threadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string name);

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int key);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hwnd, StringBuilder name, int capacity);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorPos(out POINT point);

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool TerminateProcess(IntPtr process, uint exitCode);

        // The plain versions exist on both 32-bit and 64-bit Windows and are enough for style flags.
        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hwnd, int index);

        [DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hwnd, int index, int value);
    }

    // Watches for a left-button drag that starts in File Explorer or on the desktop while the hotkey is held.
    internal sealed class MouseWatcher : IDisposable
    {
        private const int WH_MOUSE_LL = 14;
        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const uint GA_ROOT = 2;
        private const int DragThreshold = 10;
        private const int KeyPollMilliseconds = 25;

        private readonly DispatcherTimer _keyTimer;
        private readonly Native.LowLevelMouseProc _proc;
        private IntPtr _hook;
        private bool _down;
        private bool _eligible;
        private bool _fired;
        private Native.POINT _start;

        // Replaced as a whole when the setting changes, so reading it from the hook needs no lock.
        public int[] Hotkey { get; set; }

        public event Action HotkeyDragStarted;
        public event Action ButtonReleased;

        public MouseWatcher()
        {
            // No mouse events arrive while the mouse rests, and the key may be tapped right then.
            _keyTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(KeyPollMilliseconds) };
            _keyTimer.Tick += (s, e) => CheckHotkey();
            _proc = OnMouse;
            _hook = Native.SetWindowsHookEx(WH_MOUSE_LL, _proc, Native.GetModuleHandle(null), 0);
            if (_hook == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        private IntPtr OnMouse(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int message = wParam.ToInt32();
                if (message == WM_LBUTTONDOWN)
                {
                    _start = (Native.POINT)Marshal.PtrToStructure(lParam, typeof(Native.POINT));
                    _down = true;
                    _fired = false;
                    _eligible = IsFileManagerAt(_start);
                }
                else if (message == WM_MOUSEMOVE && _down && _eligible && !_fired && !_keyTimer.IsEnabled)
                {
                    var point = (Native.POINT)Marshal.PtrToStructure(lParam, typeof(Native.POINT));
                    if (Math.Abs(point.X - _start.X) >= DragThreshold || Math.Abs(point.Y - _start.Y) >= DragThreshold)
                    {
                        _keyTimer.Start();
                        CheckHotkey();
                    }
                }
                else if (message == WM_LBUTTONUP && _down)
                {
                    _down = false;
                    _keyTimer.Stop();
                    var handler = ButtonReleased;
                    if (handler != null) handler();
                }
            }
            return Native.CallNextHookEx(_hook, nCode, wParam, lParam);
        }

        private void CheckHotkey()
        {
            if (_fired || !IsHotkeyDown()) return;
            _fired = true;
            _keyTimer.Stop();
            var handler = HotkeyDragStarted;
            if (handler != null) handler();
        }

        private bool IsHotkeyDown()
        {
            int[] keys = Hotkey;
            if (keys == null || keys.Length == 0) return false;
            foreach (int key in keys)
            {
                if ((Native.GetAsyncKeyState(key) & 0x8000) == 0) return false;
            }
            return true;
        }

        private static bool IsFileManagerAt(Native.POINT point)
        {
            IntPtr root = Native.GetAncestor(Native.WindowFromPoint(point), GA_ROOT);
            if (root == IntPtr.Zero) return false;
            var name = new StringBuilder(64);
            Native.GetClassName(root, name, name.Capacity);
            string cls = name.ToString();
            return cls == "CabinetWClass" || cls == "ExploreWClass" || cls == "Progman" || cls == "WorkerW";
        }

        public void Dispose()
        {
            _keyTimer.Stop();
            if (_hook == IntPtr.Zero) return;
            Native.UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }
    }
}
