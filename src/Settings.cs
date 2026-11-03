using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Input;
using Microsoft.Win32;

namespace Fanak
{
    // Stored next to the exe so the settings travel with the folder. If the folder is read-only
    // (Program Files, for example), they go to the user's profile instead.
    internal sealed class AppSettings
    {
        private const string FileName = "settings.ini";

        public static readonly int[] DefaultHotkey = { 0xC0 };
        public const int DefaultDpiValue = 300;

        public int[] HotkeyKeys = DefaultHotkey;
        public int DefaultDpi = DefaultDpiValue;
        public bool NotifyOnDone = true;
        public bool IsFirstRun;

        private static string PortablePath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, FileName); }
        }

        private static string ProfilePath
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Fanak", FileName); }
        }

        public static AppSettings Load()
        {
            var settings = new AppSettings();
            string path = File.Exists(PortablePath) ? PortablePath : File.Exists(ProfilePath) ? ProfilePath : null;
            if (path == null)
            {
                settings.IsFirstRun = true;
                return settings;
            }

            try
            {
                foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                {
                    int split = line.IndexOf('=');
                    if (split > 0) settings.Apply(line.Substring(0, split).Trim(), line.Substring(split + 1).Trim());
                }
            }
            catch (IOException ex)
            {
                Log.Write("설정 읽기 실패: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                Log.Write("설정 읽기 실패: " + ex.Message);
            }
            return settings;
        }

        private void Apply(string key, string value)
        {
            int number;
            if (key == "hotkey")
            {
                var keys = new List<int>();
                foreach (string part in value.Split('+'))
                {
                    if (!int.TryParse(part, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out number) || number <= 0 || number > 0xFE) return;
                    keys.Add(number);
                }
                if (keys.Count > 0) HotkeyKeys = keys.ToArray();
            }
            else if (key == "default_dpi")
            {
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out number) && number >= Converter.MinDpi && number <= Converter.MaxDpi)
                    DefaultDpi = number;
            }
            else if (key == "notify_on_done")
            {
                NotifyOnDone = value != "0";
            }
        }

        public void Save()
        {
            var text = new StringBuilder();
            text.Append("hotkey=").Append(string.Join("+", HotkeyKeys.Select(k => k.ToString("X2", CultureInfo.InvariantCulture)))).Append("\r\n");
            text.Append("default_dpi=").Append(DefaultDpi.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
            text.Append("notify_on_done=").Append(NotifyOnDone ? "1" : "0").Append("\r\n");

            if (TryWrite(PortablePath, text.ToString())) return;
            Directory.CreateDirectory(Path.GetDirectoryName(ProfilePath));
            if (!TryWrite(ProfilePath, text.ToString())) Log.Write("설정 저장 실패: " + ProfilePath);
        }

        private static bool TryWrite(string path, string text)
        {
            try
            {
                File.WriteAllText(path, text, new UTF8Encoding(false));
                return true;
            }
            catch (IOException) { return false; }
            catch (UnauthorizedAccessException) { return false; }
        }
    }

    internal static class AutoStart
    {
        private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string ValueName = "FanakConvert";

        private static string ExeCommand
        {
            get { return "\"" + Assembly.GetEntryAssembly().Location + "\""; }
        }

        public static bool IsEnabled
        {
            get
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKey))
                {
                    return key != null && key.GetValue(ValueName) != null;
                }
            }
        }

        public static void Set(bool enabled)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RunKey))
            {
                if (enabled) key.SetValue(ValueName, ExeCommand);
                else key.DeleteValue(ValueName, false);
            }
        }

        // If the folder was moved, the registered path would otherwise point nowhere.
        public static void RefreshPath()
        {
            if (!IsEnabled) return;
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKey))
            {
                if ((string)key.GetValue(ValueName) == ExeCommand) return;
            }
            Set(true);
        }
    }

    internal static class HotkeyInfo
    {
        private const int Shift = 0x10;
        private const int Ctrl = 0x11;
        private const int Alt = 0x12;

        private static readonly Dictionary<int, string> Names = new Dictionary<int, string>
        {
            { Shift, "Shift" }, { Ctrl, "Ctrl" }, { Alt, "Alt" },
            { 0x09, "Tab" }, { 0x0D, "Enter" }, { 0x20, "Space" }, { 0x08, "Backspace" },
            { 0x21, "Page Up" }, { 0x22, "Page Down" }, { 0x23, "End" }, { 0x24, "Home" },
            { 0x25, "←" }, { 0x26, "↑" }, { 0x27, "→" }, { 0x28, "↓" }, { 0x2D, "Insert" }, { 0x2E, "Delete" },
            { 0xBA, ";" }, { 0xBB, "=" }, { 0xBC, "," }, { 0xBD, "-" }, { 0xBE, "." }, { 0xBF, "/" },
            { 0xC0, "`" }, { 0xDB, "[" }, { 0xDC, "\\" }, { 0xDD, "]" }, { 0xDE, "'" }
        };

        // Esc cancels the drag itself, Windows keys open Start, and the rest flip a mode that stays on.
        private static readonly HashSet<Key> Rejected = new HashSet<Key>
        {
            Key.Escape, Key.LWin, Key.RWin, Key.Apps, Key.CapsLock, Key.NumLock, Key.Scroll,
            Key.HangulMode, Key.HanjaMode, Key.PrintScreen, Key.Pause, Key.Sleep
        };

        public static bool IsRejected(Key key)
        {
            return Rejected.Contains(key);
        }

        public static int ToVirtualKey(Key key)
        {
            int vk = KeyInterop.VirtualKeyFromKey(key);
            switch (vk)
            {
                case 0xA0:
                case 0xA1: return Shift;
                case 0xA2:
                case 0xA3: return Ctrl;
                case 0xA4:
                case 0xA5: return Alt;
                default: return vk;
            }
        }

        public static int[] Sorted(IEnumerable<int> keys)
        {
            return keys.OrderBy(k => k == Ctrl ? 0 : k == Shift ? 1 : k == Alt ? 2 : 3).ThenBy(k => k).ToArray();
        }

        public static string Label(IEnumerable<int> keys)
        {
            return string.Join(" + ", keys.Select(NameOf));
        }

        private static string NameOf(int vk)
        {
            string name;
            if (Names.TryGetValue(vk, out name)) return name;
            if ((vk >= 0x30 && vk <= 0x39) || (vk >= 0x41 && vk <= 0x5A)) return ((char)vk).ToString();
            if (vk >= 0x70 && vk <= 0x87) return "F" + (vk - 0x6F);
            return KeyInterop.KeyFromVirtualKey(vk).ToString();
        }

        public static string WarningFor(int[] keys)
        {
            if (keys.All(k => k == Shift || k == Ctrl || k == Alt))
                return "Shift, Ctrl, Alt만 쓰면 탐색기의 복사·이동 드래그와 겹칩니다.";
            bool typing = keys.Any(k => k == 0x20 || k == 0x09 || k == 0x0D || (k >= 0x30 && k <= 0x5A));
            bool hasModifier = keys.Any(k => k == Ctrl || k == Alt);
            if (typing && !hasModifier)
                return "드래그 전에 누르면 입력 중인 곳에 글자가 들어갈 수 있습니다.";
            return null;
        }
    }
}
