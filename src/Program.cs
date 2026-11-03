using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Drawing = System.Drawing;
using WinForms = System.Windows.Forms;

namespace Fanak
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            // Fanak.exe --convert PNG [--dpi 300] [--resample] [--crop x,y,w,h] file1 file2 ...
            if (args.Length >= 3 && args[0] == "--convert") return ExitNow(RunCli(args));

            // Fanak.exe --quit closes the running instance.
            if (args.Length == 1 && args[0] == "--quit")
            {
                Signal(TrayApp.QuitEventName);
                return 0;
            }

            bool isFirst;
            using (var mutex = new Mutex(true, @"Local\Fanak.SingleInstance", out isFirst))
            {
                if (!isFirst)
                {
                    // Opening the exe again is the easiest way back to a tray app, so show its settings.
                    Signal(TrayApp.ShowSettingsEventName);
                    return 0;
                }
                var app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
                Theme.Apply(app);
                using (new TrayApp(app))
                {
                    app.Run();
                }
                GC.KeepAlive(mutex);
            }
            return ExitNow(0);
        }

        private static void Signal(string eventName)
        {
            using (var signal = new EventWaitHandle(false, EventResetMode.AutoReset, eventName)) signal.Set();
        }

        // After rendering a PDF, the Windows PDF engine never finishes unloading and a normal exit spins
        // forever. Every file is closed by this point, so nothing is lost by skipping that step.
        private static int ExitNow(int code)
        {
            OfficePdf.StopActive();
            Native.TerminateProcess(Native.GetCurrentProcess(), (uint)code);
            return code;
        }

        private static int RunCli(string[] args)
        {
            var request = new ConvertRequest { Format = args[1] };
            var files = new List<string>();
            for (int i = 2; i < args.Length; i++)
            {
                if (args[i] == "--dpi" && i + 1 < args.Length) request.Dpi = int.Parse(args[++i], CultureInfo.InvariantCulture);
                else if (args[i] == "--resample") request.Resample = true;
                else if (args[i] == "--crop" && i + 1 < args.Length)
                {
                    int[] v = args[++i].Split(',').Select(p => int.Parse(p, CultureInfo.InvariantCulture)).ToArray();
                    request.Crop = new Int32Rect(v[0], v[1], v[2], v[3]);
                }
                else files.Add(args[i]);
            }
            request.Files = files.ToArray();

            int failures = 0;
            try
            {
                foreach (string output in Converter.Run(request, (file, ex) =>
                {
                    failures++;
                    Log.Write("변환 실패: " + file + " - " + ex.Message);
                }))
                {
                    Log.Write("변환 완료: " + output);
                }
            }
            catch (ArgumentException ex)
            {
                failures++;
                Log.Write("변환 실패: " + ex.Message);
            }
            return failures == 0 ? 0 : 1;
        }
    }

    // Conversions run one at a time on their own thread, so the mouse hook on the UI thread never lags.
    internal sealed class ConversionQueue
    {
        private readonly BlockingCollection<ConvertRequest> _jobs = new BlockingCollection<ConvertRequest>();
        private readonly Thread _thread;
        private volatile bool _stopping;

        public event Action<ConvertRequest, List<string>, int, string> BatchFinished;

        public ConversionQueue()
        {
            _thread = new Thread(Work) { IsBackground = true, Name = "Fanak conversions" };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        public void Enqueue(ConvertRequest request)
        {
            _jobs.Add(request);
        }

        private void Work()
        {
            foreach (ConvertRequest job in _jobs.GetConsumingEnumerable())
            {
                if (_stopping) return;
                int failed = 0;
                string firstError = null;
                var outputs = new List<string>();
                try
                {
                    outputs = Converter.Run(job, (file, ex) =>
                    {
                        failed++;
                        if (firstError == null) firstError = ex.Message;
                        Log.Write("변환 실패: " + file + " - " + ex.Message);
                    });
                }
                catch (Exception ex)
                {
                    failed++;
                    firstError = ex.Message;
                    Log.Write("변환 실패: " + ex.Message);
                }
                outputs.ForEach(output => Log.Write("변환 완료: " + output));

                var handler = BatchFinished;
                if (handler != null && !_stopping) handler(job, outputs, failed, firstError);
            }
        }

        public void Stop()
        {
            _stopping = true;
            _jobs.CompleteAdding();
            _thread.Join(TimeSpan.FromSeconds(3));
        }
    }

    internal sealed class TrayApp : IDisposable
    {
        public const string QuitEventName = @"Local\Fanak.Quit";
        public const string ShowSettingsEventName = @"Local\Fanak.ShowSettings";
        private const int MaxTrayText = 63;
        private const int MaxBalloonText = 250;

        private readonly Application _app;
        private readonly Dispatcher _dispatcher;
        private readonly AppSettings _settings;
        private readonly MouseWatcher _watcher;
        private readonly RadialWindow _window;
        private readonly ConversionQueue _queue;
        private readonly WinForms.NotifyIcon _tray;
        private readonly DispatcherTimer _hideTimer;
        private readonly StatusWindow _status;
        // What the work since the queue was last empty has produced, reported together when it empties again.
        private readonly List<string> _outputs = new List<string>();
        private CancellationTokenSource _cancel = new CancellationTokenSource();
        private SettingsWindow _settingsWindow;
        private string _firstError;
        private string _note;
        private int _failed;
        private int _pending;

        public TrayApp(Application app)
        {
            _app = app;
            _dispatcher = Dispatcher.CurrentDispatcher;
            _settings = AppSettings.Load();

            _window = new RadialWindow();
            // Leaves the drop callback first, because Explorer stays blocked until it returns.
            _window.FormatChosen += (files, format) =>
                _dispatcher.BeginInvoke(new Action(() => OnFormatChosen(files, format)));

            _status = new StatusWindow(CancelAll, OpenLog);
            _queue = new ConversionQueue();
            _queue.BatchFinished += (job, outputs, failed, error) =>
                _dispatcher.BeginInvoke(new Action(() => OnBatchFinished(job, outputs, failed, error)));

            // The drop arrives just after the button is released, so hiding has to wait a moment.
            _hideTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _hideTimer.Tick += (s, e) =>
            {
                _hideTimer.Stop();
                _window.Dismiss();
            };

            var menu = new WinForms.ContextMenuStrip();
            menu.Items.Add("설정", null, (s, e) => ShowSettings(SettingsWindow.GeneralTab));
            menu.Items.Add("사용법", null, (s, e) => ShowSettings(SettingsWindow.GuideTab));
            menu.Items.Add("변환 기록 열기", null, (s, e) => OpenLog());
            menu.Items.Add(new WinForms.ToolStripSeparator());
            menu.Items.Add("종료", null, (s, e) => _app.Shutdown());
            _tray = new WinForms.NotifyIcon
            {
                Icon = LoadTrayIcon(),
                ContextMenuStrip = menu,
                Visible = true
            };
            _tray.DoubleClick += (s, e) => ShowSettings(SettingsWindow.GeneralTab);
            SetStatus("대기 중");

            _watcher = new MouseWatcher { Hotkey = _settings.HotkeyKeys };
            _watcher.HotkeyDragStarted += () => _dispatcher.BeginInvoke(new Action(() =>
            {
                _hideTimer.Stop();
                _window.ShowAtCursor();
            }));
            _watcher.ButtonReleased += () => _dispatcher.BeginInvoke(new Action(() =>
            {
                if (_window.IsVisible) _hideTimer.Start();
            }));

            WaitForSignal(QuitEventName, () => _app.Shutdown());
            WaitForSignal(ShowSettingsEventName, () => ShowSettings(SettingsWindow.GeneralTab));

            try
            {
                AutoStart.RefreshPath();
            }
            catch (Exception ex)
            {
                if (!(ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is IOException)) throw;
                Log.Write("자동 실행 경로 갱신 실패: " + ex.Message);
            }

            Log.Write("앱 시작");
            if (_settings.IsFirstRun)
            {
                _settings.Save();
                ShowSettings(SettingsWindow.GuideTab);
            }
            else if (_settings.NotifyOnDone)
            {
                ShowBalloon("Fanak 변환기 실행 중", "파일을 끌다가 " + HotkeyInfo.Label(_settings.HotkeyKeys) + " 키를 누르면 메뉴가 열립니다.", WinForms.ToolTipIcon.None);
            }
        }

        private void WaitForSignal(string eventName, Action action)
        {
            var signal = new EventWaitHandle(false, EventResetMode.AutoReset, eventName);
            ThreadPool.RegisterWaitForSingleObject(signal, (state, timedOut) =>
                _dispatcher.BeginInvoke(action), null, Timeout.Infinite, false);
        }

        private void ShowSettings(int tab)
        {
            if (_settingsWindow != null)
            {
                _settingsWindow.ShowTab(tab);
                _settingsWindow.Activate();
                return;
            }
            _settingsWindow = new SettingsWindow(_settings, tab,
                () => _watcher.Hotkey = _settings.HotkeyKeys, OpenLog, () => _app.Shutdown());
            _settingsWindow.Closed += (s, e) => _settingsWindow = null;
            _settingsWindow.Show();
        }

        private void OnFormatChosen(string[] files, string format)
        {
            _hideTimer.Stop();
            if (format == Converter.DpiAction)
            {
                new DpiWindow(files, _settings.DefaultDpi, Submit).Show();
            }
            else if (format == Converter.CropAction)
            {
                OpenCrop(files[0]);
            }
            else
            {
                Submit(new ConvertRequest { Files = files, Format = format, PdfDpi = _settings.DefaultDpi });
            }
        }

        private void Submit(ConvertRequest request)
        {
            string title = request.Format == Converter.KeepFormat ? "저장 중" : request.Format + "로 변환 중";
            if (_pending == 0)
            {
                _cancel.Dispose();
                _cancel = new CancellationTokenSource();
                _outputs.Clear();
                _firstError = null;
                _note = null;
                _failed = 0;
                _status.BeginWork(title, request.Files[0]);
            }
            _pending++;
            SetStatus("변환 중…");
            request.Cancel = _cancel.Token;
            request.FileStarted = (file, index, total) => _dispatcher.BeginInvoke(new Action(() =>
            {
                _status.SetTitle(title);
                _status.SetFile(file, index, total, _pending - 1);
            }));
            _queue.Enqueue(request);
        }

        private void CancelAll()
        {
            _cancel.Cancel();
            OfficePdf.StopActive();
        }

        // Large images decode off the UI thread so the mouse hook does not stall the whole system's pointer.
        private void OpenCrop(string file)
        {
            Task.Run(() => Converter.LoadUpright(file)).ContinueWith(task =>
            {
                if (task.IsFaulted)
                {
                    string message = task.Exception.GetBaseException().Message;
                    Log.Write("자르기 창 열기 실패: " + file + " - " + message);
                    _status.ShowFailed(0, 1, message, null);
                    return;
                }
                new CropWindow(file, task.Result, Submit).Show();
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void OnBatchFinished(ConvertRequest job, List<string> outputs, int failed, string error)
        {
            _pending--;
            _outputs.AddRange(outputs);
            _failed += failed;
            if (_firstError == null) _firstError = error;
            if (job.Dpi.HasValue) _note = job.Dpi.Value + " DPI로 저장된 것을 확인했습니다.";
            if (_pending > 0) return;

            SetStatus("대기 중");
            string last = _outputs.LastOrDefault();
            if (_failed > 0) _status.ShowFailed(_outputs.Count, _failed, _firstError, last);
            else if (_cancel.IsCancellationRequested) _status.ShowCancelled(_outputs.Count, last);
            else if (_outputs.Count > 0 && _settings.NotifyOnDone) _status.ShowDone(_outputs.ToList(), _note);
            else _status.HideCard();
        }

        private void ShowBalloon(string title, string text, WinForms.ToolTipIcon icon)
        {
            _tray.ShowBalloonTip(5000, title, text.Length > MaxBalloonText ? text.Substring(0, MaxBalloonText) : text, icon);
        }

        private void SetStatus(string status)
        {
            string text = "Fanak 변환기 - " + status;
            _tray.Text = text.Length > MaxTrayText ? text.Substring(0, MaxTrayText) : text;
        }

        private static void OpenLog()
        {
            if (File.Exists(Log.FilePath)) Process.Start("notepad.exe", "\"" + Log.FilePath + "\"");
        }

        private static Drawing.Icon LoadTrayIcon()
        {
            using (Stream stream = Theme.OpenIconStream())
            {
                return new Drawing.Icon(stream, WinForms.SystemInformation.SmallIconSize);
            }
        }

        public void Dispose()
        {
            _watcher.Dispose();
            _hideTimer.Stop();
            _tray.Visible = false;
            _tray.Dispose();
            _queue.Stop();
            Log.Write("앱 종료");
        }
    }
}
