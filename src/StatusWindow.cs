using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Threading;

namespace Fanak
{
    // The card in the screen corner that follows conversions from start to finish. Tray balloons can't hold buttons.
    // It never takes focus, because the user is usually busy in another window.
    internal sealed class StatusWindow : Window
    {
        private const double CardWidth = 360;
        private const double ShadowSpace = 14;
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        // A conversion that ends sooner than this needs no progress card, only the result.
        private static readonly TimeSpan ShowDelay = TimeSpan.FromMilliseconds(500);
        private static readonly TimeSpan DoneStays = TimeSpan.FromSeconds(8);
        private static readonly TimeSpan CancelledStays = TimeSpan.FromSeconds(4);

        private readonly Action _cancel;
        private readonly Action _openLog;
        private readonly TextBlock _title = new TextBlock { FontSize = 14, FontWeight = FontWeights.SemiBold, TextTrimming = TextTrimming.CharacterEllipsis };
        private readonly TextBlock _name = new TextBlock { Margin = new Thickness(0, 3, 0, 0), TextTrimming = TextTrimming.CharacterEllipsis };
        private readonly TextBlock _detail = new TextBlock { FontSize = 12, Margin = new Thickness(0, 3, 0, 0), TextWrapping = TextWrapping.Wrap, MaxHeight = 72, TextTrimming = TextTrimming.CharacterEllipsis };
        private readonly ProgressBar _bar = new ProgressBar { Height = 4, Margin = new Thickness(0, 10, 0, 0), BorderThickness = new Thickness(0) };
        private readonly StackPanel _buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        private readonly Button _closeButton;
        private readonly DispatcherTimer _clock = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        private readonly DispatcherTimer _showTimer = new DispatcherTimer { Interval = ShowDelay };
        private readonly DispatcherTimer _hideTimer = new DispatcherTimer();
        private readonly Stopwatch _elapsed = new Stopwatch();
        private string _progress = "";
        private bool _cancelling;

        public StatusWindow(Action cancel, Action openLog)
        {
            _cancel = cancel;
            _openLog = openLog;

            Title = "Fanak 변환기";
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            ShowInTaskbar = false;
            ShowActivated = false;
            Topmost = true;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.Manual;
            SizeToContent = SizeToContent.Height;
            Width = CardWidth + ShadowSpace * 2;
            FontFamily = new FontFamily("Segoe UI, Malgun Gothic");
            FontSize = 13;
            UseLayoutRounding = true;
            TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);

            var ink = (Brush)Application.Current.FindResource("Ink");
            var muted = (Brush)Application.Current.FindResource("Muted");
            _title.Foreground = ink;
            _name.Foreground = ink;
            _detail.Foreground = muted;
            _bar.Foreground = (Brush)Application.Current.FindResource("Accent");
            _bar.Background = (Brush)Application.Current.FindResource("Line");

            _closeButton = new Button { Content = "✕", MinWidth = 0, Width = 26, Height = 26, Padding = new Thickness(0), FontSize = 11, BorderBrush = Brushes.Transparent, Background = Brushes.Transparent, VerticalAlignment = VerticalAlignment.Top };
            _closeButton.Click += (s, e) => HideCard();

            var words = new StackPanel { Margin = new Thickness(12, 0, 6, 0) };
            words.Children.Add(_title);
            words.Children.Add(_name);
            words.Children.Add(_detail);

            var top = new Grid();
            top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            top.ColumnDefinitions.Add(new ColumnDefinition());
            top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            var icon = new Image { Source = Theme.AppIcon, Width = 32, Height = 32, VerticalAlignment = VerticalAlignment.Top };
            RenderOptions.SetBitmapScalingMode(icon, BitmapScalingMode.HighQuality);
            Grid.SetColumn(words, 1);
            Grid.SetColumn(_closeButton, 2);
            top.Children.Add(icon);
            top.Children.Add(words);
            top.Children.Add(_closeButton);

            var body = new StackPanel();
            body.Children.Add(top);
            body.Children.Add(_bar);
            body.Children.Add(_buttons);

            Content = new Border
            {
                Background = Brushes.White,
                BorderBrush = (Brush)Application.Current.FindResource("Line"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(14, 12, 12, 12),
                Margin = new Thickness(ShadowSpace),
                Effect = new DropShadowEffect { BlurRadius = 14, ShadowDepth = 2, Opacity = 0.22 },
                Child = body
            };

            _clock.Tick += (s, e) => ShowProgress();
            _showTimer.Tick += (s, e) =>
            {
                _showTimer.Stop();
                ShowCard();
            };
            _hideTimer.Tick += (s, e) => HideCard();
            // Someone reaching for a button shouldn't have it vanish under the pointer.
            MouseEnter += (s, e) => _hideTimer.Stop();
            MouseLeave += (s, e) =>
            {
                if (_hideTimer.Interval > TimeSpan.Zero && !_clock.IsEnabled && IsVisible) _hideTimer.Start();
            };
            SizeChanged += (s, e) => PlaceInCorner();
            SourceInitialized += (s, e) =>
            {
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                Native.SetWindowLong(hwnd, GWL_EXSTYLE, Native.GetWindowLong(hwnd, GWL_EXSTYLE) | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW);
            };
        }

        public void BeginWork(string title, string firstFile)
        {
            _hideTimer.Stop();
            _hideTimer.Interval = TimeSpan.Zero;
            _cancelling = false;
            _title.Text = title;
            SetName(firstFile);
            _progress = "";
            _bar.Visibility = Visibility.Visible;
            _bar.IsIndeterminate = true;
            _closeButton.Visibility = Visibility.Collapsed;
            SetButtons(new Button[] { MakeButton("취소", false, RequestCancel) });
            _elapsed.Restart();
            _clock.Start();
            ShowProgress();
            if (IsVisible) return;
            _showTimer.Start();
        }

        public void SetTitle(string title)
        {
            if (!_cancelling) _title.Text = title;
        }

        // Counted by file. Word and Hangul don't report how far along they are, so one file has no percentage to show.
        public void SetFile(string file, int index, int total, int jobsWaiting)
        {
            if (_cancelling) return;
            SetName(file);
            _progress = total > 1 ? index + " / " + total + "개" : "";
            if (jobsWaiting > 0) _progress += (_progress.Length > 0 ? " · " : "") + "대기 " + jobsWaiting + "건";
            _bar.IsIndeterminate = total <= 1;
            if (total > 1)
            {
                _bar.Maximum = total;
                _bar.Value = index - 1;
            }
            ShowProgress();
        }

        private void ShowProgress()
        {
            string clock = (int)_elapsed.Elapsed.TotalMinutes + ":" + _elapsed.Elapsed.Seconds.ToString("00");
            _detail.Text = _cancelling
                ? "진행 중인 파일을 멈추고 있습니다. 조금 걸릴 수 있습니다."
                : (_progress.Length > 0 ? _progress + " · " : "") + clock;
        }

        private void RequestCancel()
        {
            _cancelling = true;
            _title.Text = "취소하는 중";
            ShowProgress();
            foreach (object button in _buttons.Children) ((Button)button).IsEnabled = false;
            _cancel();
        }

        public void ShowDone(IList<string> outputs, string note)
        {
            string last = outputs[outputs.Count - 1];
            var buttons = new List<Button>();
            if (outputs.Count == 1) buttons.Add(MakeButton("파일 열기", true, () => OpenFile(last)));
            buttons.Add(MakeButton("폴더 열기", outputs.Count > 1, () => OpenFolder(last)));
            Finish("저장 완료", outputs.Count == 1 ? Path.GetFileName(last) : outputs.Count + "개 파일", note, buttons, DoneStays);
        }

        public void ShowFailed(int succeeded, int failed, string error, string lastOutput)
        {
            var buttons = new List<Button>();
            buttons.Add(MakeButton("변환 기록", false, _openLog));
            if (lastOutput != null) buttons.Add(MakeButton("폴더 열기", false, () => OpenFolder(lastOutput)));
            string summary = succeeded > 0 ? "성공 " + succeeded + "개, 실패 " + failed + "개" : failed > 1 ? failed + "개 파일" : "";
            // Nothing else records that this happened, so it waits to be read.
            Finish("변환 실패", summary, error, buttons, TimeSpan.Zero);
        }

        public void ShowCancelled(int kept, string lastOutput)
        {
            var buttons = new List<Button>();
            if (lastOutput != null) buttons.Add(MakeButton("폴더 열기", false, () => OpenFolder(lastOutput)));
            Finish("변환 취소", kept > 0 ? "이미 끝난 " + kept + "개는 남겨 두었습니다." : "저장된 파일은 없습니다.", null, buttons, CancelledStays);
        }

        public void HideCard()
        {
            _showTimer.Stop();
            _hideTimer.Stop();
            _clock.Stop();
            _elapsed.Stop();
            Hide();
        }

        private void Finish(string title, string name, string detail, IList<Button> buttons, TimeSpan stays)
        {
            _showTimer.Stop();
            _clock.Stop();
            _elapsed.Stop();
            _title.Text = title;
            SetName(name);
            _detail.Text = detail ?? "";
            _detail.Visibility = string.IsNullOrEmpty(detail) ? Visibility.Collapsed : Visibility.Visible;
            _bar.IsIndeterminate = false;
            _bar.Visibility = Visibility.Collapsed;
            _closeButton.Visibility = Visibility.Visible;
            SetButtons(buttons);
            ShowCard();

            _hideTimer.Stop();
            _hideTimer.Interval = stays;
            if (stays > TimeSpan.Zero && !IsMouseOver) _hideTimer.Start();
        }

        private void SetName(string name)
        {
            string shown = name != null && Path.IsPathRooted(name) ? Path.GetFileName(name) : name;
            _name.Text = shown ?? "";
            _name.ToolTip = string.IsNullOrEmpty(shown) ? null : shown;
            _name.Visibility = string.IsNullOrEmpty(shown) ? Visibility.Collapsed : Visibility.Visible;
            _detail.Visibility = Visibility.Visible;
        }

        private void SetButtons(IList<Button> buttons)
        {
            _buttons.Children.Clear();
            _buttons.Margin = new Thickness(0, buttons.Count > 0 ? 10 : 0, 0, 0);
            foreach (Button button in buttons) _buttons.Children.Add(button);
        }

        private Button MakeButton(string label, bool primary, Action action)
        {
            var button = new Button { Content = label, Margin = new Thickness(6, 0, 0, 0), Padding = new Thickness(12, 4, 12, 4) };
            if (primary) button.Style = Ui.Find("PrimaryButton");
            button.Click += (s, e) => action();
            return button;
        }

        private void OpenFile(string path)
        {
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                HideCard();
            }
            catch (Win32Exception ex)
            {
                Log.Write("파일 열기 실패: " + path + " - " + ex.Message);
                _detail.Text = "이 파일을 열 프로그램을 찾지 못했습니다.";
                _detail.Visibility = Visibility.Visible;
            }
        }

        private void OpenFolder(string path)
        {
            Process.Start("explorer.exe", "/select,\"" + path + "\"");
            HideCard();
        }

        private void ShowCard()
        {
            Show();
            PlaceInCorner();
        }

        private void PlaceInCorner()
        {
            Rect area = SystemParameters.WorkArea;
            Left = area.Right - ActualWidth;
            Top = area.Bottom - ActualHeight;
        }
    }
}
