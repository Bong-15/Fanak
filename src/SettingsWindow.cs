using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Fanak
{
    internal sealed class SettingsWindow : Window
    {
        public const int GeneralTab = 0;
        public const int GuideTab = 1;
        private const int MaxHotkeyKeys = 3;
        private const string Version = "261103 ver 1.0";

        private readonly AppSettings _settings;
        private readonly Action _changed;
        private readonly TabControl _tabs;
        private readonly TextBlock _keyLabel;
        private readonly TextBlock _keyHint;
        private readonly TextBlock _guideKey;
        private readonly Button _changeKey;
        private readonly ComboBox _defaultDpi;
        private readonly TextBlock _dpiHint;
        private readonly TextBlock _startHint;
        private readonly List<int> _held = new List<int>();
        private bool _capturing;

        public SettingsWindow(AppSettings settings, int tab, Action changed, Action openLog, Action quit)
        {
            _settings = settings;
            _changed = changed;

            Title = "Fanak 변환기 설정";
            Width = 500;
            SizeToContent = SizeToContent.Height;
            ResizeMode = ResizeMode.NoResize;
            Ui.Prepare(this);

            _keyLabel = new TextBlock { FontSize = 16, FontWeight = FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Center };
            _keyHint = Ui.Hint("");
            _guideKey = new TextBlock { FontWeight = FontWeights.SemiBold };
            _changeKey = new Button { Content = "변경", Margin = new Thickness(12, 0, 0, 0) };
            _changeKey.Click += (s, e) => BeginCapture();
            _defaultDpi = Ui.DpiBox(settings.DefaultDpi.ToString(CultureInfo.InvariantCulture), false);
            _dpiHint = Ui.Hint("DPI 창이 처음 열릴 때의 값이자, PDF를 이미지로 바꿀 때의 해상도입니다.");
            _startHint = Ui.Hint("켜 두면 Windows에 로그인할 때 자동으로 실행됩니다.");

            _tabs = new TabControl();
            _tabs.Items.Add(new TabItem { Header = "일반", Content = BuildGeneral() });
            _tabs.Items.Add(new TabItem { Header = "사용법", Content = BuildGuide() });
            _tabs.Items.Add(new TabItem { Header = "정보", Content = BuildAbout(openLog, quit) });
            _tabs.SelectedIndex = tab;

            var close = new Button { Content = "닫기", Style = Ui.Find("PrimaryButton"), IsCancel = true, HorizontalAlignment = HorizontalAlignment.Right };
            close.Click += (s, e) => Close();
            var bottom = new Border { BorderBrush = (Brush)Application.Current.FindResource("Line"), BorderThickness = new Thickness(0, 1, 0, 0), Padding = new Thickness(20, 12, 20, 12), Child = close };

            var layout = new DockPanel();
            DockPanel.SetDock(bottom, Dock.Bottom);
            layout.Children.Add(bottom);
            layout.Children.Add(_tabs);
            Content = layout;

            ShowHotkey();

            PreviewKeyDown += OnPreviewKeyDown;
            PreviewKeyUp += OnPreviewKeyUp;
            Deactivated += (s, e) => CancelCapture();
            Closing += (s, e) => CommitDefaultDpi();
        }

        public void ShowTab(int tab)
        {
            _tabs.SelectedIndex = tab;
        }

        // ---- General ----

        private UIElement BuildGeneral()
        {
            var page = new StackPanel { Margin = new Thickness(20, 18, 20, 8) };

            var keyRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 8) };
            keyRow.Children.Add(new Border { Style = Ui.Find("KeyChip"), Child = _keyLabel });
            keyRow.Children.Add(_changeKey);
            var reset = new Button { Content = "기본값으로", Margin = new Thickness(8, 0, 0, 0) };
            reset.Click += (s, e) => SetHotkey(AppSettings.DefaultHotkey);
            keyRow.Children.Add(reset);

            var key = new StackPanel();
            key.Children.Add(Ui.Heading("메뉴 호출 키"));
            key.Children.Add(keyRow);
            key.Children.Add(_keyHint);
            page.Children.Add(Ui.Card(key));

            var dpi = new StackPanel();
            dpi.Children.Add(Ui.Heading("기본 DPI"));
            _defaultDpi.Margin = new Thickness(0, 10, 0, 8);
            _defaultDpi.HorizontalAlignment = HorizontalAlignment.Left;
            _defaultDpi.LostKeyboardFocus += (s, e) => CommitDefaultDpi();
            _defaultDpi.SelectionChanged += (s, e) => Dispatcher.BeginInvoke(new Action(CommitDefaultDpi));
            dpi.Children.Add(_defaultDpi);
            dpi.Children.Add(_dpiHint);
            page.Children.Add(Ui.Card(dpi));

            var behavior = new StackPanel();
            behavior.Children.Add(Ui.Heading("앱 동작"));
            var notify = new CheckBox { Content = "변환이 끝나면 알림 표시", IsChecked = _settings.NotifyOnDone, Margin = new Thickness(0, 10, 0, 0) };
            notify.Checked += (s, e) => SetNotify(true);
            notify.Unchecked += (s, e) => SetNotify(false);
            behavior.Children.Add(notify);
            TextBlock notifyHint = Ui.Hint("변환에 실패하면 이 설정과 상관없이 항상 알려 줍니다.");
            notifyHint.Margin = new Thickness(0, 2, 0, 0);
            behavior.Children.Add(notifyHint);
            var start = new CheckBox { Content = "Windows 시작 시 자동 실행", IsChecked = AutoStart.IsEnabled, Margin = new Thickness(0, 12, 0, 0) };
            start.Checked += (s, e) => SetAutoStart(start, true);
            start.Unchecked += (s, e) => SetAutoStart(start, false);
            behavior.Children.Add(start);
            _startHint.Margin = new Thickness(0, 2, 0, 0);
            behavior.Children.Add(_startHint);
            page.Children.Add(Ui.Card(behavior));
            return page;
        }

        private void SetNotify(bool enabled)
        {
            _settings.NotifyOnDone = enabled;
            _settings.Save();
        }

        private void SetAutoStart(CheckBox box, bool enabled)
        {
            try
            {
                AutoStart.Set(enabled);
            }
            catch (Exception ex)
            {
                if (!(ex is UnauthorizedAccessException || ex is SecurityException || ex is System.IO.IOException)) throw;
                _startHint.Text = "자동 실행 설정을 바꾸지 못했습니다: " + ex.Message;
                _startHint.Foreground = Ui.ErrorBrush;
            }
        }

        private void CommitDefaultDpi()
        {
            int? dpi;
            if (Ui.TryReadDpi(_defaultDpi, out dpi) && dpi.HasValue)
            {
                _dpiHint.Foreground = (Brush)Application.Current.FindResource("Muted");
                _dpiHint.Text = "DPI 창이 처음 열릴 때의 값이자, PDF를 이미지로 바꿀 때의 해상도입니다.";
                if (dpi.Value == _settings.DefaultDpi) return;
                _settings.DefaultDpi = dpi.Value;
                _settings.Save();
                return;
            }
            _dpiHint.Text = Ui.DpiError + " 이전 값 " + _settings.DefaultDpi + "을 유지합니다.";
            _dpiHint.Foreground = Ui.ErrorBrush;
            _defaultDpi.Text = _settings.DefaultDpi.ToString(CultureInfo.InvariantCulture);
        }

        // ---- Hotkey capture ----

        private void BeginCapture()
        {
            CommitDefaultDpi();
            _capturing = true;
            _held.Clear();
            _keyLabel.Text = "…";
            _changeKey.IsEnabled = false;
            _keyHint.Foreground = (Brush)Application.Current.FindResource("Muted");
            _keyHint.Text = "원하는 키를 누르세요. 최대 3개까지 조합할 수 있고, Esc를 누르면 취소됩니다.";
            Keyboard.Focus(_tabs);
        }

        private void CancelCapture()
        {
            if (!_capturing) return;
            _capturing = false;
            _changeKey.IsEnabled = true;
            ShowHotkey();
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!_capturing) return;
            e.Handled = true;

            Key key = e.Key;
            if (key == Key.System) key = e.SystemKey;
            else if (key == Key.ImeProcessed) key = e.ImeProcessedKey;
            else if (key == Key.DeadCharProcessed) key = e.DeadCharProcessedKey;

            if (key == Key.Escape)
            {
                CancelCapture();
                return;
            }
            if (HotkeyInfo.IsRejected(key))
            {
                _keyHint.Text = key + " 키는 쓸 수 없습니다. 다른 키를 누르세요.";
                _keyHint.Foreground = Ui.ErrorBrush;
                return;
            }

            int vk = HotkeyInfo.ToVirtualKey(key);
            if (vk != 0 && !_held.Contains(vk) && _held.Count < MaxHotkeyKeys) _held.Add(vk);
            if (_held.Count > 0) _keyLabel.Text = HotkeyInfo.Label(HotkeyInfo.Sorted(_held));
        }

        private void OnPreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (!_capturing) return;
            e.Handled = true;
            if (_held.Count == 0) return;

            _capturing = false;
            _changeKey.IsEnabled = true;
            SetHotkey(HotkeyInfo.Sorted(_held));
            _changeKey.Focus();
        }

        private void SetHotkey(int[] keys)
        {
            _settings.HotkeyKeys = keys;
            _settings.Save();
            _changed();
            ShowHotkey();
        }

        private void ShowHotkey()
        {
            string label = HotkeyInfo.Label(_settings.HotkeyKeys);
            _keyLabel.Text = label;
            _guideKey.Text = label;

            string warning = HotkeyInfo.WarningFor(_settings.HotkeyKeys);
            _keyHint.Text = warning ?? "파일을 끄는 중에 이 키를 누르면 변환 메뉴가 열립니다.";
            _keyHint.Foreground = warning != null ? Ui.ErrorBrush : (Brush)Application.Current.FindResource("Muted");
        }

        // ---- Guide ----

        private UIElement BuildGuide()
        {
            var page = new StackPanel { Margin = new Thickness(20, 18, 20, 8) };

            var steps = new StackPanel();
            steps.Children.Add(Ui.Heading("이렇게 쓰세요"));
            steps.Children.Add(Step("1", "탐색기나 바탕 화면에서 파일을 끌기 시작합니다."));
            var keyStep = new StackPanel { Orientation = Orientation.Horizontal };
            keyStep.Children.Add(new TextBlock { Text = "끄는 중에 " });
            keyStep.Children.Add(_guideKey);
            keyStep.Children.Add(new TextBlock { Text = " 키를 누르면 메뉴가 열립니다." });
            steps.Children.Add(Step("2", keyStep));
            steps.Children.Add(Step("3", "원하는 항목 위에 놓습니다. 가운데에 놓으면 취소됩니다."));
            page.Children.Add(Ui.Card(steps));

            var features = new Grid();
            features.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(104) });
            features.ColumnDefinitions.Add(new ColumnDefinition());
            string[,] rows =
            {
                { "이미지 변환", "PNG, JPG, TIFF, BMP, GIF" },
                { "PDF 만들기", "이미지를 PDF로 저장합니다. 여러 장은 이름순으로 한 파일에 합칩니다." },
                { "PDF → 이미지", "페이지마다 이미지 파일 하나로 저장합니다." },
                { "자르기", "창에서 영역을 지정해 저장합니다. 이미지 한 장일 때 쓸 수 있습니다." },
                { "DPI", "값을 바꿔 저장한 뒤, 저장된 파일을 다시 읽어 확인합니다." },
                { "MD 변환", "Word, PowerPoint, Excel, 한글(.hwp, .hwpx 등), PDF, 텍스트(.txt, .csv, .html, .rtf 등)를 Markdown으로 저장합니다. 그림은 옆 폴더에 담습니다. 스캔한 PDF는 글자 인식으로 읽습니다." },
                { "문서 → PDF", "Word, PowerPoint, Excel, 한글, 텍스트(.txt, .html, .rtf) 파일을 PDF로 저장합니다. 해당 프로그램이 설치된 PC에서만 쓸 수 있으며, 한글 파일은 열린 한글 창을 닫은 뒤 변환합니다." },
                { "한글 ↔ 워드", "한글 문서는 DOCX로, Word 문서는 HWP로 저장합니다. 한컴오피스가 필요하며, 쪽 배치는 달라질 수 있습니다." },
                { "표 → Excel", ".csv와 .tsv를 XLSX로 저장합니다." }
            };
            for (int i = 0; i < rows.GetLength(0); i++)
            {
                features.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var name = new TextBlock { Text = rows[i, 0], FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 8, 7) };
                TextBlock text = Ui.Hint(rows[i, 1]);
                text.Margin = new Thickness(0, 0, 0, 7);
                Grid.SetRow(name, i);
                Grid.SetRow(text, i);
                Grid.SetColumn(text, 1);
                features.Children.Add(name);
                features.Children.Add(text);
            }
            var featureBox = new StackPanel();
            TextBlock heading = Ui.Heading("할 수 있는 일");
            heading.Margin = new Thickness(0, 0, 0, 10);
            featureBox.Children.Add(heading);
            featureBox.Children.Add(features);
            page.Children.Add(Ui.Card(featureBox));

            TextBlock note = Ui.Hint("결과는 원본과 같은 폴더에 저장됩니다. 원본은 바뀌지 않고, 같은 이름이 있으면 번호를 붙입니다.");
            note.Margin = new Thickness(2, 0, 2, 8);
            page.Children.Add(note);
            return page;
        }

        private static UIElement Step(string number, string text)
        {
            return Step(number, new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap });
        }

        private static UIElement Step(string number, UIElement content)
        {
            var row = new DockPanel { Margin = new Thickness(0, 10, 0, 0) };
            var badge = new TextBlock { Text = number, FontWeight = FontWeights.Bold, Foreground = (Brush)Application.Current.FindResource("Accent"), Width = 20 };
            DockPanel.SetDock(badge, Dock.Left);
            row.Children.Add(badge);
            row.Children.Add(content);
            return row;
        }

        // ---- About ----

        private UIElement BuildAbout(Action openLog, Action quit)
        {
            var page = new StackPanel { Margin = new Thickness(20, 18, 20, 8) };

            var about = new DockPanel();
            var fox = new Image { Width = 64, Height = 64, Margin = new Thickness(0, 0, 14, 0), VerticalAlignment = VerticalAlignment.Top };
            using (Stream icon = Theme.OpenIconStream())
            {
                BitmapDecoder decoder = BitmapDecoder.Create(icon, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
                fox.Source = decoder.Frames.OrderByDescending(f => f.PixelWidth).First();
            }
            DockPanel.SetDock(fox, Dock.Left);
            about.Children.Add(fox);
            var aboutText = new StackPanel();
            aboutText.Children.Add(Ui.Heading("Fanak 변환기 " + Version));
            TextBlock text = Ui.Hint("Fanak(فنك)은 사막여우를 뜻하는 아랍어입니다. 이 폴더만 복사하면 다른 PC에서도 설치 없이 실행됩니다. 파일은 PC 안에서만 처리하며 외부로 보내지 않습니다.");
            text.Margin = new Thickness(0, 6, 0, 0);
            aboutText.Children.Add(text);
            about.Children.Add(aboutText);
            page.Children.Add(Ui.Card(about));

            var actions = new StackPanel();
            actions.Children.Add(Ui.Heading("문제가 생겼을 때"));
            TextBlock hint = Ui.Hint("변환 기록에서 실패 원인을 볼 수 있습니다.");
            hint.Margin = new Thickness(0, 6, 0, 10);
            actions.Children.Add(hint);
            var row = new StackPanel { Orientation = Orientation.Horizontal };
            var log = new Button { Content = "변환 기록 열기" };
            log.Click += (s, e) => openLog();
            var exit = new Button { Content = "프로그램 종료", Margin = new Thickness(8, 0, 0, 0) };
            exit.Click += (s, e) => quit();
            row.Children.Add(log);
            row.Children.Add(exit);
            actions.Children.Add(row);
            page.Children.Add(Ui.Card(actions));
            return page;
        }
    }
}
