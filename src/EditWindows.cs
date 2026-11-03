using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using IOPath = System.IO.Path;

namespace Fanak
{
    internal static class Ui
    {
        private static readonly string[] DpiChoices = { "72", "96", "150", "300", "600", "1000" };
        private const string KeepDpiLabel = "원본 유지";
        private const string KeepFormatLabel = "원본 형식";

        public static readonly Brush ErrorBrush = Brushes.Firebrick;

        public static void Prepare(Window window)
        {
            window.FontFamily = new FontFamily("Segoe UI, Malgun Gothic");
            window.FontSize = 13;
            window.Background = Brushes.White;
            window.Icon = Theme.AppIcon;
            window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            window.UseLayoutRounding = true;
            TextOptions.SetTextFormattingMode(window, TextFormattingMode.Display);
            // The drop came from another app, so Windows may otherwise open this behind Explorer.
            window.Topmost = true;
            window.ContentRendered += (s, e) =>
            {
                window.Activate();
                window.Topmost = false;
                // Enter and Esc only reach the default and cancel buttons once a control has focus.
                window.MoveFocus(new TraversalRequest(FocusNavigationDirection.First));
            };
        }

        public static Style Find(string key)
        {
            return (Style)Application.Current.FindResource(key);
        }

        public static TextBlock Heading(string text)
        {
            return new TextBlock { Text = text, Style = Find("Heading") };
        }

        public static TextBlock Hint(string text)
        {
            return new TextBlock { Text = text, Style = Find("Hint") };
        }

        public static Border Card(UIElement content)
        {
            return new Border { Style = Find("Card"), Child = content };
        }

        public static ComboBox DpiBox(string initial, bool allowKeep)
        {
            var box = new ComboBox { IsEditable = true, Width = 104, VerticalContentAlignment = VerticalAlignment.Center, Padding = new Thickness(6, 5, 6, 5) };
            if (allowKeep) box.Items.Add(KeepDpiLabel);
            foreach (string choice in DpiChoices) box.Items.Add(choice);
            box.Text = initial;
            return box;
        }

        public static ComboBox FormatBox(bool allowKeep)
        {
            var box = new ComboBox { Width = 104, Padding = new Thickness(6, 5, 6, 5) };
            if (allowKeep) box.Items.Add(KeepFormatLabel);
            foreach (string format in new[] { "PNG", "JPG", "TIFF" }) box.Items.Add(format);
            box.SelectedIndex = 0;
            return box;
        }

        public static string SelectedFormat(ComboBox box)
        {
            string text = (string)box.SelectedItem;
            return text == KeepFormatLabel ? Converter.KeepFormat : text;
        }

        public static string KeepDpi
        {
            get { return KeepDpiLabel; }
        }

        public static bool TryReadDpi(ComboBox box, out int? dpi)
        {
            dpi = null;
            string text = box.Text.Trim();
            if (text == KeepDpiLabel) return true;
            int value;
            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value)) return false;
            if (value < Converter.MinDpi || value > Converter.MaxDpi) return false;
            dpi = value;
            return true;
        }

        public static string DpiError
        {
            get { return "DPI는 " + Converter.MinDpi + "~" + Converter.MaxDpi + " 사이의 숫자로 입력하세요."; }
        }

        public static UIElement Labeled(string label, UIElement control)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 0, 14, 0) };
            panel.Children.Add(new TextBlock { Text = label, Style = Find("Hint"), Margin = new Thickness(0, 0, 0, 4) });
            panel.Children.Add(control);
            return panel;
        }

        public static StackPanel SaveCancelButtons(Action save, Action close)
        {
            var save1 = new Button { Content = "저장", Style = Find("PrimaryButton"), IsDefault = true };
            var cancel = new Button { Content = "취소", IsCancel = true, Margin = new Thickness(8, 0, 0, 0) };
            save1.Click += (s, e) => save();
            cancel.Click += (s, e) => close();
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(save1);
            buttons.Children.Add(cancel);
            return buttons;
        }
    }

    internal sealed class DpiWindow : Window
    {
        private readonly string[] _files;
        private readonly Action<ConvertRequest> _onSave;
        private readonly ComboBox _dpi;
        private readonly ComboBox _format;
        private readonly CheckBox _resample;
        private readonly TextBlock _preview;
        private readonly TextBlock _modeHint;
        private readonly bool _hasInfo;
        private readonly int _pixelWidth;
        private readonly int _pixelHeight;
        private readonly double _sourceDpi;

        public DpiWindow(string[] files, int defaultDpi, Action<ConvertRequest> onSave)
        {
            _files = files;
            _onSave = onSave;
            bool anyPdf = files.Any(f => Converter.KindOf(f) == FileKind.Pdf);
            bool anyImage = files.Any(f => Converter.KindOf(f) == FileKind.Image);

            Title = "DPI 설정";
            SizeToContent = SizeToContent.Height;
            Width = 430;
            ResizeMode = ResizeMode.NoResize;
            Ui.Prepare(this);

            var panel = new StackPanel { Margin = new Thickness(20) };

            var info = new StackPanel();
            string summary = files.Length == 1 ? IOPath.GetFileName(files[0]) : files.Length + "개 파일";
            info.Children.Add(new TextBlock { Text = summary, Style = Ui.Find("Heading"), TextTrimming = TextTrimming.CharacterEllipsis });
            if (files.Length == 1 && anyImage)
                _hasInfo = Converter.TryReadImageInfo(files[0], out _pixelWidth, out _pixelHeight, out _sourceDpi);
            if (_hasInfo)
            {
                TextBlock current = Ui.Hint("현재  " + Describe(_sourceDpi, _pixelWidth, _pixelHeight));
                current.Margin = new Thickness(0, 4, 0, 0);
                info.Children.Add(current);
            }
            _preview = Ui.Hint("");
            _preview.Foreground = (Brush)Application.Current.FindResource("Ink");
            _preview.Margin = new Thickness(0, 2, 0, 0);
            info.Children.Add(_preview);
            panel.Children.Add(Ui.Card(info));

            var row = new StackPanel { Orientation = Orientation.Horizontal };
            _dpi = Ui.DpiBox(defaultDpi.ToString(CultureInfo.InvariantCulture), false);
            _format = Ui.FormatBox(!anyPdf);
            row.Children.Add(Ui.Labeled("DPI", _dpi));
            row.Children.Add(Ui.Labeled("저장 형식", _format));
            panel.Children.Add(row);

            _resample = new CheckBox { Content = "인쇄 크기 유지 (픽셀 수를 다시 계산)", Margin = new Thickness(0, 14, 0, 0) };
            if (anyImage) panel.Children.Add(_resample);
            _modeHint = Ui.Hint("");
            _modeHint.Margin = new Thickness(0, 6, 0, 0);
            panel.Children.Add(_modeHint);

            StackPanel buttons = Ui.SaveCancelButtons(Save, Close);
            buttons.Margin = new Thickness(0, 18, 0, 0);
            panel.Children.Add(buttons);
            Content = panel;

            _dpi.AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler((s, e) => UpdatePreview()));
            _dpi.SelectionChanged += (s, e) => Dispatcher.BeginInvoke(new Action(UpdatePreview));
            _resample.Checked += (s, e) => UpdatePreview();
            _resample.Unchecked += (s, e) => UpdatePreview();
            UpdatePreview();
        }

        private static string Describe(double dpi, int width, int height)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:0.#} DPI  ·  {1:N0} × {2:N0} 픽셀  ·  인쇄 {3:0.#} × {4:0.#} cm",
                dpi, width, height, width / dpi * 2.54, height / dpi * 2.54);
        }

        private void UpdatePreview()
        {
            bool resample = _resample.IsChecked == true;
            int? dpi;
            bool valid = Ui.TryReadDpi(_dpi, out dpi) && dpi.HasValue;

            if (!valid)
            {
                _preview.Text = Ui.DpiError;
                _preview.Foreground = Ui.ErrorBrush;
            }
            else if (_hasInfo)
            {
                double scale = resample ? dpi.Value / _sourceDpi : 1;
                _preview.Text = "변경 후  " + Describe(dpi.Value, (int)Math.Round(_pixelWidth * scale), (int)Math.Round(_pixelHeight * scale));
                _preview.Foreground = (Brush)Application.Current.FindResource("Ink");
            }
            else
            {
                _preview.Text = "변경 후  " + dpi.Value + " DPI";
                _preview.Foreground = (Brush)Application.Current.FindResource("Ink");
            }

            if (_files.All(f => Converter.KindOf(f) == FileKind.Pdf)) _modeHint.Text = "PDF 페이지를 이 해상도의 이미지로 저장합니다.";
            else if (resample) _modeHint.Text = "인쇄 크기가 같도록 픽셀 수를 바꿉니다. 픽셀을 늘려도 원본보다 선명해지지는 않습니다.";
            else _modeHint.Text = "픽셀은 그대로 두고 DPI 값만 바꿉니다. 인쇄 크기만 달라집니다.";
        }

        private void Save()
        {
            int? dpi;
            if (!Ui.TryReadDpi(_dpi, out dpi) || !dpi.HasValue)
            {
                _dpi.Focus();
                return;
            }
            _onSave(new ConvertRequest
            {
                Files = _files,
                Format = Ui.SelectedFormat(_format),
                Dpi = dpi,
                Resample = _resample.IsChecked == true
            });
            Close();
        }
    }

    internal sealed class CropWindow : Window
    {
        private const double GrabMargin = 8;
        private const double MinSelection = 4;
        public const double StagePadding = 16;

        private readonly string _file;
        private readonly BitmapSource _image;
        private readonly Action<ConvertRequest> _onSave;
        private readonly double _scale;
        private readonly double _surfaceWidth;
        private readonly double _surfaceHeight;
        private readonly Canvas _surface;
        private readonly Path _shade;
        private readonly Rectangle _frame;
        private readonly TextBlock _sizeLabel;
        private readonly TextBlock _message;
        private readonly ComboBox _ratio;
        private readonly ComboBox _dpi;
        private readonly ComboBox _format;

        private Rect _selection;
        private double _aspect;
        private bool _dragging;
        private bool _moving;
        private bool _movesLeft, _movesRight, _movesTop, _movesBottom;
        private Point _dragStart;
        private Rect _startSelection;

        public CropWindow(string file, BitmapSource image, Action<ConvertRequest> onSave)
        {
            _file = file;
            _image = image;
            _onSave = onSave;

            Title = "이미지 자르기 - " + IOPath.GetFileName(file);
            SizeToContent = SizeToContent.WidthAndHeight;
            ResizeMode = ResizeMode.NoResize;
            Ui.Prepare(this);

            // Large images shrink to fit; tiny ones grow so the selection handles stay easy to grab.
            double maxWidth = SystemParameters.WorkArea.Width * 0.85;
            double maxHeight = SystemParameters.WorkArea.Height * 0.68;
            double fit = Math.Min(maxWidth / image.PixelWidth, maxHeight / image.PixelHeight);
            double grow = Math.Max(1, Math.Min(4, 360.0 / Math.Max(image.PixelWidth, image.PixelHeight)));
            _scale = Math.Min(fit, grow);
            _surfaceWidth = Math.Max(1, image.PixelWidth * _scale);
            _surfaceHeight = Math.Max(1, image.PixelHeight * _scale);

            _surface = new Canvas { Width = _surfaceWidth, Height = _surfaceHeight, Background = Brushes.Black, ClipToBounds = true, Cursor = Cursors.Cross };
            var picture = new Image { Source = image, Width = _surfaceWidth, Height = _surfaceHeight, Stretch = Stretch.Fill, IsHitTestVisible = false };
            RenderOptions.SetBitmapScalingMode(picture, _scale > 1 ? BitmapScalingMode.NearestNeighbor : BitmapScalingMode.HighQuality);
            _shade = new Path { Fill = new SolidColorBrush(Color.FromArgb(0xA0, 0, 0, 0)), IsHitTestVisible = false };
            _frame = new Rectangle { Stroke = Brushes.White, StrokeThickness = 2, IsHitTestVisible = false };
            _surface.Children.Add(picture);
            _surface.Children.Add(_shade);
            _surface.Children.Add(_frame);
            _surface.MouseLeftButtonDown += OnSurfaceDown;
            _surface.MouseMove += OnSurfaceMove;
            _surface.MouseLeftButtonUp += OnSurfaceUp;

            var stage = new Grid { Background = new SolidColorBrush(Color.FromRgb(0x2B, 0x2B, 0x31)) };
            _surface.Margin = new Thickness(StagePadding);
            stage.Children.Add(_surface);

            _ratio = new ComboBox { Width = 84, Padding = new Thickness(6, 5, 6, 5) };
            foreach (string item in new[] { "자유", "1:1", "4:3", "3:2", "16:9" }) _ratio.Items.Add(item);
            _ratio.SelectedIndex = 0;
            _ratio.SelectionChanged += OnRatioChanged;
            _format = Ui.FormatBox(true);
            _dpi = Ui.DpiBox(Ui.KeepDpi, true);
            var whole = new Button { Content = "전체 선택", VerticalAlignment = VerticalAlignment.Bottom };
            whole.Click += (s, e) => SelectWhole();

            _sizeLabel = new TextBlock { Style = Ui.Find("Heading") };
            _message = Ui.Hint("이미지 위를 드래그해 영역을 지정하세요. 테두리를 끌면 크기가 바뀌고, 안쪽을 끌면 이동합니다.");

            var status = new DockPanel { Margin = new Thickness(0, 0, 0, 12) };
            _sizeLabel.Margin = new Thickness(0, 0, 16, 0);
            DockPanel.SetDock(_sizeLabel, Dock.Left);
            status.Children.Add(_sizeLabel);
            status.Children.Add(_message);

            var options = new StackPanel { Orientation = Orientation.Horizontal };
            options.Children.Add(Ui.Labeled("비율", _ratio));
            options.Children.Add(Ui.Labeled("저장 형식", _format));
            options.Children.Add(Ui.Labeled("DPI", _dpi));
            options.Children.Add(whole);

            var bar = new DockPanel { LastChildFill = false };
            DockPanel.SetDock(options, Dock.Left);
            bar.Children.Add(options);
            StackPanel buttons = Ui.SaveCancelButtons(Save, Close);
            buttons.VerticalAlignment = VerticalAlignment.Bottom;
            buttons.Margin = new Thickness(24, 0, 0, 0);
            DockPanel.SetDock(buttons, Dock.Right);
            bar.Children.Add(buttons);

            var footer = new StackPanel { Margin = new Thickness(18, 14, 18, 16), MinWidth = 640 };
            // Keeps a long hint from widening the window past the image.
            footer.Children.Add(status);
            footer.Children.Add(bar);

            var layout = new StackPanel();
            layout.Children.Add(stage);
            layout.Children.Add(footer);
            Content = layout;

            _selection = new Rect(_surfaceWidth * 0.1, _surfaceHeight * 0.1, _surfaceWidth * 0.8, _surfaceHeight * 0.8);
            Redraw();
        }

        private void SelectWhole()
        {
            _ratio.SelectedIndex = 0;
            _selection = new Rect(0, 0, _surfaceWidth, _surfaceHeight);
            Redraw();
        }

        private void OnSurfaceDown(object sender, MouseButtonEventArgs e)
        {
            Point p = Limit(e.GetPosition(_surface));
            _dragStart = p;
            _startSelection = _selection;
            ReadGrab(p, out _movesLeft, out _movesRight, out _movesTop, out _movesBottom, out _moving);

            if (!_moving && !_movesLeft && !_movesRight && !_movesTop && !_movesBottom)
            {
                _startSelection = new Rect(p, new Size(0, 0));
                _selection = _startSelection;
                _movesRight = true;
                _movesBottom = true;
            }
            _dragging = true;
            _surface.CaptureMouse();
            e.Handled = true;
        }

        private void OnSurfaceMove(object sender, MouseEventArgs e)
        {
            Point p = Limit(e.GetPosition(_surface));
            if (!_dragging)
            {
                bool left, right, top, bottom, inside;
                ReadGrab(p, out left, out right, out top, out bottom, out inside);
                if ((left && top) || (right && bottom)) _surface.Cursor = Cursors.SizeNWSE;
                else if ((right && top) || (left && bottom)) _surface.Cursor = Cursors.SizeNESW;
                else if (left || right) _surface.Cursor = Cursors.SizeWE;
                else if (top || bottom) _surface.Cursor = Cursors.SizeNS;
                else _surface.Cursor = inside ? Cursors.SizeAll : Cursors.Cross;
                return;
            }

            if (_moving)
            {
                double x = Clamp(_startSelection.X + p.X - _dragStart.X, 0, _surfaceWidth - _startSelection.Width);
                double y = Clamp(_startSelection.Y + p.Y - _dragStart.Y, 0, _surfaceHeight - _startSelection.Height);
                _selection = new Rect(x, y, _startSelection.Width, _startSelection.Height);
            }
            else
            {
                Resize(p);
            }
            Redraw();
        }

        private void OnSurfaceUp(object sender, MouseButtonEventArgs e)
        {
            _dragging = false;
            _surface.ReleaseMouseCapture();
            Redraw();
        }

        private void Resize(Point p)
        {
            Rect s = _startSelection;
            double x0 = _movesLeft ? p.X : s.Left;
            double x1 = _movesRight ? p.X : s.Right;
            double y0 = _movesTop ? p.Y : s.Top;
            double y1 = _movesBottom ? p.Y : s.Bottom;

            // The edge being dragged may cross the fixed edge, so work from whichever one stays put.
            double anchorX = _movesLeft ? x1 : x0;
            double anchorY = _movesTop ? y1 : y0;
            double movingX = _movesLeft ? x0 : x1;
            double movingY = _movesTop ? y0 : y1;
            bool towardLeft = movingX < anchorX;
            bool towardTop = movingY < anchorY;
            double width = Math.Abs(x1 - x0);
            double height = Math.Abs(y1 - y0);

            if (_aspect > 0)
            {
                bool heightDrives = !_movesLeft && !_movesRight;
                if (heightDrives) width = height * _aspect;
                else height = width / _aspect;
                double roomX = towardLeft ? anchorX : _surfaceWidth - anchorX;
                double roomY = towardTop ? anchorY : _surfaceHeight - anchorY;
                double fit = Math.Min(1, Math.Min(width > 0 ? roomX / width : 1, height > 0 ? roomY / height : 1));
                width *= fit;
                height *= fit;
            }

            _selection = new Rect(towardLeft ? anchorX - width : anchorX, towardTop ? anchorY - height : anchorY, width, height);
        }

        private void ReadGrab(Point p, out bool left, out bool right, out bool top, out bool bottom, out bool inside)
        {
            Rect s = _selection;
            bool withinX = p.X >= s.Left - GrabMargin && p.X <= s.Right + GrabMargin;
            bool withinY = p.Y >= s.Top - GrabMargin && p.Y <= s.Bottom + GrabMargin;
            left = withinY && Math.Abs(p.X - s.Left) <= GrabMargin;
            right = withinY && !left && Math.Abs(p.X - s.Right) <= GrabMargin;
            top = withinX && Math.Abs(p.Y - s.Top) <= GrabMargin;
            bottom = withinX && !top && Math.Abs(p.Y - s.Bottom) <= GrabMargin;
            inside = !left && !right && !top && !bottom && s.Contains(p);
        }

        private void OnRatioChanged(object sender, SelectionChangedEventArgs e)
        {
            _aspect = 0;
            string[] parts = ((string)_ratio.SelectedItem).Split(':');
            if (parts.Length == 2) _aspect = double.Parse(parts[0], CultureInfo.InvariantCulture) / double.Parse(parts[1], CultureInfo.InvariantCulture);
            if (_aspect <= 0) return;

            double width = _selection.Width;
            double height = width / _aspect;
            if (height > _selection.Height)
            {
                height = _selection.Height;
                width = height * _aspect;
            }
            _selection = new Rect(_selection.X + (_selection.Width - width) / 2, _selection.Y + (_selection.Height - height) / 2, width, height);
            Redraw();
        }

        private Int32Rect SelectionPixels()
        {
            int x = (int)Math.Round(_selection.X / _scale);
            int y = (int)Math.Round(_selection.Y / _scale);
            int width = (int)Math.Round(_selection.Width / _scale);
            int height = (int)Math.Round(_selection.Height / _scale);
            x = Math.Max(0, Math.Min(x, _image.PixelWidth - 1));
            y = Math.Max(0, Math.Min(y, _image.PixelHeight - 1));
            width = Math.Min(width, _image.PixelWidth - x);
            height = Math.Min(height, _image.PixelHeight - y);
            return new Int32Rect(x, y, width, height);
        }

        private void Redraw()
        {
            var outer = new RectangleGeometry(new Rect(0, 0, _surfaceWidth, _surfaceHeight));
            _shade.Data = new CombinedGeometry(GeometryCombineMode.Exclude, outer, new RectangleGeometry(_selection));
            Canvas.SetLeft(_frame, _selection.X);
            Canvas.SetTop(_frame, _selection.Y);
            _frame.Width = _selection.Width;
            _frame.Height = _selection.Height;
            Int32Rect pixels = SelectionPixels();
            _sizeLabel.Text = string.Format(CultureInfo.InvariantCulture, "{0:N0} × {1:N0} 픽셀", pixels.Width, pixels.Height);
        }

        private void Save()
        {
            int? dpi;
            if (!Ui.TryReadDpi(_dpi, out dpi))
            {
                ShowError(Ui.DpiError);
                return;
            }
            if (_selection.Width < MinSelection || _selection.Height < MinSelection)
            {
                ShowError("자를 영역을 먼저 드래그해 주세요.");
                return;
            }
            _onSave(new ConvertRequest
            {
                Files = new[] { _file },
                Format = Ui.SelectedFormat(_format),
                Dpi = dpi,
                Crop = SelectionPixels()
            });
            Close();
        }

        private void ShowError(string message)
        {
            _message.Text = message;
            _message.Foreground = Ui.ErrorBrush;
        }

        private Point Limit(Point p)
        {
            return new Point(Clamp(p.X, 0, _surfaceWidth), Clamp(p.Y, 0, _surfaceHeight));
        }

        private static double Clamp(double value, double low, double high)
        {
            return Math.Max(low, Math.Min(value, high));
        }
    }
}
