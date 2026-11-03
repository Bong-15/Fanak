using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using File = System.IO.File;

namespace Fanak
{
    // A click-free drop target. It stays almost fully transparent until a file drag actually enters it,
    // so a hotkey drag that is not a file drag (such as a selection box) shows nothing.
    internal sealed class RadialWindow : Window
    {
        private const double MenuSize = 420;
        private const double InnerRadius = 72;
        private const double OuterRadius = 190;
        private const double GapDegrees = 1.5;
        private const double GroupGapDegrees = 4.5;
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        private static readonly Brush ImageBrush = Frozen(Color.FromArgb(0xF2, 0x2B, 0x2B, 0x2D));
        private static readonly Brush DocumentBrush = Frozen(Color.FromArgb(0xF2, 0x38, 0x38, 0x3A));
        private static readonly Brush ToolBrush = Frozen(Color.FromArgb(0xF2, 0x47, 0x47, 0x4A));
        private static readonly Brush HoverBrush = Frozen(Color.FromRgb(0xF2, 0x8C, 0x28));
        private static readonly Brush DisabledBrush = Frozen(Color.FromArgb(0xB0, 0x26, 0x26, 0x2B));
        private static readonly Brush CenterBrush = Frozen(Color.FromArgb(0xF5, 0x18, 0x18, 0x1C));
        private static readonly Brush TextBrush = Frozen(Colors.White);
        private static readonly FontFamily MenuFont = new FontFamily("Segoe UI, Malgun Gothic");

        private readonly Canvas _menu;
        private readonly List<Path> _sectors = new List<Path>();
        private readonly TextBlock _centerText;
        private string[] _files;
        private List<string> _formats = new List<string>();
        private string _message;
        private int _selected = -1;

        public event Action<string[], string> FormatChosen;

        public RadialWindow()
        {
            Title = "Fanak 변환기";
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            ShowInTaskbar = false;
            ShowActivated = false;
            Topmost = true;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Width = MenuSize;
            Height = MenuSize;
            AllowDrop = true;

            _menu = new Canvas { Width = MenuSize, Height = MenuSize, Opacity = 0, IsHitTestVisible = false };
            _centerText = new TextBlock
            {
                Width = InnerRadius * 2 - 20,
                Foreground = TextBrush,
                FontFamily = MenuFont,
                FontSize = 13,
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap
            };

            // Alpha 1 keeps the layered window hit-testable, which OLE needs to find a drop target.
            var root = new Grid { Background = Frozen(Color.FromArgb(1, 0, 0, 0)) };
            root.Children.Add(_menu);
            Content = root;

            DragEnter += OnDragEnter;
            DragOver += OnDragOver;
            Drop += OnDrop;
            SourceInitialized += (s, e) => ApplyWindowStyle();
        }

        private static Brush Frozen(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        private void ApplyWindowStyle()
        {
            IntPtr hwnd = new WindowInteropHelper(this).Handle;
            int style = Native.GetWindowLong(hwnd, GWL_EXSTYLE);
            Native.SetWindowLong(hwnd, GWL_EXSTYLE, style | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW);
        }

        public void ShowAtCursor()
        {
            IntPtr hwnd = new WindowInteropHelper(this).EnsureHandle();
            Native.POINT cursor;
            if (!Native.GetCursorPos(out cursor)) return;

            Point center = HwndSource.FromHwnd(hwnd).CompositionTarget.TransformFromDevice.Transform(new Point(cursor.X, cursor.Y));
            double maxLeft = SystemParameters.VirtualScreenLeft + SystemParameters.VirtualScreenWidth - MenuSize;
            double maxTop = SystemParameters.VirtualScreenTop + SystemParameters.VirtualScreenHeight - MenuSize;
            Left = Math.Max(SystemParameters.VirtualScreenLeft, Math.Min(center.X - MenuSize / 2, maxLeft));
            Top = Math.Max(SystemParameters.VirtualScreenTop, Math.Min(center.Y - MenuSize / 2, maxTop));
            Show();
        }

        public void Dismiss()
        {
            Hide();
            _menu.BeginAnimation(OpacityProperty, null);
            _menu.Opacity = 0;
            _files = null;
            _formats = new List<string>();
            _message = null;
            _selected = -1;
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            e.Handled = true;
            string[] files = ReadFiles(e.Data);
            if (files == null)
            {
                e.Effects = DragDropEffects.None;
                return;
            }

            if (_files == null)
            {
                _files = files;
                MenuPlan plan = Converter.Plan(files);
                _formats = plan.Formats;
                _message = plan.Message;
                BuildMenu();
                _menu.BeginAnimation(OpacityProperty, new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(90)));
            }
            UpdateSelection(e);
        }

        private void OnDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
            if (_files == null)
            {
                e.Effects = DragDropEffects.None;
                return;
            }
            UpdateSelection(e);
        }

        private void OnDrop(object sender, DragEventArgs e)
        {
            e.Handled = true;
            if (_files == null)
            {
                e.Effects = DragDropEffects.None;
                return;
            }
            UpdateSelection(e);

            string[] files = _files;
            string format = _selected >= 0 ? _formats[_selected] : null;
            Dismiss();

            var handler = FormatChosen;
            if (format != null && handler != null) handler(files, format);
        }

        private static string[] ReadFiles(IDataObject data)
        {
            if (!data.GetDataPresent(DataFormats.FileDrop)) return null;
            var paths = data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null) return null;
            string[] files = paths.Where(File.Exists).ToArray();
            return files.Length > 0 ? files : null;
        }

        // Reports a copy, never a move: Explorer would delete the original after a move, and it must stay put.
        private void UpdateSelection(DragEventArgs e)
        {
            int selected = SectorAt(e.GetPosition(this));
            if (selected != _selected)
            {
                _selected = selected;
                for (int i = 0; i < _sectors.Count; i++)
                    _sectors[i].Fill = i == _selected ? HoverBrush : RestBrush(_formats[i]);
                UpdateCenterText();
            }

            if (_selected < 0) e.Effects = DragDropEffects.None;
            else if ((e.AllowedEffects & DragDropEffects.Copy) != 0) e.Effects = DragDropEffects.Copy;
            else if ((e.AllowedEffects & DragDropEffects.Link) != 0) e.Effects = DragDropEffects.Link;
            else e.Effects = DragDropEffects.None;
        }

        private int SectorAt(Point position)
        {
            if (_formats.Count == 0) return -1;
            double dx = position.X - MenuSize / 2;
            double dy = position.Y - MenuSize / 2;
            if (Math.Sqrt(dx * dx + dy * dy) < InnerRadius) return -1;

            double step = 360.0 / _formats.Count;
            double angle = (Math.Atan2(dy, dx) * 180 / Math.PI + 90 + 360) % 360;
            return (int)Math.Round(angle / step) % _formats.Count;
        }

        private enum MenuGroup { Image, Document, Tool }

        // Crop and DPI open a window instead of converting right away, so they form a group of their own.
        private static MenuGroup GroupOf(string format)
        {
            if (format == Converter.CropAction || format == Converter.DpiAction) return MenuGroup.Tool;
            return format == "PDF" || format == "MD" || format == "DOCX" || format == "HWP" || format == "XLSX" ? MenuGroup.Document : MenuGroup.Image;
        }

        private static Brush RestBrush(string format)
        {
            switch (GroupOf(format))
            {
                case MenuGroup.Tool: return ToolBrush;
                case MenuGroup.Document: return DocumentBrush;
                default: return ImageBrush;
            }
        }

        // The wider gap only falls where one group meets the next.
        private double GapBetween(int first, int second)
        {
            return GroupOf(_formats[first]) == GroupOf(_formats[second]) ? GapDegrees : GroupGapDegrees;
        }

        private void BuildMenu()
        {
            _menu.Children.Clear();
            _sectors.Clear();
            var center = new Point(MenuSize / 2, MenuSize / 2);

            if (_formats.Count == 0)
            {
                var ring = new Path
                {
                    Data = new CombinedGeometry(GeometryCombineMode.Exclude,
                        new EllipseGeometry(center, OuterRadius, OuterRadius),
                        new EllipseGeometry(center, InnerRadius, InnerRadius)),
                    Fill = DisabledBrush
                };
                _menu.Children.Add(ring);
            }

            double step = 360.0 / Math.Max(1, _formats.Count);
            for (int i = 0; i < _formats.Count; i++)
            {
                double middle = -90 + i * step;
                int previous = (i + _formats.Count - 1) % _formats.Count;
                int next = (i + 1) % _formats.Count;
                var sector = new Path
                {
                    Data = Sector(center, InnerRadius, OuterRadius, middle - step / 2 + GapBetween(previous, i), middle + step / 2 - GapBetween(i, next)),
                    Fill = RestBrush(_formats[i])
                };
                _sectors.Add(sector);
                _menu.Children.Add(sector);
            }

            for (int i = 0; i < _formats.Count; i++)
            {
                var label = new TextBlock
                {
                    Text = Converter.Label(_formats[i]),
                    Foreground = TextBrush,
                    FontFamily = MenuFont,
                    FontSize = 17,
                    FontWeight = FontWeights.SemiBold
                };
                label.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                Point spot = Polar(center, (InnerRadius + OuterRadius) / 2, -90 + i * step);
                Canvas.SetLeft(label, spot.X - label.DesiredSize.Width / 2);
                Canvas.SetTop(label, spot.Y - label.DesiredSize.Height / 2);
                _menu.Children.Add(label);
            }

            _menu.Children.Add(new Ellipse
            {
                Width = (InnerRadius - 4) * 2,
                Height = (InnerRadius - 4) * 2,
                Fill = CenterBrush,
                Margin = new Thickness(center.X - InnerRadius + 4, center.Y - InnerRadius + 4, 0, 0)
            });
            _menu.Children.Add(_centerText);
            UpdateCenterText();
        }

        private void UpdateCenterText()
        {
            if (_message != null) _centerText.Text = _message;
            else if (_selected >= 0 && _formats[_selected] == Converter.CropAction) _centerText.Text = "이미지 자르기";
            else if (_selected >= 0 && _formats[_selected] == Converter.DpiAction) _centerText.Text = "DPI 설정";
            else if (_selected >= 0) _centerText.Text = _formats[_selected] + "로 변환";
            else if (_files != null && _files.Length > 1) _centerText.Text = _files.Length + "개 파일\n놓으면 취소";
            else if (_files != null) _centerText.Text = Shorten(System.IO.Path.GetFileName(_files[0])) + "\n놓으면 취소";

            _centerText.Measure(new Size(_centerText.Width, double.PositiveInfinity));
            Canvas.SetLeft(_centerText, (MenuSize - _centerText.Width) / 2);
            Canvas.SetTop(_centerText, (MenuSize - _centerText.DesiredSize.Height) / 2);
        }

        private static string Shorten(string name)
        {
            return name.Length <= 16 ? name : name.Substring(0, 15) + "…";
        }

        private static Point Polar(Point center, double radius, double degrees)
        {
            double radians = degrees * Math.PI / 180;
            return new Point(center.X + radius * Math.Cos(radians), center.Y + radius * Math.Sin(radians));
        }

        private static Geometry Sector(Point center, double inner, double outer, double from, double to)
        {
            bool large = to - from > 180;
            var figure = new PathFigure { StartPoint = Polar(center, inner, from), IsClosed = true };
            figure.Segments.Add(new LineSegment(Polar(center, outer, from), true));
            figure.Segments.Add(new ArcSegment(Polar(center, outer, to), new Size(outer, outer), 0, large, SweepDirection.Clockwise, true));
            figure.Segments.Add(new LineSegment(Polar(center, inner, to), true));
            figure.Segments.Add(new ArcSegment(Polar(center, inner, from), new Size(inner, inner), 0, large, SweepDirection.Counterclockwise, true));
            var geometry = new PathGeometry();
            geometry.Figures.Add(figure);
            return geometry;
        }
    }
}
