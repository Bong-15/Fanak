using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;

namespace Fanak
{
    // A PDF only says where each letter sits on the page. This rebuilds lines, columns, tables,
    // paragraphs and headings from those positions, so the result is a good guess rather than an exact one.
    internal sealed class PdfToMd
    {
        private sealed class Frag
        {
            public double L, T, R, B, Size, Base;
            public List<PdfTextItem> Items;
            public List<List<string>> Table;
            public PdfImageMark Image;
            public bool IsText { get { return Items != null; } }
        }

        private sealed class TextRun
        {
            public string Text;
            public bool Bold, Italic;
        }

        private enum ParaKind { Text, Table, Image, Raw }

        private sealed class Para
        {
            public ParaKind Kind;
            public List<TextRun> Runs = new List<TextRun>();
            public List<List<string>> Table;
            public string Markdown;
            public double Size, Left, Right;
            public int Lines, ListLevel, Number, Page;
            public bool IsList, Ordered, Ocr, FilledLast;
        }

        private const double RenderDpi = 150;
        private const double OcrDpi = 300;
        private const double UnreadableShare = 0.2;
        private const double SmallestImage = 30;

        private static readonly Regex BulletStart = new Regex(@"^\s*[•·▪▫◦○●■□◆◇▶▷►‣∙⁃]\s*|^\s*[-–—*]\s+");
        private static readonly Regex NumberStart = new Regex(@"^\s*(\d{1,3})\.\s+");
        // Not "가." style markers: Korean sentences end in "다.", and a line often starts with the end of one.
        private static readonly Regex OtherListStart = new Regex(@"^\s*(\(?\d{1,3}\)|\(?[가-힣]\)\s|[①-⑳]|※)");
        private static readonly Regex PageNumberOnly = new Regex(@"^[-–—\s]*(page\s*)?#([\s/]*(of\s*)?#)?[-–—\s]*$", RegexOptions.IgnoreCase);
        private static readonly Regex SentenceEnd = new Regex(@"[.!?:;。)\]""'”’다요죠음함임됨]$");

        private readonly string _source;
        private readonly MdDoc _doc;
        private readonly MdImages _images;
        private List<PdfPageContent> _pages;
        private PdfDocument _windowsDoc;
        private OcrEngine _ocr;
        private bool _ocrChecked;

        public PdfToMd(string source, MdDoc doc, MdImages images)
        {
            _source = source;
            _doc = doc;
            _images = images;
        }

        public void Run()
        {
            OpenWindowsDocument();

            List<PdfPageContent> pages = ReadText();
            uint expected = _windowsDoc != null ? _windowsDoc.PageCount : (uint)(pages == null ? 0 : pages.Count);
            if (pages == null || pages.Count != expected)
            {
                // The text could not be read at all (encrypted, for example), so every page is scanned instead.
                if (_windowsDoc == null) throw new InvalidOperationException("PDF를 열 수 없습니다. 파일이 손상됐거나 암호가 걸려 있을 수 있습니다.");
                pages = new List<PdfPageContent>();
                for (uint i = 0; i < _windowsDoc.PageCount; i++)
                {
                    using (PdfPage page = _windowsDoc.GetPage(i))
                        pages.Add(new PdfPageContent { Width = page.Size.Width * 72 / 96, Height = page.Size.Height * 72 / 96, Ocr = true });
                }
            }

            _pages = pages;
            var scanned = new List<int>();
            var unreadable = new List<int>();
            var paras = new List<Para>();
            for (int i = 0; i < pages.Count; i++)
            {
                if (NeedsOcr(pages[i]) && !ApplyOcr(pages[i], (uint)i))
                {
                    AddWholePage(pages[i], (uint)i, paras);
                    unreadable.Add(i + 1);
                    continue;
                }
                if (pages[i].Ocr) scanned.Add(i + 1);
            }

            RemoveHeadersAndFooters(pages);
            RemoveRepeatedImages(pages);
            for (int i = 0; i < pages.Count; i++)
            {
                if (unreadable.Contains(i + 1)) continue;
                paras.AddRange(LayoutPage(pages[i], i));
            }
            paras = paras.OrderBy(p => p.Page).ToList();

            if (scanned.Count > 0) _doc.AddRaw("<!-- 글자 인식(OCR)으로 읽은 쪽: " + string.Join(", ", scanned) + " -->");
            if (unreadable.Count > 0) _doc.AddRaw("<!-- 글자를 읽지 못해 그림으로 넣은 쪽: " + string.Join(", ", unreadable) + " -->");
            Emit(paras);
        }

        private void OpenWindowsDocument()
        {
            try
            {
                StorageFile file = Converter.Wait(StorageFile.GetFileFromPathAsync(_source));
                _windowsDoc = Converter.Wait(PdfDocument.LoadFromFileAsync(file));
            }
            catch (Exception ex)
            {
                if (ex is OutOfMemoryException) throw;
                _windowsDoc = null;
                Log.Write("Windows PDF 엔진으로 열지 못함: " + _source + " - " + ex.Message);
            }
        }

        private List<PdfPageContent> ReadText()
        {
            try
            {
                var file = new PdfFile(File.ReadAllBytes(_source));
                if (file.Encrypted) return null;
                var reader = new PdfPageReader(file);
                var pages = new List<PdfPageContent>();
                foreach (PdfPageRef page in file.Pages()) pages.Add(reader.Read(page));
                return pages;
            }
            catch (Exception ex)
            {
                if (ex is OutOfMemoryException || ex is IOException || ex is UnauthorizedAccessException) throw;
                Log.Write("PDF 글자를 직접 읽지 못함: " + _source + " - " + ex);
                return null;
            }
        }

        // ---- Scanned pages ----

        private static bool NeedsOcr(PdfPageContent page)
        {
            if (page.Ocr) return true;
            int chars = page.Text.Sum(t => t.Text.Length);
            int bad = page.Text.Sum(t => t.Text.Count(c => c == '�'));
            if (chars > 0 && bad > chars * UnreadableShare) return true;
            double area = page.Width * page.Height;
            double covered = page.Images.Sum(m => (m.X1 - m.X0) * (m.Y1 - m.Y0));
            return chars < 10 && covered > area * 0.5;
        }

        private bool ApplyOcr(PdfPageContent page, uint index)
        {
            OcrEngine engine = Engine();
            if (engine == null || _windowsDoc == null) return false;

            double dpi = Math.Min(OcrDpi, OcrEngine.MaxImageDimension * 72 / Math.Max(1, Math.Max(page.Width, page.Height)));
            var items = new List<PdfTextItem>();
            using (PdfPage pdfPage = _windowsDoc.GetPage(index))
            using (var memory = new InMemoryRandomAccessStream())
            {
                var options = new PdfPageRenderOptions
                {
                    DestinationWidth = (uint)Math.Max(1, Math.Round(pdfPage.Size.Width * dpi / 96)),
                    DestinationHeight = (uint)Math.Max(1, Math.Round(pdfPage.Size.Height * dpi / 96))
                };
                Converter.Wait(pdfPage.RenderToStreamAsync(memory, options));
                BitmapDecoder decoder = Converter.Wait(BitmapDecoder.CreateAsync(memory));
                using (SoftwareBitmap bitmap = Converter.Wait(decoder.GetSoftwareBitmapAsync()))
                {
                    OcrResult result = Converter.Wait(engine.RecognizeAsync(bitmap));
                    double scale = 72 / dpi;
                    foreach (OcrLine line in result.Lines)
                    {
                        var words = line.Words.ToList();
                        if (words.Count == 0) continue;
                        double height = words.Max(w => (double)w.BoundingRect.Height) * scale;
                        double bottom = words.Max(w => (double)w.BoundingRect.Y + w.BoundingRect.Height) * scale;
                        foreach (OcrWord word in words)
                        {
                            items.Add(new PdfTextItem
                            {
                                X = word.BoundingRect.X * scale,
                                EndX = (word.BoundingRect.X + word.BoundingRect.Width) * scale,
                                Y = bottom,
                                Size = Math.Max(1, Math.Round(height)),
                                Text = word.Text
                            });
                        }
                    }
                }
            }

            if (items.Count == 0) return false;
            page.Text = items;
            page.Sideways.Clear();
            page.Lines.Clear();
            page.Images.Clear();
            page.Ocr = true;
            return true;
        }

        private OcrEngine Engine()
        {
            if (_ocrChecked) return _ocr;
            _ocrChecked = true;
            _ocr = OcrEngine.TryCreateFromUserProfileLanguages();
            if (_ocr == null)
            {
                var language = OcrEngine.AvailableRecognizerLanguages.FirstOrDefault();
                if (language != null) _ocr = OcrEngine.TryCreateFromLanguage(language);
            }
            if (_ocr == null) Log.Write("글자 인식(OCR) 언어가 이 PC에 설치되어 있지 않습니다.");
            return _ocr;
        }

        private void AddWholePage(PdfPageContent page, uint index, List<Para> paras)
        {
            if (_windowsDoc == null) return;
            byte[] png = RenderRegion(index, page, 0, 0, page.Width, page.Height);
            if (png == null) return;
            paras.Add(new Para { Kind = ParaKind.Image, Markdown = MdImages.Markdown("쪽 " + (index + 1), _images.Save(png, "png")), Page = (int)index });
        }

        // ---- Things that repeat on every page ----

        private static string FragKey(string text)
        {
            return Regex.Replace(Regex.Replace(text, @"\d+", "#"), @"\s+", " ").Trim();
        }

        // Losing real text is much worse than keeping a header, so a line has to repeat at the same
        // height on most pages. Only short lines may differ by a number, as "Page 3" does.
        private static void RemoveHeadersAndFooters(List<PdfPageContent> pages)
        {
            if (pages.Count < 2) return;
            const double Margin = 0.1;
            const int LongestNumberedLine = 40;
            var counts = new Dictionary<string, int>();
            var lines = new List<List<KeyValuePair<string, List<PdfTextItem>>>>();
            foreach (PdfPageContent page in pages)
            {
                var keys = new List<KeyValuePair<string, List<PdfTextItem>>>();
                foreach (List<PdfTextItem> line in Lines(page.Text))
                {
                    double y = line[0].Y;
                    if (y > page.Height * Margin && y < page.Height * (1 - Margin)) continue;
                    string text = string.Concat(line.Select(t => t.Text));
                    string key = FragKey(text);
                    if (key.Length == 0) continue;
                    if (key.Length > LongestNumberedLine) key = Regex.Replace(text, @"\s+", " ").Trim();
                    keys.Add(new KeyValuePair<string, List<PdfTextItem>>(Math.Round(y / 2) + "|" + key, line));
                }
                foreach (string key in keys.Select(k => k.Key).Distinct())
                {
                    int count;
                    counts.TryGetValue(key, out count);
                    counts[key] = count + 1;
                }
                lines.Add(keys);
            }

            for (int i = 0; i < pages.Count; i++)
            {
                double top = pages[i].Text.Count == 0 ? 0 : pages[i].Text.Min(t => t.Y);
                double bottom = pages[i].Text.Count == 0 ? 0 : pages[i].Text.Max(t => t.Y);
                foreach (KeyValuePair<string, List<PdfTextItem>> line in lines[i])
                {
                    string text = line.Key.Substring(line.Key.IndexOf('|') + 1);
                    bool repeated = pages.Count >= 3 && counts[line.Key] >= Math.Max(3, pages.Count * 0.5);
                    bool outermost = Math.Abs(line.Value[0].Y - top) < 1 || Math.Abs(line.Value[0].Y - bottom) < 1;
                    if (!repeated && !(outermost && PageNumberOnly.IsMatch(text))) continue;
                    foreach (PdfTextItem item in line.Value) pages[i].Text.Remove(item);
                }
            }
        }

        private static void RemoveRepeatedImages(List<PdfPageContent> pages)
        {
            var usage = new Dictionary<int, int>();
            foreach (PdfPageContent page in pages)
            {
                foreach (int id in page.Images.Where(m => m.Id > 0).Select(m => m.Id).Distinct())
                {
                    int count;
                    usage.TryGetValue(id, out count);
                    usage[id] = count + 1;
                }
            }
            foreach (PdfPageContent page in pages)
            {
                double area = page.Width * page.Height;
                bool hasText = page.Text.Count > 0;
                page.Images.RemoveAll(m =>
                {
                    int count;
                    bool repeated = m.Id > 0 && usage.TryGetValue(m.Id, out count) && count >= 3 && count >= pages.Count * 0.5;
                    bool background = hasText && (m.X1 - m.X0) * (m.Y1 - m.Y0) > area * 0.8;
                    bool tiny = m.X1 - m.X0 < SmallestImage || m.Y1 - m.Y0 < SmallestImage;
                    return repeated || background || tiny;
                });
            }
        }

        // ---- Page layout ----

        private List<Para> LayoutPage(PdfPageContent page, int index)
        {
            AddDrawnBullets(page);
            var frags = new List<Frag>();
            frags.AddRange(RuledTables(page));
            foreach (List<PdfTextItem> line in Lines(page.Text)) frags.AddRange(SplitLine(line));
            frags.AddRange(ImageFrags(page, index));

            var paras = new List<Para>();
            foreach (List<Frag> leaf in ReadingOrder(frags)) LayoutLeaf(leaf, index, page.Ocr, paras);

            if (page.Sideways.Count > 0)
            {
                string text = string.Join(" ", page.Sideways.Select(t => t.Text.Trim()));
                var para = new Para { Page = index, Size = 0, Lines = 1 };
                para.Runs.Add(new TextRun { Text = text });
                paras.Add(para);
            }
            foreach (Para para in paras) para.Page = index;
            return paras;
        }

        // Web browsers draw list bullets as small shapes rather than as a character.
        private static void AddDrawnBullets(PdfPageContent page)
        {
            foreach (double[] dot in page.Dots)
            {
                double middle = (dot[1] + dot[3]) / 2;
                List<PdfTextItem> sameLine = page.Text.Where(t => middle > t.Y - t.Size && middle < t.Y + t.Size * 0.1).ToList();
                PdfTextItem next = sameLine.Where(t => t.X >= dot[2]).OrderBy(t => t.X).FirstOrDefault();
                if (next == null || next.X - dot[2] > next.Size * 2.5) continue;
                if (sameLine.Any(t => t.EndX <= dot[0] + 1 && dot[0] - t.EndX < next.Size * 2)) continue;
                page.Text.Add(new PdfTextItem { X = dot[0], EndX = dot[2], Y = next.Y, Size = next.Size, Text = "• " });
            }
        }

        private static List<List<PdfTextItem>> Lines(List<PdfTextItem> items)
        {
            var lines = new List<List<PdfTextItem>>();
            List<PdfTextItem> current = null;
            double y = 0, size = 0;
            foreach (PdfTextItem item in items.OrderBy(t => t.Y))
            {
                if (current != null && Math.Abs(item.Y - y) <= 0.45 * Math.Max(item.Size, size))
                {
                    current.Add(item);
                    size = Math.Max(size, item.Size);
                    continue;
                }
                current = new List<PdfTextItem> { item };
                lines.Add(current);
                y = item.Y;
                size = item.Size;
            }
            foreach (List<PdfTextItem> line in lines) line.Sort((a, b) => a.X.CompareTo(b.X));
            return lines;
        }

        // A wide gap on one line means two separate things: columns, or table cells.
        private static IEnumerable<Frag> SplitLine(List<PdfTextItem> line)
        {
            var current = new List<PdfTextItem>();
            double end = 0;
            foreach (PdfTextItem item in line)
            {
                if (current.Count > 0 && item.X - end > 1.5 * Math.Max(item.Size, current[0].Size))
                {
                    yield return MakeFrag(current);
                    current = new List<PdfTextItem>();
                }
                current.Add(item);
                end = current.Count == 1 ? item.EndX : Math.Max(end, item.EndX);
            }
            if (current.Count > 0) yield return MakeFrag(current);
        }

        private static Frag MakeFrag(List<PdfTextItem> items)
        {
            double size = items.GroupBy(t => Math.Round(t.Size * 2) / 2).OrderByDescending(g => g.Sum(t => t.Text.Length)).First().Key;
            double y = items.OrderByDescending(t => t.Text.Length).First().Y;
            return new Frag
            {
                Items = items,
                L = items.Min(t => t.X),
                R = items.Max(t => t.EndX),
                T = y - size * 0.8,
                B = y + size * 0.25,
                Size = size,
                Base = y
            };
        }

        // Figures are copied from the rendered page, which handles every image format inside PDFs.
        private IEnumerable<Frag> ImageFrags(PdfPageContent page, int index)
        {
            if (_windowsDoc == null) yield break;
            var marks = new List<PdfImageMark>(page.Images);
            bool merged = true;
            while (merged)
            {
                merged = false;
                for (int i = 0; i < marks.Count && !merged; i++)
                {
                    for (int k = i + 1; k < marks.Count && !merged; k++)
                    {
                        PdfImageMark a = marks[i], b = marks[k];
                        if (a.X0 - 2 > b.X1 || b.X0 - 2 > a.X1 || a.Y0 - 2 > b.Y1 || b.Y0 - 2 > a.Y1) continue;
                        marks[i] = new PdfImageMark { Id = -1, X0 = Math.Min(a.X0, b.X0), Y0 = Math.Min(a.Y0, b.Y0), X1 = Math.Max(a.X1, b.X1), Y1 = Math.Max(a.Y1, b.Y1) };
                        marks.RemoveAt(k);
                        merged = true;
                    }
                }
            }
            foreach (PdfImageMark mark in marks)
            {
                yield return new Frag { Image = mark, L = mark.X0, T = mark.Y0, R = mark.X1, B = mark.Y1, Size = 0 };
            }
        }

        private byte[] RenderRegion(uint index, PdfPageContent content, double x0, double y0, double x1, double y1)
        {
            x0 = Math.Max(0, x0);
            y0 = Math.Max(0, y0);
            x1 = Math.Min(content.Width, x1);
            y1 = Math.Min(content.Height, y1);
            if (x1 - x0 < 1 || y1 - y0 < 1) return null;

            using (PdfPage page = _windowsDoc.GetPage(index))
            using (var memory = new InMemoryRandomAccessStream())
            {
                double factor = page.Size.Width / Math.Max(1, content.Width);
                var options = new PdfPageRenderOptions
                {
                    SourceRect = new Windows.Foundation.Rect(x0 * factor, y0 * factor, (x1 - x0) * factor, (y1 - y0) * factor),
                    DestinationWidth = (uint)Math.Max(1, Math.Round((x1 - x0) * RenderDpi / 72)),
                    DestinationHeight = (uint)Math.Max(1, Math.Round((y1 - y0) * RenderDpi / 72))
                };
                Converter.Wait(page.RenderToStreamAsync(memory, options));
                var png = new byte[(int)memory.Size];
                using (var reader = new DataReader(memory.GetInputStreamAt(0)))
                {
                    Converter.Wait(reader.LoadAsync((uint)png.Length));
                    reader.ReadBytes(png);
                }
                return png;
            }
        }

        // ---- Tables drawn with lines ----

        private static List<double> Cluster(IEnumerable<double> values, double tolerance)
        {
            var result = new List<double>();
            foreach (double value in values.OrderBy(v => v))
            {
                if (result.Count == 0 || value - result[result.Count - 1] > tolerance) result.Add(value);
            }
            return result;
        }

        private static List<Frag> RuledTables(PdfPageContent page)
        {
            const double Touch = 2.5;
            var tables = new List<Frag>();
            List<PdfLineMark> marks = page.Lines;
            if (marks.Count < 4 || marks.Count > 20000) return tables;

            var parent = Enumerable.Range(0, marks.Count).ToArray();
            Func<int, int> find = null;
            find = i =>
            {
                while (parent[i] != i)
                {
                    parent[i] = parent[parent[i]];
                    i = parent[i];
                }
                return i;
            };
            for (int i = 0; i < marks.Count; i++)
            {
                if (!marks[i].Horizontal) continue;
                for (int k = 0; k < marks.Count; k++)
                {
                    if (marks[k].Horizontal) continue;
                    bool touches = marks[k].Position >= marks[i].From - Touch && marks[k].Position <= marks[i].To + Touch
                        && marks[i].Position >= marks[k].From - Touch && marks[i].Position <= marks[k].To + Touch;
                    if (touches) parent[find(i)] = find(k);
                }
            }

            foreach (IGrouping<int, int> group in Enumerable.Range(0, marks.Count).GroupBy(find))
            {
                List<PdfLineMark> lines = group.Select(i => marks[i]).ToList();
                List<PdfLineMark> across = lines.Where(l => l.Horizontal).ToList();
                List<PdfLineMark> down = lines.Where(l => !l.Horizontal).ToList();
                if (across.Count == 0 || down.Count == 0) continue;

                // Borders may be left open at the sides, so the far ends of the lines count as edges as well.
                List<double> xs = Cluster(down.Select(l => l.Position).Concat(across.Select(l => l.From)).Concat(across.Select(l => l.To)), Touch);
                List<double> ys = Cluster(across.Select(l => l.Position).Concat(down.Select(l => l.From)).Concat(down.Select(l => l.To)), Touch);
                if (xs.Count < 3 || ys.Count < 3) continue;

                Frag table = BuildGrid(page, xs, ys);
                if (table != null) tables.Add(table);
            }
            return tables;
        }

        private static Frag BuildGrid(PdfPageContent page, List<double> xs, List<double> ys)
        {
            int columns = xs.Count - 1, rows = ys.Count - 1;
            var cells = new List<PdfTextItem>[rows, columns];
            var used = new List<PdfTextItem>();
            foreach (PdfTextItem item in page.Text)
            {
                double midY = item.Y - item.Size * 0.3;
                double startX = item.X + Math.Min(1, (item.EndX - item.X) / 2);
                if (startX < xs[0] || startX > xs[columns] || midY < ys[0] || midY > ys[rows]) continue;
                int column = xs.FindLastIndex(x => x <= startX);
                int row = ys.FindLastIndex(y => y <= midY);
                column = Math.Min(column, columns - 1);
                row = Math.Min(row, rows - 1);
                if (cells[row, column] == null) cells[row, column] = new List<PdfTextItem>();
                cells[row, column].Add(item);
                used.Add(item);
            }

            int filledColumns = Enumerable.Range(0, columns).Count(c => Enumerable.Range(0, rows).Any(r => cells[r, c] != null));
            int filledRows = Enumerable.Range(0, rows).Count(r => Enumerable.Range(0, columns).Any(c => cells[r, c] != null));
            if (filledColumns < 2 || filledRows < 2) return null;

            var result = new List<List<string>>();
            for (int r = 0; r < rows; r++)
            {
                var row = new List<string>();
                for (int c = 0; c < columns; c++) row.Add(cells[r, c] == null ? "" : CellMarkdown(cells[r, c]));
                result.Add(row);
            }
            foreach (PdfTextItem item in used) page.Text.Remove(item);
            return new Frag { Table = result, L = xs[0], R = xs[columns], T = ys[0], B = ys[rows], Size = 0 };
        }

        private static string CellMarkdown(List<PdfTextItem> items)
        {
            var inline = new MdInline();
            bool first = true;
            foreach (List<PdfTextItem> line in Lines(items))
            {
                if (!first) inline.Add(" ");
                first = false;
                foreach (TextRun run in LineRuns(line)) inline.Add(run.Text, run.Bold, run.Italic, false);
            }
            return inline.Render(false);
        }

        // ---- Reading order ----

        // Bands are stacked strips of the page. Neighbouring bands split by the same gutter form a set of
        // columns, which are read one after the other. Anything else is read from top to bottom.
        private static List<List<Frag>> ReadingOrder(List<Frag> frags)
        {
            var leaves = new List<List<Frag>>();
            if (frags.Count == 0) return leaves;

            double typical = Median(frags.Where(f => f.IsText).Select(f => f.Size), 10);
            var bands = new List<List<Frag>>();
            List<Frag> band = null;
            double bottom = 0;
            foreach (Frag frag in frags.OrderBy(f => f.T))
            {
                if (band == null || frag.T > bottom + typical * 0.8)
                {
                    band = new List<Frag>();
                    bands.Add(band);
                    bottom = frag.B;
                }
                band.Add(frag);
                bottom = Math.Max(bottom, frag.B);
            }

            var gutters = bands.Select(b => Gutter(b, typical)).ToList();
            int index = 0;
            while (index < bands.Count)
            {
                double[] gutter = gutters[index];
                int end = index;
                var group = new List<Frag>(bands[index]);
                while (gutter != null && end + 1 < bands.Count)
                {
                    double[] next = gutters[end + 1];
                    if (next == null || next[1] < gutter[0] || next[0] > gutter[1]) break;
                    gutter = new[] { Math.Max(gutter[0], next[0]), Math.Min(gutter[1], next[1]) };
                    group.AddRange(bands[++end]);
                }

                if (gutter != null && IsColumnPair(group, gutter))
                {
                    double middle = (gutter[0] + gutter[1]) / 2;
                    leaves.AddRange(ReadingOrder(group.Where(f => f.R <= middle).ToList()));
                    leaves.AddRange(ReadingOrder(group.Where(f => f.R > middle).ToList()));
                }
                else
                {
                    leaves.Add(group.OrderBy(f => f.IsText ? f.Base : f.T).ThenBy(f => f.L).ToList());
                }
                index = end + 1;
            }
            return leaves;
        }

        private static double[] Gutter(List<Frag> band, double typical)
        {
            double left = band.Min(f => f.L), right = band.Max(f => f.R);
            double width = right - left;
            if (width <= 0) return null;

            double[] best = null;
            double reach = left;
            foreach (Frag frag in band.OrderBy(f => f.L))
            {
                double gap = frag.L - reach;
                bool centered = reach - left > width * 0.15 && right - frag.L > width * 0.15;
                if (gap >= typical * 0.8 && centered && (best == null || gap > best[1] - best[0])) best = new[] { reach, frag.L };
                reach = Math.Max(reach, frag.R);
            }
            return best;
        }

        // Real text columns have mostly full lines, while the gaps in a table have short pieces beside them.
        private static bool IsColumnPair(List<Frag> group, double[] gutter)
        {
            double middle = (gutter[0] + gutter[1]) / 2;
            return IsProseSide(group.Where(f => f.R <= middle).ToList()) && IsProseSide(group.Where(f => f.R > middle).ToList());
        }

        private static bool IsProseSide(List<Frag> side)
        {
            List<Frag> text = side.Where(f => f.IsText).ToList();
            if (text.Count < 3) return false;
            double width = text.Max(f => f.R) - text.Min(f => f.L);
            return text.Count(f => f.R - f.L >= width * 0.6) >= text.Count * 0.5;
        }

        private static double Median(IEnumerable<double> values, double fallback)
        {
            List<double> sorted = values.OrderBy(v => v).ToList();
            return sorted.Count == 0 ? fallback : sorted[sorted.Count / 2];
        }

        // ---- Lines to paragraphs ----

        private void LayoutLeaf(List<Frag> leaf, int pageIndex, bool ocr, List<Para> output)
        {
            int i = 0;
            while (i < leaf.Count)
            {
                Frag frag = leaf[i];
                if (frag.Table != null)
                {
                    output.Add(new Para { Kind = ParaKind.Table, Table = frag.Table, Ocr = ocr });
                    i++;
                }
                else if (frag.Image != null)
                {
                    AddImage(frag, pageIndex, output);
                    i++;
                }
                else
                {
                    int stop = i;
                    while (stop < leaf.Count && leaf[stop].IsText) stop++;
                    Paragraphs(leaf.GetRange(i, stop - i), ocr, output);
                    i = stop;
                }
            }
        }

        private void AddImage(Frag frag, int pageIndex, List<Para> output)
        {
            byte[] png = RenderRegion((uint)pageIndex, _pages[pageIndex], frag.L, frag.T, frag.R, frag.B);
            if (png != null) output.Add(new Para { Kind = ParaKind.Image, Markdown = MdImages.Markdown(null, _images.Save(png, "png")) });
        }

        private static void Paragraphs(List<Frag> frags, bool ocr, List<Para> output)
        {
            // Pieces on the same baseline are one visual line.
            var lines = new List<List<Frag>>();
            foreach (Frag frag in frags.OrderBy(f => f.Base))
            {
                List<Frag> last = lines.LastOrDefault();
                if (last != null && Math.Abs(frag.Base - last[0].Base) <= 0.45 * Math.Max(frag.Size, last[0].Size)) last.Add(frag);
                else lines.Add(new List<Frag> { frag });
            }
            foreach (List<Frag> line in lines) line.Sort((a, b) => a.L.CompareTo(b.L));

            double regionRight = frags.Max(f => f.R);
            double regionLeft = frags.Min(f => f.L);
            Para para = null;
            List<Frag> previous = null;
            double lastGap = 0;
            for (int i = 0; i < lines.Count; i++)
            {
                List<List<string>> table;
                int end = AlignedTable(lines, i, out table);
                if (end > i)
                {
                    output.Add(new Para { Kind = ParaKind.Table, Table = table, Ocr = ocr });
                    para = null;
                    previous = null;
                    i = end - 1;
                    continue;
                }

                List<Frag> line = lines[i];
                List<TextRun> runs = new List<TextRun>();
                foreach (Frag frag in line)
                {
                    if (runs.Count > 0) runs.Add(new TextRun { Text = " " });
                    runs.AddRange(LineRuns(frag.Items));
                }
                string text = string.Concat(runs.Select(r => r.Text));
                double size = line.Max(f => f.Size);
                double left = line[0].L, right = line[line.Count - 1].R;

                Match bullet = BulletStart.Match(text);
                Match number = NumberStart.Match(text);
                bool otherList = OtherListStart.IsMatch(text);
                bool startsItem = bullet.Success || number.Success || otherList;

                double gap = previous == null ? 0 : line[0].Base - previous[0].Base;
                bool continues = para != null && previous != null && !startsItem
                    && Math.Abs(size - para.Size) <= para.Size * (ocr ? 0.25 : 0.1)
                    && gap <= (lastGap > 0 ? lastGap * 1.25 : size * 1.75)
                    && previous[previous.Count - 1].R >= regionRight - size * 3
                    && (para.IsList || left - previous[0].L <= size * 0.8);

                if (continues)
                {
                    AppendLine(para, runs);
                    para.Lines++;
                    para.FilledLast = right >= regionRight - size * 3;
                    lastGap = gap;
                }
                else
                {
                    para = new Para { Size = size, Lines = 1, Left = left, Ocr = ocr, FilledLast = right >= regionRight - size * 3 };
                    if (bullet.Success)
                    {
                        para.IsList = true;
                        para.ListLevel = Math.Max(0, Math.Min(3, (int)Math.Round((left - regionLeft) / Math.Max(1, size * 1.5))));
                        DropLeading(runs, bullet.Length);
                    }
                    else if (number.Success)
                    {
                        para.IsList = true;
                        para.Ordered = true;
                        para.Number = int.Parse(number.Groups[1].Value, CultureInfo.InvariantCulture);
                        para.ListLevel = Math.Max(0, Math.Min(3, (int)Math.Round((left - regionLeft) / Math.Max(1, size * 1.5))));
                        DropLeading(runs, number.Length);
                    }
                    para.Runs = runs;
                    output.Add(para);
                    lastGap = 0;
                }
                para.Right = Math.Max(para.Right, right);
                previous = line;
            }
        }

        private static List<TextRun> LineRuns(List<PdfTextItem> items)
        {
            var runs = new List<TextRun>();
            PdfTextItem previous = null;
            foreach (PdfTextItem item in items)
            {
                string text = item.Text;
                if (previous != null && item.X - previous.EndX > 0.18 * Math.Max(item.Size, previous.Size)
                    && !previous.Text.EndsWith(" ", StringComparison.Ordinal) && !text.StartsWith(" ", StringComparison.Ordinal))
                {
                    text = " " + text;
                }
                TextRun last = runs.LastOrDefault();
                if (last != null && last.Bold == item.Bold && last.Italic == item.Italic) last.Text += text;
                else runs.Add(new TextRun { Text = text, Bold = item.Bold, Italic = item.Italic });
                previous = item;
            }
            return runs;
        }

        private static void AppendLine(Para para, List<TextRun> runs)
        {
            TextRun last = para.Runs.LastOrDefault();
            string firstText = runs.Count > 0 ? runs[0].Text : "";
            if (last != null && firstText.Length > 0)
            {
                string tail = last.Text.TrimEnd();
                char head = firstText[0];
                // Ordinal, because culture-aware comparison ignores the soft hyphen and would always match it.
                bool hyphen = tail.EndsWith("-", StringComparison.Ordinal) || tail.EndsWith("­", StringComparison.Ordinal);
                if (hyphen && tail.Length > 1 && char.IsLetter(tail[tail.Length - 2]) && tail[tail.Length - 2] < 0x250 && char.IsLower(head))
                {
                    last.Text = tail.Substring(0, tail.Length - 1);
                }
                else
                {
                    bool noSpace = tail.Length > 0 && IsSpacelessScript(tail[tail.Length - 1]) && IsSpacelessScript(head);
                    last.Text = tail + (noSpace ? "" : " ");
                }
            }
            para.Runs.AddRange(runs);
        }

        // Chinese and Japanese don't put spaces between words, but Korean does.
        private static bool IsSpacelessScript(char c)
        {
            return (c >= 0x3040 && c <= 0x30FF) || (c >= 0x3400 && c <= 0x9FFF);
        }

        private static void DropLeading(List<TextRun> runs, int count)
        {
            while (count > 0 && runs.Count > 0)
            {
                if (runs[0].Text.Length <= count)
                {
                    count -= runs[0].Text.Length;
                    runs.RemoveAt(0);
                }
                else
                {
                    runs[0].Text = runs[0].Text.Substring(count);
                    count = 0;
                }
            }
        }

        // Tables without lines show up as several lines whose pieces share the same columns.
        private static int AlignedTable(List<List<Frag>> lines, int start, out List<List<string>> rows)
        {
            rows = null;
            if (lines[start].Count < 2) return start;

            int end = start;
            int multi = 0;
            while (end < lines.Count)
            {
                if (lines[end].Count >= 2) multi++;
                else if (end + 1 >= lines.Count || lines[end + 1].Count < 2) break;
                end++;
            }
            if (multi < 2) return start;

            var columns = new List<double[]>();
            foreach (Frag frag in lines.Skip(start).Take(end - start).Where(l => l.Count >= 2).SelectMany(l => l).OrderBy(f => f.L))
            {
                double[] last = columns.LastOrDefault();
                if (last != null && frag.L <= last[1]) last[1] = Math.Max(last[1], frag.R);
                else columns.Add(new[] { frag.L, frag.R });
            }
            if (columns.Count < 2 || (multi < 3 && columns.Count < 3)) return start;

            rows = new List<List<string>>();
            for (int i = start; i < end; i++)
            {
                if (lines[i].Count == 1 && rows.Count > 0)
                {
                    int index = ColumnOf(columns, lines[i][0]);
                    List<string> row = rows[rows.Count - 1];
                    string text = CellMarkdown(lines[i][0].Items);
                    row[index] = row[index].Length == 0 ? text : row[index] + " " + text;
                    continue;
                }
                var cells = columns.Select(c => "").ToList();
                foreach (Frag frag in lines[i])
                {
                    int index = ColumnOf(columns, frag);
                    string text = CellMarkdown(frag.Items);
                    cells[index] = cells[index].Length == 0 ? text : cells[index] + " " + text;
                }
                rows.Add(cells);
            }
            return end;
        }

        private static int ColumnOf(List<double[]> columns, Frag frag)
        {
            int index = columns.FindLastIndex(c => c[0] <= frag.L + 1);
            return Math.Max(0, index);
        }

        // ---- Writing ----

        private void Emit(List<Para> paras)
        {
            double body = BodySize(paras.Where(p => !p.Ocr));
            double ocrBody = BodySize(paras.Where(p => p.Ocr));
            List<double> headingSizes = paras
                .Where(p => p.Kind == ParaKind.Text && !p.Ocr && IsHeading(p, body, ocrBody))
                .Select(p => Math.Round(p.Size * 2) / 2).Distinct().OrderByDescending(s => s).ToList();

            for (int i = 0; i < paras.Count; i++)
            {
                Para para = paras[i];
                switch (para.Kind)
                {
                    case ParaKind.Table: _doc.AddTable(para.Table); continue;
                    case ParaKind.Image:
                    case ParaKind.Raw: _doc.AddRaw(para.Markdown); continue;
                }

                bool heading = IsHeading(para, body, ocrBody);

                // A sentence cut by a page break continues on the next page.
                while (!heading && !para.IsList && i + 1 < paras.Count)
                {
                    Para next = paras[i + 1];
                    string tail = string.Concat(para.Runs.Select(r => r.Text)).TrimEnd();
                    bool joins = next.Kind == ParaKind.Text && next.Page == para.Page + 1 && !next.IsList && para.FilledLast
                        && Math.Abs(next.Size - para.Size) <= para.Size * 0.1 && !SentenceEnd.IsMatch(tail)
                        && !IsHeading(next, body, ocrBody) && !OtherListStart.IsMatch(string.Concat(next.Runs.Select(r => r.Text)));
                    if (!joins) break;
                    AppendLine(para, next.Runs);
                    para.Page = next.Page;
                    para.FilledLast = next.FilledLast;
                    i++;
                }

                var inline = new MdInline();
                if (heading && para.Ordered) inline.Add(para.Number + ". ");
                foreach (TextRun run in para.Runs) inline.Add(run.Text, run.Bold, run.Italic, false);
                if (heading)
                {
                    int level = para.Ocr ? 2 : headingSizes.IndexOf(Math.Round(para.Size * 2) / 2) + 1;
                    _doc.AddHeading(Math.Max(1, Math.Min(4, level)), inline);
                }
                else if (para.IsList)
                {
                    _doc.AddListItem(para.ListLevel, para.Ordered, para.Number, inline);
                }
                else
                {
                    _doc.AddParagraph(inline);
                }
            }
        }

        private static double BodySize(IEnumerable<Para> paras)
        {
            var weights = new Dictionary<double, int>();
            foreach (Para para in paras.Where(p => p.Kind == ParaKind.Text && p.Size > 0))
            {
                double size = Math.Round(para.Size * 2) / 2;
                int weight;
                weights.TryGetValue(size, out weight);
                weights[size] = weight + para.Runs.Sum(r => r.Text.Length);
            }
            return weights.Count == 0 ? 0 : weights.OrderByDescending(w => w.Value).First().Key;
        }

        private static bool IsHeading(Para para, double body, double ocrBody)
        {
            if (para.Kind != ParaKind.Text || (para.IsList && !para.Ordered) || para.Lines > 3) return false;
            int length = para.Runs.Sum(r => r.Text.Length);
            if (length == 0 || length > 150) return false;
            double reference = para.Ocr ? ocrBody : body;
            return reference > 0 && para.Size >= reference * (para.Ocr ? 1.6 : 1.15);
        }
    }
}
