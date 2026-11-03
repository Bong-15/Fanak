using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media.Imaging;
using System.Xml;

namespace Fanak
{
    // Plain text formats (.txt, .csv, .tsv, .html, .rtf). None of these need another program to read.
    internal static class TextDocs
    {
        private const int MaxExcelRows = 1048576;
        private const int MaxExcelColumns = 16384;
        private const int MaxExcelCellLength = 32767;
        private const int MaxPreciseDigits = 15;
        private const int ColumnWidthSampleRows = 1000;
        private const int MinColumnWidth = 8;
        private const int MaxColumnWidth = 60;

        public static readonly string[] DelimitedExts = { "csv", "tsv" };
        public static readonly string[] WordOpensExts = { "txt", "html", "htm", "rtf" };

        private static readonly Regex PlainNumber = new Regex(@"^-?(?:0|[1-9]\d*)(?:\.\d+)?$");
        private static readonly Regex NotXml = new Regex(@"[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]");
        private static readonly Regex HtmlCharset = new Regex(@"<meta\b[^>]*charset\s*=\s*[""']?([\w-]+)", RegexOptions.IgnoreCase);
        private static readonly byte[] Utf8Mark = { 0xEF, 0xBB, 0xBF };

        public static bool IsTextExt(string ext)
        {
            return DelimitedExts.Contains(ext) || WordOpensExts.Contains(ext);
        }

        public static void Read(string source, string ext, MdDoc doc, MdImages images)
        {
            byte[] data = HangulFile.ReadShared(source);
            if (DelimitedExts.Contains(ext)) ReadDelimited(source, doc);
            else if (ext == "txt") ReadPlain(Decode(data), doc);
            else if (ext == "rtf") new RtfReader(doc, images).Run(data);
            else new HtmlReader(doc, images, Path.GetDirectoryName(source)).Run(DecodeHtml(data));
        }

        // Korean text files are UTF-8 or CP949, usually with nothing to say which. Bytes that break UTF-8's strict
        // rules almost never occur by chance, so this can tell them apart.
        public static string Decode(byte[] data)
        {
            if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) return new UTF8Encoding(false).GetString(data, 3, data.Length - 3);
            if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xFE) return Encoding.Unicode.GetString(data, 2, data.Length - 2);
            if (data.Length >= 2 && data[0] == 0xFE && data[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(data, 2, data.Length - 2);
            if (Array.IndexOf(data, (byte)0) >= 0) throw new InvalidOperationException("텍스트 파일이 아니거나 손상된 파일입니다.");
            try
            {
                return new UTF8Encoding(false, true).GetString(data);
            }
            catch (DecoderFallbackException)
            {
                return Encoding.GetEncoding(949).GetString(data);
            }
        }

        // A page declares its own encoding, which wins over a guess unless there is a byte order mark.
        public static string DecodeHtml(byte[] data)
        {
            bool marked = data.Length >= 2 && (data[0] == 0xFF && data[1] == 0xFE || data[0] == 0xFE && data[1] == 0xFF || data[0] == 0xEF && data[1] == 0xBB);
            if (!marked && Array.IndexOf(data, (byte)0) < 0)
            {
                string head = Encoding.GetEncoding(28591).GetString(data, 0, Math.Min(data.Length, 4096));
                Match declared = HtmlCharset.Match(head);
                if (declared.Success)
                {
                    try { return Encoding.GetEncoding(declared.Groups[1].Value).GetString(data); }
                    catch (ArgumentException) { }
                }
            }
            return Decode(data);
        }

        // ---- .txt ----

        private static void ReadPlain(string text, MdDoc doc)
        {
            var lines = new List<string>();
            foreach (string raw in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
            {
                if (raw.Trim().Length > 0)
                {
                    lines.Add(raw);
                    continue;
                }
                doc.AddLines(lines);
                lines.Clear();
            }
            doc.AddLines(lines);
        }

        // Word guesses a text file's encoding, and guesses wrong on UTF-8 without a mark. Writing one settles it.
        public static byte[] PrepareForWord(string source, string ext)
        {
            byte[] data = HangulFile.ReadShared(source);
            string text = ext == "txt" ? Decode(data) : CleanHtml(DecodeHtml(data), Path.GetDirectoryName(source));
            return Utf8Mark.Concat(new UTF8Encoding(false).GetBytes(text)).ToArray();
        }

        // ---- .csv, .tsv ----

        private static void ReadDelimited(string source, MdDoc doc)
        {
            List<List<string>> rows = ReadRows(source);
            doc.AddTable(rows.Select(r => r.Select(MdInline.Escape).ToList()).ToList());
        }

        private static List<List<string>> ReadRows(string source)
        {
            string text = Decode(HangulFile.ReadShared(source));
            bool tabs = Path.GetExtension(source).Equals(".tsv", StringComparison.OrdinalIgnoreCase);
            return ParseDelimited(text, tabs ? '\t' : GuessDelimiter(text));
        }

        // Where a comma is the decimal point, spreadsheet programs write semicolons instead, whatever the extension says.
        private static char GuessDelimiter(string text)
        {
            int end = text.IndexOfAny(new[] { '\r', '\n' });
            string first = end < 0 ? text : text.Substring(0, end);
            first = Regex.Replace(first, "\"[^\"]*\"", "");
            return new[] { ',', ';', '\t' }.OrderByDescending(d => first.Count(c => c == d)).First();
        }

        private static List<List<string>> ParseDelimited(string text, char delimiter)
        {
            var rows = new List<List<string>>();
            var row = new List<string>();
            var cell = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                bool crlf = c == '\r' && i + 1 < text.Length && text[i + 1] == '\n';
                if (quoted)
                {
                    if (c != '"')
                    {
                        if (crlf) continue;
                        cell.Append(c);
                    }
                    else if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        cell.Append('"');
                        i++;
                    }
                    else quoted = false;
                }
                else if (c == '"' && cell.Length == 0) quoted = true;
                else if (c == delimiter)
                {
                    row.Add(cell.ToString());
                    cell.Clear();
                }
                else if (c == '\r' || c == '\n')
                {
                    if (crlf) i++;
                    row.Add(cell.ToString());
                    cell.Clear();
                    rows.Add(row);
                    row = new List<string>();
                }
                else cell.Append(c);
            }
            if (cell.Length > 0 || row.Count > 0)
            {
                row.Add(cell.ToString());
                rows.Add(row);
            }
            return rows;
        }

        public static string ToXlsx(string source)
        {
            List<List<string>> rows = ReadRows(source);
            if (rows.Count == 0) throw new InvalidOperationException("파일에 내용이 없습니다.");
            if (rows.Count > MaxExcelRows || rows.Max(r => r.Count) > MaxExcelColumns)
                throw new NotSupportedException("Excel 한 시트에 담을 수 있는 크기를 넘습니다.");
            if (rows.Any(r => r.Any(c => c.Length > MaxExcelCellLength)))
                throw new NotSupportedException("Excel 한 칸에 담을 수 있는 " + MaxExcelCellLength + "자를 넘는 값이 있습니다.");

            string output = Converter.UniquePath(source, "", "xlsx");
            try
            {
                using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write))
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
                {
                    WriteEntry(zip, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
                    WriteEntry(zip, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                    WriteEntry(zip, "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                    WriteEntry(zip, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                    using (Stream entry = zip.CreateEntry("xl/worksheets/sheet1.xml").Open())
                    {
                        WriteSheet(entry, rows);
                    }
                }
                return output;
            }
            catch
            {
                Converter.TryDelete(output);
                throw;
            }
        }

        private static void WriteEntry(ZipArchive zip, string name, string body)
        {
            byte[] bytes = new UTF8Encoding(false).GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + body);
            using (Stream entry = zip.CreateEntry(name).Open()) entry.Write(bytes, 0, bytes.Length);
        }

        private static void WriteSheet(Stream stream, List<List<string>> rows)
        {
            const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            using (XmlWriter xml = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
            {
                xml.WriteStartElement("worksheet", Main);
                xml.WriteStartElement("cols", Main);
                int columns = rows.Max(r => r.Count);
                for (int c = 0; c < columns; c++)
                {
                    int index = c;
                    int width = rows.Take(ColumnWidthSampleRows).Where(r => index < r.Count).Select(r => DisplayWidth(r[index])).DefaultIfEmpty(0).Max() + 2;
                    string size = Math.Max(MinColumnWidth, Math.Min(MaxColumnWidth, width)).ToString(CultureInfo.InvariantCulture);
                    string number = (c + 1).ToString(CultureInfo.InvariantCulture);
                    xml.WriteStartElement("col", Main);
                    xml.WriteAttributeString("min", number);
                    xml.WriteAttributeString("max", number);
                    xml.WriteAttributeString("width", size);
                    xml.WriteAttributeString("customWidth", "1");
                    xml.WriteEndElement();
                }
                xml.WriteEndElement();

                xml.WriteStartElement("sheetData", Main);
                for (int r = 0; r < rows.Count; r++)
                {
                    string rowNumber = (r + 1).ToString(CultureInfo.InvariantCulture);
                    xml.WriteStartElement("row", Main);
                    xml.WriteAttributeString("r", rowNumber);
                    for (int c = 0; c < rows[r].Count; c++)
                    {
                        if (rows[r][c].Length > 0) WriteCell(xml, Main, ColumnName(c) + rowNumber, rows[r][c]);
                    }
                    xml.WriteEndElement();
                }
                xml.WriteEndElement();
                xml.WriteEndElement();
            }
        }

        // Only an exact number becomes one. Codes such as 007 keep their zeros, and text stays text, so a value
        // beginning with "=" can never run as a formula.
        private static void WriteCell(XmlWriter xml, string ns, string reference, string value)
        {
            xml.WriteStartElement("c", ns);
            xml.WriteAttributeString("r", reference);
            if (PlainNumber.IsMatch(value) && value.Count(char.IsDigit) <= MaxPreciseDigits)
            {
                xml.WriteElementString("v", ns, value);
            }
            else
            {
                xml.WriteAttributeString("t", "inlineStr");
                xml.WriteStartElement("is", ns);
                xml.WriteStartElement("t", ns);
                xml.WriteAttributeString("xml", "space", null, "preserve");
                xml.WriteString(NotXml.Replace(value, ""));
                xml.WriteEndElement();
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }

        private static string ColumnName(int index)
        {
            string name = "";
            for (int n = index + 1; n > 0; n = (n - 1) / 26) name = (char)('A' + (n - 1) % 26) + name;
            return name;
        }

        private static int DisplayWidth(string text)
        {
            return text.Sum(c => c < 0x80 ? 1 : 2);
        }

        // ---- .html for Word ----

        // The manual promises that files stay on this PC, but Word would fetch whatever a page links to.
        // Local pictures still work: <base> points relative links back at the original folder.
        private static string CleanHtml(string html, string folder)
        {
            const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.Singleline;
            html = Regex.Replace(html, @"<(script|iframe|object|frameset)\b.*?</\1\s*>", "", Options);
            html = Regex.Replace(html, @"<(?:embed|base)\b[^>]*>", "", Options);
            html = Regex.Replace(html, @"\ssrcset\s*=\s*(?:""[^""]*""|'[^']*'|[^\s>]+)", "", Options);
            html = Regex.Replace(html, @"@import\b[^;]*;", "", Options);
            html = Regex.Replace(html, @"url\(\s*['""]?(?:[a-z][a-z0-9+.-]*:|//|\\\\)[^)]*\)", "none", Options);
            html = Regex.Replace(html, @"(<meta\b[^>]*charset\s*=\s*[""']?)[\w-]+", "$1utf-8", Options);
            html = Regex.Replace(html, @"<(?!a\b)[a-z][^>]*>", tag => Regex.Replace(tag.Value,
                @"\b(src|srcset|href|poster|background|data)\s*=\s*(?:""\s*(?:[a-z][a-z0-9+.-]*:|//|\\\\)[^""]*""|'\s*(?:[a-z][a-z0-9+.-]*:|//|\\\\)[^']*'|(?:[a-z][a-z0-9+.-]*:|//|\\\\)[^\s>]*)",
                match => Regex.IsMatch(match.Value, @"=\s*[""']?\s*data:", Options) ? match.Value : match.Groups[1].Value + "=\"\"",
                Options), Options);

            string baseTag = "<base href=\"" + WebUtility.HtmlEncode(new Uri(folder.TrimEnd('\\') + "\\").AbsoluteUri) + "\">";
            Match head = Regex.Match(html, @"<head\b[^>]*>", Options);
            return head.Success ? html.Insert(head.Index + head.Length, baseTag) : baseTag + html;
        }
    }

    // Enough of HTML to bring over what a saved page says. It keeps no styling and runs no scripts.
    internal sealed class HtmlReader
    {
        private sealed class OpenList
        {
            public bool Ordered;
            public int Number;
        }

        private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.Singleline;
        private static readonly Regex Token = new Regex(
            @"<!--.*?-->|<(script|style|noscript|head|title|template|svg)\b.*?</\1\s*>|<[!/a-z?](?:[^>""']|""[^""]*""|'[^']*')*>|<[!/a-z?][^>]*>|[^<]+|<",
            Options);
        private static readonly Regex TagName = new Regex(@"^<(/?)([a-z][\w:-]*)", Options);
        private static readonly string[] ImageExts = { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg", ".tif", ".tiff" };

        private readonly MdDoc _doc;
        private readonly MdImages _images;
        private readonly string _folder;
        private readonly Stack<OpenList> _lists = new Stack<OpenList>();
        private MdInline _current = new MdInline();
        private MdInline _beforeLink;
        private string _href;
        private StringBuilder _pre;
        private List<List<string>> _rows;
        private List<string> _row;
        private bool _inCell;
        private int _tableDepth, _heading, _quote, _bold, _italic, _strike;

        public HtmlReader(MdDoc doc, MdImages images, string folder)
        {
            _doc = doc;
            _images = images;
            _folder = folder;
        }

        public void Run(string html)
        {
            foreach (Match token in Token.Matches(html))
            {
                string text = token.Value;
                if (token.Groups[1].Success || text.StartsWith("<!", StringComparison.Ordinal) || text.StartsWith("<?", StringComparison.Ordinal)) continue;
                Match tag = TagName.Match(text);
                if (tag.Success) OnTag(tag.Groups[2].Value.ToLowerInvariant(), tag.Groups[1].Length > 0, text);
                else if (!(text.StartsWith("</", StringComparison.Ordinal) && text.Length > 2)) OnText(WebUtility.HtmlDecode(text));
            }
            EndTable();
            Flush();
        }

        private void OnText(string text)
        {
            if (_pre != null)
            {
                _pre.Append(text);
                return;
            }
            _current.Add(Regex.Replace(text, @"\s+", " "), _bold > 0, _italic > 0, _strike > 0);
        }

        private void OnTag(string name, bool closing, string tag)
        {
            if (_pre != null && name != "pre")
            {
                if (name == "br") _pre.Append('\n');
                return;
            }
            switch (name)
            {
                case "h1": case "h2": case "h3": case "h4": case "h5": case "h6":
                    Flush();
                    _heading = closing ? 0 : name[1] - '0';
                    return;
                case "p": case "div": case "section": case "article": case "header": case "footer": case "nav": case "main":
                case "aside": case "figure": case "figcaption": case "address": case "dl": case "dt": case "dd": case "form": case "center":
                    Flush();
                    return;
                case "br": _current.AddBreak(); return;
                case "hr":
                    Flush();
                    if (_tableDepth == 0) _doc.AddRaw("---");
                    return;
                case "b": case "strong": _bold = Math.Max(0, _bold + (closing ? -1 : 1)); return;
                case "i": case "em": _italic = Math.Max(0, _italic + (closing ? -1 : 1)); return;
                case "s": case "strike": case "del": _strike = Math.Max(0, _strike + (closing ? -1 : 1)); return;
                case "a":
                    EndLink();
                    if (!closing) StartLink(Attr(tag, "href"));
                    return;
                case "img": AddImage(tag); return;
                case "ul": case "ol":
                    Flush();
                    if (!closing) _lists.Push(new OpenList { Ordered = name == "ol" });
                    else if (_lists.Count > 0) _lists.Pop();
                    return;
                case "li":
                    Flush();
                    if (!closing && _lists.Count > 0) _lists.Peek().Number++;
                    return;
                case "blockquote":
                    Flush();
                    _quote = Math.Max(0, _quote + (closing ? -1 : 1));
                    return;
                case "pre":
                    if (!closing)
                    {
                        Flush();
                        _pre = new StringBuilder();
                    }
                    else if (_pre != null) EndPre();
                    return;
                case "table": OnTable(closing); return;
                case "caption":
                    if (closing && _tableDepth == 1)
                    {
                        EndLink();
                        _doc.AddParagraph(_current);
                        _current = new MdInline();
                    }
                    return;
                case "tr":
                    if (_tableDepth != 1) return;
                    EndCell();
                    EndRow();
                    if (!closing) _row = new List<string>();
                    return;
                case "td": case "th":
                    if (_tableDepth != 1) return;
                    EndCell();
                    if (!closing)
                    {
                        if (_row == null) _row = new List<string>();
                        _current = new MdInline();
                        _inCell = true;
                    }
                    return;
            }
        }

        private void OnTable(bool closing)
        {
            // A table inside a table can't exist in Markdown, so its cells just join the outer cell's text.
            if (!closing)
            {
                if (_tableDepth++ == 0)
                {
                    Flush();
                    _rows = new List<List<string>>();
                }
            }
            else if (_tableDepth > 0 && --_tableDepth == 0)
            {
                _tableDepth = 1;
                EndTable();
            }
        }

        private void EndTable()
        {
            if (_tableDepth == 0) return;
            EndCell();
            EndRow();
            _tableDepth = 0;
            _doc.AddTable(_rows);
            _rows = null;
            _current = new MdInline();
        }

        private void EndCell()
        {
            if (!_inCell) return;
            EndLink();
            _row.Add(_current.Render(false));
            _current = new MdInline();
            _inCell = false;
        }

        private void EndRow()
        {
            if (_row != null && _row.Count > 0) _rows.Add(_row);
            _row = null;
        }

        private void EndPre()
        {
            string body = _pre.ToString().Trim('\r', '\n').Replace("\r\n", "\n");
            _pre = null;
            if (body.Length == 0) return;
            string fence = body.Contains("```") ? "````" : "```";
            _doc.AddRaw(fence + "\n" + body + "\n" + fence);
        }

        private void StartLink(string href)
        {
            _beforeLink = _current;
            _href = href;
            _current = new MdInline();
        }

        private void EndLink()
        {
            if (_beforeLink == null) return;
            MdInline outer = _beforeLink;
            string label = _current.Render(false);
            string href = (_href ?? "").Trim().Replace(" ", "%20").Replace(")", "%29");
            _beforeLink = null;
            _current = outer;
            if (label.Length == 0) return;
            bool linkable = href.Length > 0 && !href.StartsWith("#", StringComparison.Ordinal)
                && !Regex.IsMatch(href, @"^(?!https?:|mailto:)[a-z][a-z0-9+.-]*:", RegexOptions.IgnoreCase);
            _current.AddRaw(linkable ? "[" + label + "](" + href + ")" : label);
        }

        private void AddImage(string tag)
        {
            string alt = Attr(tag, "alt");
            string src = (Attr(tag, "src") ?? "").Trim();
            string link = null;
            if (Regex.IsMatch(src, "^(?:https?:)?//", RegexOptions.IgnoreCase)) link = src.Replace(" ", "%20").Replace(")", "%29");
            else if (src.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase)) link = SaveDataImage(src);
            else if (src.Length > 0) link = SaveLocalImage(src);

            if (link != null) _current.AddRaw(MdImages.Markdown(alt, link));
            else if (!string.IsNullOrWhiteSpace(alt)) _current.Add(alt, _bold > 0, _italic > 0, _strike > 0);
        }

        private string SaveDataImage(string src)
        {
            int comma = src.IndexOf(',');
            if (comma < 0 || src.IndexOf(";base64", 0, comma, StringComparison.OrdinalIgnoreCase) < 0) return null;
            Match type = Regex.Match(src, @"^data:image/(\w+)", RegexOptions.IgnoreCase);
            try
            {
                return _images.SaveEmbedded(System.Convert.FromBase64String(src.Substring(comma + 1)), "image." + type.Groups[1].Value);
            }
            catch (FormatException ex)
            {
                Log.Write("문서 속 그림을 읽지 못함: " + ex.Message);
                return null;
            }
        }

        private string SaveLocalImage(string src)
        {
            string relative = Uri.UnescapeDataString(src.Split('?', '#')[0]).Replace('/', '\\');
            if (Regex.IsMatch(relative, @"^(?:[a-z][a-z0-9+.-]*:|\\\\)", RegexOptions.IgnoreCase)) return null;
            string path;
            try { path = Path.GetFullPath(Path.Combine(_folder, relative.TrimStart('\\'))); }
            catch (ArgumentException) { return null; }
            catch (NotSupportedException) { return null; }
            catch (PathTooLongException) { return null; }
            if (!File.Exists(path) || !ImageExts.Contains(Path.GetExtension(path).ToLowerInvariant())) return null;
            return _images.SaveEmbedded(File.ReadAllBytes(path), Path.GetFileName(path));
        }

        // Ends the block being collected and sends it to the document according to where it sits.
        private void Flush()
        {
            EndLink();
            if (_tableDepth > 0)
            {
                _current.Add(" ");
                return;
            }
            MdInline text = _current;
            _current = new MdInline();
            if (_heading > 0) _doc.AddHeading(_heading, text);
            else if (_quote > 0) _doc.AddQuote(text.Render(false));
            else if (_lists.Count > 0)
            {
                OpenList list = _lists.Peek();
                _doc.AddListItem(_lists.Count - 1, list.Ordered, Math.Max(1, list.Number), text);
            }
            else _doc.AddParagraph(text);
        }

        private static string Attr(string tag, string name)
        {
            Match match = Regex.Match(tag, @"\s" + name + @"\s*=\s*(?:""([^""]*)""|'([^']*)'|([^\s>]+))", Options);
            if (!match.Success) return null;
            string value = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Success ? match.Groups[2].Value : match.Groups[3].Value;
            return WebUtility.HtmlDecode(value);
        }
    }

    // WPF's rich text box already reads RTF, so this walks the document it builds. Needs an STA thread.
    internal sealed class RtfReader
    {
        private readonly MdDoc _doc;
        private readonly MdImages _images;

        public RtfReader(MdDoc doc, MdImages images)
        {
            _doc = doc;
            _images = images;
        }

        public void Run(byte[] data)
        {
            if (data.Length < 5 || Encoding.ASCII.GetString(data, 0, 5) != "{\\rtf")
                throw new InvalidOperationException("RTF 문서가 아니거나 손상된 파일입니다.");
            var flow = new FlowDocument();
            using (var stream = new MemoryStream(data))
            {
                try
                {
                    new TextRange(flow.ContentStart, flow.ContentEnd).Load(stream, DataFormats.Rtf);
                }
                catch (ArgumentException ex)
                {
                    throw new InvalidOperationException("RTF 문서를 읽지 못했습니다. 손상된 파일일 수 있습니다.", ex);
                }
            }
            ReadBlocks(flow.Blocks, 0);
        }

        private void ReadBlocks(BlockCollection blocks, int listLevel)
        {
            foreach (Block block in blocks)
            {
                var paragraph = block as Paragraph;
                var list = block as List;
                var table = block as Table;
                var section = block as Section;
                if (paragraph != null) _doc.AddParagraph(InlinesOf(paragraph));
                else if (list != null) ReadList(list, listLevel);
                else if (table != null) ReadTable(table);
                else if (section != null) ReadBlocks(section.Blocks, listLevel);
            }
        }

        private void ReadList(List list, int level)
        {
            bool ordered = list.MarkerStyle != TextMarkerStyle.Disc && list.MarkerStyle != TextMarkerStyle.Circle
                && list.MarkerStyle != TextMarkerStyle.Square && list.MarkerStyle != TextMarkerStyle.Box && list.MarkerStyle != TextMarkerStyle.None;
            int number = 0;
            foreach (ListItem item in list.ListItems)
            {
                number++;
                foreach (Block block in item.Blocks)
                {
                    var paragraph = block as Paragraph;
                    var nested = block as List;
                    if (paragraph != null) _doc.AddListItem(level, ordered, number, InlinesOf(paragraph));
                    else if (nested != null) ReadList(nested, level + 1);
                }
            }
        }

        private void ReadTable(Table table)
        {
            var rows = new List<List<string>>();
            foreach (TableRowGroup group in table.RowGroups)
            {
                foreach (TableRow row in group.Rows)
                {
                    rows.Add(row.Cells.Select(cell => string.Join("  \n", cell.Blocks.OfType<Paragraph>()
                        .Select(p => InlinesOf(p).Render(false)).Where(t => t.Length > 0))).ToList());
                }
            }
            _doc.AddTable(rows);
        }

        private MdInline InlinesOf(Paragraph paragraph)
        {
            var result = new MdInline();
            AddInlines(paragraph.Inlines, result, false);
            return result;
        }

        private void AddInlines(InlineCollection inlines, MdInline output, bool strike)
        {
            foreach (Inline inline in inlines)
            {
                bool struck = strike || (inline.TextDecorations != null && inline.TextDecorations.Any(d => d.Location == TextDecorationLocation.Strikethrough));
                var run = inline as Run;
                var link = inline as Hyperlink;
                var span = inline as Span;
                var picture = inline as InlineUIContainer;
                if (run != null) output.Add(run.Text, run.FontWeight.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight(), run.FontStyle == FontStyles.Italic, struck);
                else if (inline is LineBreak) output.AddBreak();
                else if (link != null && link.NavigateUri != null)
                {
                    var label = new MdInline();
                    AddInlines(link.Inlines, label, struck);
                    string text = label.Render(false);
                    if (text.Length > 0) output.AddRaw("[" + text + "](" + link.NavigateUri.AbsoluteUri.Replace(")", "%29") + ")");
                }
                else if (span != null) AddInlines(span.Inlines, output, struck);
                else if (picture != null) AddPicture(picture, output);
            }
        }

        private void AddPicture(InlineUIContainer picture, MdInline output)
        {
            var image = picture.Child as Image;
            var bitmap = image == null ? null : image.Source as BitmapSource;
            if (bitmap == null) return;
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bitmap));
            using (var stream = new MemoryStream())
            {
                encoder.Save(stream);
                output.AddRaw(MdImages.Markdown(null, _images.Save(stream.ToArray(), "png")));
            }
        }
    }
}
