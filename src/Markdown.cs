using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Fanak
{
    // Entry point for "MD": every reader fills an MdDoc, so they all produce the same style of Markdown.
    internal static class MarkdownConverter
    {
        public static string Convert(string source)
        {
            string output = Converter.UniquePath(source, "", "md");
            var images = new MdImages(output);
            var doc = new MdDoc();
            bool created = false;
            try
            {
                string ext = Path.GetExtension(source).TrimStart('.').ToLowerInvariant();
                if (ext == "pdf")
                {
                    // Windows.Data.Pdf blocks on async calls, which must not happen on the STA conversion thread.
                    Task.Run(() => new PdfToMd(source, doc, images).Run()).GetAwaiter().GetResult();
                }
                else if (HangulFile.Exts.Contains(ext))
                {
                    HangulFile.Read(source, doc, images);
                }
                else if (TextDocs.IsTextExt(ext))
                {
                    TextDocs.Read(source, ext, doc, images);
                }
                else
                {
                    ReadPackage(source, ext, doc, images);
                }

                string text = doc.Render();
                if (string.IsNullOrWhiteSpace(text)) throw new InvalidOperationException("문서에서 내용을 찾지 못했습니다.");

                byte[] bytes = new UTF8Encoding(false).GetBytes(text);
                using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write))
                {
                    created = true;
                    stream.Write(bytes, 0, bytes.Length);
                }
                return output;
            }
            catch
            {
                images.Remove();
                if (created) Converter.TryDelete(output);
                throw;
            }
        }

        private static void ReadPackage(string source, string ext, MdDoc doc, MdImages images)
        {
            try
            {
                using (var package = new DocPackage(source))
                {
                    switch (ext)
                    {
                        case "docx":
                        case "docm": new DocxReader(package, doc, images).Run(); break;
                        case "pptx":
                        case "pptm": new PptxReader(package, doc, images).Run(); break;
                        case "xlsx":
                        case "xlsm": new XlsxReader(package, doc).Run(); break;
                        default: throw new NotSupportedException(ext + " 형식은 지원하지 않습니다.");
                    }
                }
            }
            catch (InvalidDataException ex)
            {
                throw new InvalidOperationException("파일을 열 수 없습니다. 손상됐거나 암호가 걸려 있을 수 있습니다.", ex);
            }
            catch (System.Xml.XmlException ex)
            {
                throw new InvalidOperationException("파일을 읽을 수 없습니다. 손상된 문서일 수 있습니다.", ex);
            }
        }
    }

    // A run of text that shares one format. Raw pieces (links, images) are already Markdown.
    internal sealed class MdInline
    {
        private sealed class Piece
        {
            public string Text;
            public bool Raw, Bold, Italic, Strike;
        }

        private readonly List<Piece> _pieces = new List<Piece>();

        public void Add(string text, bool bold, bool italic, bool strike)
        {
            if (string.IsNullOrEmpty(text)) return;
            Piece last = _pieces.LastOrDefault();
            if (last != null && !last.Raw && last.Bold == bold && last.Italic == italic && last.Strike == strike)
                last.Text += text;
            else
                _pieces.Add(new Piece { Text = text, Bold = bold, Italic = italic, Strike = strike });
        }

        public void Add(string text)
        {
            Add(text, false, false, false);
        }

        public void AddRaw(string markdown)
        {
            _pieces.Add(new Piece { Text = markdown, Raw = true });
        }

        public void AddBreak()
        {
            AddRaw("  \n");
        }

        public bool IsEmpty
        {
            get { return _pieces.All(p => !p.Raw && string.IsNullOrWhiteSpace(p.Text)) || Render(false).Length == 0; }
        }

        public string PlainText
        {
            get { return string.Concat(_pieces.Where(p => !p.Raw).Select(p => p.Text)); }
        }

        // Headings are already emphasized, so they drop the bold and italic.
        public string Render(bool plain)
        {
            var result = new StringBuilder();
            foreach (Piece piece in _pieces)
            {
                if (piece.Raw)
                {
                    result.Append(piece.Text);
                    continue;
                }
                string text = Regex.Replace(piece.Text.Replace(' ', ' ').Replace('\t', ' '), " {2,}", " ");
                string core = text.Trim();
                if (core.Length == 0)
                {
                    if (result.Length > 0 && result[result.Length - 1] != ' ' && result[result.Length - 1] != '\n') result.Append(' ');
                    continue;
                }

                string mark = "";
                if (!plain && piece.Bold) mark += "**";
                if (!plain && piece.Italic) mark += "*";
                string strike = !plain && piece.Strike ? "~~" : "";
                string closing = new string(mark.Reverse().ToArray());
                if (text[0] == ' ' && result.Length > 0 && result[result.Length - 1] != ' ' && result[result.Length - 1] != '\n') result.Append(' ');
                result.Append(strike).Append(mark).Append(Escape(core)).Append(closing).Append(strike);
                if (text[text.Length - 1] == ' ') result.Append(' ');
            }
            return result.ToString().Trim(' ', '\n');
        }

        public static string Escape(string text)
        {
            var result = new StringBuilder(text.Length + 8);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\\' || c == '`' || c == '*' || c == '[' || c == ']' || c == '<' || c == '~')
                {
                    result.Append('\\');
                }
                else if (c == '_')
                {
                    bool inWord = i > 0 && i + 1 < text.Length && char.IsLetterOrDigit(text[i - 1]) && char.IsLetterOrDigit(text[i + 1]);
                    if (!inWord) result.Append('\\');
                }
                result.Append(c);
            }
            return result.ToString();
        }
    }

    internal sealed class MdDoc
    {
        private static readonly Regex LooksLikeBlock = new Regex(@"^(#{1,6}\s|[-+]\s|>|\d{1,9}[.)]\s|-{3,}\s*$|={3,}\s*$)");

        private sealed class Block
        {
            public string Text;
            public bool ListItem;
            public int Level;
        }

        private readonly List<Block> _blocks = new List<Block>();
        private readonly List<string> _footnotes = new List<string>();

        public void AddHeading(int level, MdInline text)
        {
            string body = text.Render(true).Replace("  \n", " ").Replace("\n", " ");
            if (body.Length == 0) return;
            _blocks.Add(new Block { Text = new string('#', Math.Max(1, Math.Min(6, level))) + " " + body });
        }

        public void AddParagraph(MdInline text)
        {
            string body = text.Render(false);
            if (body.Length == 0) return;
            if (LooksLikeBlock.IsMatch(body)) body = EscapeBlockStart(body);
            _blocks.Add(new Block { Text = body });
        }

        // Lines of plain text, kept as they break. Each one could look like Markdown, not only the first.
        public void AddLines(IEnumerable<string> lines)
        {
            AddRaw(string.Join("  \n", lines.Select(line =>
            {
                string text = MdInline.Escape(line.Trim());
                return LooksLikeBlock.IsMatch(text) ? EscapeBlockStart(text) : text;
            })));
        }

        private static string EscapeBlockStart(string body)
        {
            Match number = Regex.Match(body, @"^\d{1,9}(?=[.)])");
            return number.Success ? body.Insert(number.Length, "\\") : "\\" + body;
        }

        public void AddQuote(string markdown)
        {
            if (markdown.Length == 0) return;
            _blocks.Add(new Block { Text = "> " + markdown.Replace("\n", "\n> ") });
        }

        public void AddRaw(string markdown)
        {
            if (markdown.Length > 0) _blocks.Add(new Block { Text = markdown });
        }

        public void AddListItem(int level, bool ordered, int number, MdInline text)
        {
            string body = text.Render(false);
            if (body.Length == 0) return;
            // Indenting without a parent item would turn the line into a code block.
            Block previous = _blocks.LastOrDefault();
            int deepest = previous != null && previous.ListItem ? previous.Level + 1 : 0;
            level = Math.Max(0, Math.Min(level, deepest));
            string indent = new string(' ', level * 4);
            string marker = ordered ? number + ". " : "- ";
            body = body.Replace("\n", "\n" + indent + new string(' ', marker.Length));
            _blocks.Add(new Block { Text = indent + marker + body, ListItem = true, Level = level });
        }

        public void AddTable(List<List<string>> rows)
        {
            rows = rows.Where(r => r.Any(c => !string.IsNullOrWhiteSpace(c))).ToList();
            if (rows.Count == 0) return;
            int columns = rows.Max(r => r.Count);
            var keep = Enumerable.Range(0, columns)
                .Where(c => rows.Any(r => c < r.Count && !string.IsNullOrWhiteSpace(r[c]))).ToList();

            var table = new StringBuilder();
            for (int i = 0; i < rows.Count; i++)
            {
                table.Append('|');
                foreach (int c in keep) table.Append(' ').Append(c < rows[i].Count ? Cell(rows[i][c]) : "").Append(" |");
                if (i == 0)
                {
                    table.Append("\n|");
                    foreach (int c in keep) table.Append(" --- |");
                }
                if (i + 1 < rows.Count) table.Append('\n');
            }
            _blocks.Add(new Block { Text = table.ToString() });
        }

        private static string Cell(string markdown)
        {
            return markdown.Trim().Replace("|", "\\|").Replace("  \n", "<br>").Replace("\n", " ");
        }

        public void AddFootnote(int number, string markdown)
        {
            _footnotes.Add("[^" + number + "]: " + markdown.Replace("  \n", " ").Replace("\n", " "));
        }

        public string Render()
        {
            var result = new StringBuilder();
            for (int i = 0; i < _blocks.Count; i++)
            {
                if (i > 0) result.Append(_blocks[i].ListItem && _blocks[i - 1].ListItem ? "\n" : "\n\n");
                result.Append(_blocks[i].Text);
            }
            if (_footnotes.Count > 0)
            {
                if (result.Length > 0) result.Append("\n\n");
                result.Append(string.Join("\n", _footnotes));
            }
            if (result.Length > 0) result.Append('\n');
            return result.ToString();
        }
    }

    // Pictures go in a folder beside the .md, created only when there is a picture to put in it.
    internal sealed class MdImages
    {
        private static readonly string[] ViewableExts = { "png", "jpg", "jpeg", "gif", "bmp", "webp", "svg" };

        private readonly string _mdPath;
        private readonly Dictionary<string, string> _saved = new Dictionary<string, string>();
        private string _folder;
        private int _count;

        public MdImages(string mdPath)
        {
            _mdPath = mdPath;
        }

        // Returns the link to use in the document.
        public string Save(byte[] data, string ext)
        {
            string key;
            using (var md5 = MD5.Create()) key = System.Convert.ToBase64String(md5.ComputeHash(data));
            string link;
            if (_saved.TryGetValue(key, out link)) return link;

            if (_folder == null)
            {
                string stem = Path.Combine(Path.GetDirectoryName(_mdPath), Path.GetFileNameWithoutExtension(_mdPath) + "_images");
                _folder = stem;
                for (int n = 1; Directory.Exists(_folder) || File.Exists(_folder); n++) _folder = stem + " (" + n + ")";
                Directory.CreateDirectory(_folder);
            }

            _count++;
            string name = "image" + _count.ToString("D3") + "." + ext;
            File.WriteAllBytes(Path.Combine(_folder, name), data);
            link = Encode(Path.GetFileName(_folder)) + "/" + name;
            _saved[key] = link;
            return link;
        }

        // Pictures from inside a document. Viewers can't show metafiles and TIFF, so those become PNG.
        public string SaveEmbedded(byte[] data, string entryName)
        {
            string ext = Path.GetExtension(entryName).TrimStart('.').ToLowerInvariant();
            if (ViewableExts.Contains(ext)) return Save(data, ext);
            try
            {
                using (var input = new MemoryStream(data))
                using (System.Drawing.Image picture = System.Drawing.Image.FromStream(input))
                using (var bitmap = new System.Drawing.Bitmap(picture))
                using (var output = new MemoryStream())
                {
                    bitmap.Save(output, System.Drawing.Imaging.ImageFormat.Png);
                    return Save(output.ToArray(), "png");
                }
            }
            catch (ArgumentException ex)
            {
                Log.Write("문서 속 그림을 변환하지 못함: " + entryName + " - " + ex.Message);
                return null;
            }
        }

        public static string Markdown(string alt, string link)
        {
            alt = (alt ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            return "![" + MdInline.Escape(alt) + "](" + link + ")";
        }

        private static string Encode(string name)
        {
            return name.Replace("%", "%25").Replace(" ", "%20").Replace("(", "%28").Replace(")", "%29").Replace("#", "%23").Replace("?", "%3F");
        }

        public void Remove()
        {
            if (_folder == null || !Directory.Exists(_folder)) return;
            try { Directory.Delete(_folder, true); }
            catch (IOException ex) { Log.Write("임시 그림 폴더 삭제 실패: " + _folder + " - " + ex.Message); }
            catch (UnauthorizedAccessException ex) { Log.Write("임시 그림 폴더 삭제 실패: " + _folder + " - " + ex.Message); }
        }
    }
}
