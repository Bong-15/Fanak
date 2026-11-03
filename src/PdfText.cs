using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using PdfDict = System.Collections.Generic.Dictionary<string, object>;

namespace Fanak
{
    // Windows.Data.Pdf can only draw pages, so reading the text takes this small PDF parser.
    // It does not handle every PDF; PdfToMd falls back to OCR for what comes out unreadable here.
    internal sealed class PdfName
    {
        public readonly string Value;
        public PdfName(string value) { Value = value; }
    }

    internal sealed class PdfRef
    {
        public readonly int Number;
        public PdfRef(int number) { Number = number; }
    }

    internal sealed class PdfOp
    {
        public readonly string Value;
        public PdfOp(string value) { Value = value; }
    }

    internal sealed class PdfStream
    {
        public PdfDict Dict;
        public byte[] Source;
        public int Start;
    }

    internal sealed class PdfLexer
    {
        public static readonly object Eof = new object();

        private readonly byte[] _d;
        private readonly bool _content;
        public int Pos;

        public PdfLexer(byte[] data, int position, bool content)
        {
            _d = data;
            Pos = position;
            _content = content;
        }

        private static bool IsSpace(byte b)
        {
            return b == 32 || b == 10 || b == 13 || b == 9 || b == 12 || b == 0;
        }

        private static bool IsDelimiter(byte b)
        {
            return b == '(' || b == ')' || b == '<' || b == '>' || b == '[' || b == ']' || b == '{' || b == '}' || b == '/' || b == '%';
        }

        public void SkipSpace()
        {
            while (Pos < _d.Length)
            {
                if (IsSpace(_d[Pos])) Pos++;
                else if (_d[Pos] == '%') { while (Pos < _d.Length && _d[Pos] != 10 && _d[Pos] != 13) Pos++; }
                else break;
            }
        }

        public object Read()
        {
            SkipSpace();
            if (Pos >= _d.Length) return Eof;
            byte c = _d[Pos];
            switch ((char)c)
            {
                case '/': return ReadName();
                case '(': return ReadLiteral();
                case '[': Pos++; return ReadArray();
                case ']': Pos++; return new PdfOp("]");
                case '{': Pos++; return new PdfOp("{");
                case '}': Pos++; return new PdfOp("}");
                case ')': Pos++; return Read();
                case '<':
                    if (Pos + 1 < _d.Length && _d[Pos + 1] == '<') { Pos += 2; return ReadDict(); }
                    return ReadHex();
                case '>':
                    Pos += Pos + 1 < _d.Length && _d[Pos + 1] == '>' ? 2 : 1;
                    return new PdfOp(">>");
            }

            int start = Pos;
            while (Pos < _d.Length && !IsSpace(_d[Pos]) && !IsDelimiter(_d[Pos])) Pos++;
            string token = Encoding.GetEncoding(28591).GetString(_d, start, Pos - start);

            char first = token[0];
            if (char.IsDigit(first) || first == '-' || first == '+' || first == '.')
            {
                double number;
                if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) return 0.0;
                if (!_content && number >= 0 && number == Math.Floor(number))
                {
                    PdfRef reference = TryReadRef((int)number);
                    if (reference != null) return reference;
                }
                return number;
            }
            switch (token)
            {
                case "true": return true;
                case "false": return false;
                case "null": return null;
            }
            return new PdfOp(token);
        }

        private PdfRef TryReadRef(int number)
        {
            int saved = Pos;
            SkipSpace();
            int digits = Pos;
            while (Pos < _d.Length && _d[Pos] >= '0' && _d[Pos] <= '9') Pos++;
            if (Pos > digits)
            {
                SkipSpace();
                if (Pos < _d.Length && _d[Pos] == 'R' && (Pos + 1 >= _d.Length || IsSpace(_d[Pos + 1]) || IsDelimiter(_d[Pos + 1])))
                {
                    Pos++;
                    return new PdfRef(number);
                }
            }
            Pos = saved;
            return null;
        }

        private PdfName ReadName()
        {
            Pos++;
            var name = new StringBuilder();
            while (Pos < _d.Length && !IsSpace(_d[Pos]) && !IsDelimiter(_d[Pos]))
            {
                byte b = _d[Pos++];
                if (b == '#' && Pos + 1 < _d.Length && HexValue(_d[Pos]) >= 0 && HexValue(_d[Pos + 1]) >= 0)
                {
                    b = (byte)(HexValue(_d[Pos]) * 16 + HexValue(_d[Pos + 1]));
                    Pos += 2;
                }
                name.Append((char)b);
            }
            return new PdfName(name.ToString());
        }

        private static int HexValue(byte b)
        {
            if (b >= '0' && b <= '9') return b - '0';
            if (b >= 'a' && b <= 'f') return b - 'a' + 10;
            if (b >= 'A' && b <= 'F') return b - 'A' + 10;
            return -1;
        }

        private byte[] ReadLiteral()
        {
            Pos++;
            var bytes = new List<byte>();
            int depth = 1;
            while (Pos < _d.Length)
            {
                byte b = _d[Pos++];
                if (b == '(') { depth++; bytes.Add(b); }
                else if (b == ')')
                {
                    if (--depth == 0) break;
                    bytes.Add(b);
                }
                else if (b == '\\' && Pos < _d.Length)
                {
                    byte e = _d[Pos++];
                    switch ((char)e)
                    {
                        case 'n': bytes.Add(10); break;
                        case 'r': bytes.Add(13); break;
                        case 't': bytes.Add(9); break;
                        case 'b': bytes.Add(8); break;
                        case 'f': bytes.Add(12); break;
                        case '\r': if (Pos < _d.Length && _d[Pos] == 10) Pos++; break;
                        case '\n': break;
                        default:
                            if (e >= '0' && e <= '7')
                            {
                                int value = e - '0';
                                for (int i = 0; i < 2 && Pos < _d.Length && _d[Pos] >= '0' && _d[Pos] <= '7'; i++) value = value * 8 + (_d[Pos++] - '0');
                                bytes.Add((byte)value);
                            }
                            else bytes.Add(e);
                            break;
                    }
                }
                else bytes.Add(b);
            }
            return bytes.ToArray();
        }

        private byte[] ReadHex()
        {
            Pos++;
            var bytes = new List<byte>();
            int high = -1;
            while (Pos < _d.Length)
            {
                byte b = _d[Pos++];
                if (b == '>') break;
                int value = HexValue(b);
                if (value < 0) continue;
                if (high < 0) high = value;
                else
                {
                    bytes.Add((byte)(high * 16 + value));
                    high = -1;
                }
            }
            if (high >= 0) bytes.Add((byte)(high * 16));
            return bytes.ToArray();
        }

        private List<object> ReadArray()
        {
            var items = new List<object>();
            while (true)
            {
                object item = Read();
                if (item == Eof) break;
                var op = item as PdfOp;
                if (op != null)
                {
                    if (op.Value == "]") break;
                    if (op.Value == ">>") break;
                    continue;
                }
                items.Add(item);
            }
            return items;
        }

        private PdfDict ReadDict()
        {
            var dict = new PdfDict();
            while (true)
            {
                object key = Read();
                if (key == Eof) break;
                var op = key as PdfOp;
                if (op != null && op.Value == ">>") break;
                var name = key as PdfName;
                if (name == null) continue;
                object value = Read();
                if (value == Eof) break;
                op = value as PdfOp;
                if (op != null)
                {
                    if (op.Value == ">>") break;
                    continue;
                }
                dict[name.Value] = value;
            }
            return dict;
        }

        // Inline image data has no length, so the end has to be found by looking for EI.
        public void SkipInlineImage()
        {
            Pos++;
            while (Pos + 2 < _d.Length)
            {
                if (IsSpace(_d[Pos]) && _d[Pos + 1] == 'E' && _d[Pos + 2] == 'I' && (Pos + 3 >= _d.Length || IsSpace(_d[Pos + 3])))
                {
                    Pos += 3;
                    return;
                }
                Pos++;
            }
            Pos = _d.Length;
        }
    }

    internal sealed class PdfPageRef
    {
        public PdfDict Dict, Resources;
        public double X0, Y0, X1, Y1;
        public int Rotate;
        public double Width { get { return Rotate % 180 == 0 ? X1 - X0 : Y1 - Y0; } }
        public double Height { get { return Rotate % 180 == 0 ? Y1 - Y0 : X1 - X0; } }
    }

    internal sealed class PdfFile
    {
        private readonly byte[] _d;
        private readonly Dictionary<int, object> _objects = new Dictionary<int, object>();
        private PdfDict _root;

        public bool Encrypted { get; private set; }

        public PdfFile(byte[] data)
        {
            _d = data;
            Index();
        }

        // Scanning for "n g obj" instead of trusting the xref table also opens files with a damaged
        // table. Later copies replace earlier ones, which is what incremental saves expect.
        private void Index()
        {
            object rootRef = null;
            for (int i = 0; i + 3 < _d.Length; i++)
            {
                if (_d[i] == 'o' && _d[i + 1] == 'b' && _d[i + 2] == 'j')
                {
                    int number;
                    if (!ObjectHeaderBefore(i, out number)) continue;
                    object value = ParseObjectAt(i + 3);
                    _objects[number] = value;

                    var stream = value as PdfStream;
                    PdfDict dict = stream != null ? stream.Dict : value as PdfDict;
                    if (dict == null) continue;
                    string type = NameOf(dict, "Type");
                    if (type == "XRef" && stream != null)
                    {
                        if (dict.ContainsKey("Root")) rootRef = dict["Root"];
                        if (dict.ContainsKey("Encrypt")) Encrypted = true;
                    }
                    else if (type == "ObjStm" && stream != null && !Encrypted)
                    {
                        ReadObjectStream(stream);
                    }
                    else if (type == "Catalog" && rootRef == null)
                    {
                        _root = dict;
                    }
                }
                else if (_d[i] == 't' && Matches(i, "trailer"))
                {
                    var lexer = new PdfLexer(_d, i + 7, false);
                    var trailer = lexer.Read() as PdfDict;
                    if (trailer == null) continue;
                    if (trailer.ContainsKey("Root")) rootRef = trailer["Root"];
                    if (trailer.ContainsKey("Encrypt")) Encrypted = true;
                }
            }
            if (rootRef != null)
            {
                var resolved = Get(rootRef) as PdfDict;
                if (resolved != null) _root = resolved;
            }
        }

        private bool Matches(int position, string text)
        {
            if (position + text.Length > _d.Length) return false;
            for (int i = 0; i < text.Length; i++)
            {
                if (_d[position + i] != text[i]) return false;
            }
            return true;
        }

        private bool ObjectHeaderBefore(int obj, out int number)
        {
            number = 0;
            int i = obj - 1;
            while (i >= 0 && (_d[i] == 32 || _d[i] == 10 || _d[i] == 13 || _d[i] == 9)) i--;
            int genEnd = i;
            while (i >= 0 && _d[i] >= '0' && _d[i] <= '9') i--;
            if (i == genEnd || i < 0) return false;
            int gap = i;
            while (i >= 0 && (_d[i] == 32 || _d[i] == 10 || _d[i] == 13 || _d[i] == 9)) i--;
            if (i == gap) return false;
            int numEnd = i;
            while (i >= 0 && _d[i] >= '0' && _d[i] <= '9') i--;
            if (i == numEnd || numEnd - i > 9) return false;
            for (int k = i + 1; k <= numEnd; k++) number = number * 10 + (_d[k] - '0');
            return true;
        }

        private object ParseObjectAt(int position)
        {
            var lexer = new PdfLexer(_d, position, false);
            object value = lexer.Read();
            if (value == PdfLexer.Eof) return null;
            var dict = value as PdfDict;
            if (dict == null) return value is PdfOp ? null : value;

            lexer.SkipSpace();
            if (!Matches(lexer.Pos, "stream")) return dict;
            int start = lexer.Pos + 6;
            if (start < _d.Length && _d[start] == 13) start++;
            if (start < _d.Length && _d[start] == 10) start++;
            return new PdfStream { Dict = dict, Source = _d, Start = start };
        }

        private void ReadObjectStream(PdfStream stream)
        {
            try
            {
                byte[] data = Decode(stream);
                int count = (int)Number(stream.Dict, "N", 0);
                int first = (int)Number(stream.Dict, "First", 0);
                var header = new PdfLexer(data, 0, true);
                for (int i = 0; i < count; i++)
                {
                    object number = header.Read();
                    object offset = header.Read();
                    if (!(number is double) || !(offset is double)) break;
                    var lexer = new PdfLexer(data, first + (int)(double)offset, false);
                    object value = lexer.Read();
                    _objects[(int)(double)number] = value == PdfLexer.Eof || value is PdfOp ? null : value;
                }
            }
            catch (Exception ex)
            {
                if (!(ex is InvalidDataException || ex is IndexOutOfRangeException || ex is ArgumentException)) throw;
                Log.Write("PDF 개체 묶음을 읽지 못함: " + ex.Message);
            }
        }

        public object Get(object value)
        {
            for (int depth = 0; depth < 32; depth++)
            {
                var reference = value as PdfRef;
                if (reference == null) return value;
                object next;
                if (!_objects.TryGetValue(reference.Number, out next)) return null;
                value = next;
            }
            return null;
        }

        public object Get(PdfDict dict, string key)
        {
            object value;
            return dict != null && dict.TryGetValue(key, out value) ? Get(value) : null;
        }

        public PdfDict DictOf(PdfDict dict, string key)
        {
            object value = Get(dict, key);
            var stream = value as PdfStream;
            return stream != null ? stream.Dict : value as PdfDict;
        }

        public List<object> ArrayOf(PdfDict dict, string key)
        {
            return Get(dict, key) as List<object>;
        }

        public string NameOf(PdfDict dict, string key)
        {
            var name = Get(dict, key) as PdfName;
            return name == null ? null : name.Value;
        }

        public double Number(PdfDict dict, string key, double fallback)
        {
            object value = Get(dict, key);
            return value is double ? (double)value : fallback;
        }

        public List<PdfPageRef> Pages()
        {
            var pages = new List<PdfPageRef>();
            if (_root != null) CollectPages(Get(_root, "Pages") as PdfDict, null, null, null, 0, pages, new HashSet<PdfDict>());
            return pages;
        }

        private void CollectPages(PdfDict node, PdfDict resources, List<object> media, List<object> crop, int rotate, List<PdfPageRef> pages, HashSet<PdfDict> seen)
        {
            if (node == null || !seen.Add(node)) return;
            resources = DictOf(node, "Resources") ?? resources;
            media = ArrayOf(node, "MediaBox") ?? media;
            crop = ArrayOf(node, "CropBox") ?? crop;
            object rotation = Get(node, "Rotate");
            if (rotation is double) rotate = (int)(double)rotation;

            List<object> kids = ArrayOf(node, "Kids");
            if (kids != null)
            {
                foreach (object kid in kids) CollectPages(Get(kid) as PdfDict, resources, media, crop, rotate, pages, seen);
                return;
            }

            double[] box = BoxOf(crop) ?? BoxOf(media) ?? new double[] { 0, 0, 595, 842 };
            double[] outer = BoxOf(media);
            if (outer != null)
            {
                double x0 = Math.Max(box[0], outer[0]), y0 = Math.Max(box[1], outer[1]);
                double x1 = Math.Min(box[2], outer[2]), y1 = Math.Min(box[3], outer[3]);
                if (x1 > x0 && y1 > y0) box = new[] { x0, y0, x1, y1 };
            }
            pages.Add(new PdfPageRef
            {
                Dict = node,
                Resources = resources,
                X0 = box[0],
                Y0 = box[1],
                X1 = box[2],
                Y1 = box[3],
                Rotate = ((rotate % 360) + 360) % 360
            });
        }

        private double[] BoxOf(List<object> values)
        {
            if (values == null || values.Count < 4) return null;
            var n = values.Take(4).Select(v => Get(v)).Select(v => v is double ? (double)v : 0).ToArray();
            return new[] { Math.Min(n[0], n[2]), Math.Min(n[1], n[3]), Math.Max(n[0], n[2]), Math.Max(n[1], n[3]) };
        }

        public byte[] Decode(PdfStream stream)
        {
            byte[] data = RawBytes(stream);
            object filter = Get(stream.Dict, "Filter");
            object parms = Get(stream.Dict, "DecodeParms") ?? Get(stream.Dict, "DP");
            var filters = filter as List<object> ?? (filter == null ? new List<object>() : new List<object> { filter });
            var allParms = parms as List<object> ?? new List<object> { parms };

            for (int i = 0; i < filters.Count; i++)
            {
                var name = Get(filters[i]) as PdfName;
                if (name == null) continue;
                PdfDict options = i < allParms.Count ? Get(allParms[i]) as PdfDict : null;
                switch (name.Value)
                {
                    case "FlateDecode":
                    case "Fl":
                        data = Predict(Inflate(data), options);
                        break;
                    case "LZWDecode":
                    case "LZW":
                        data = Predict(Lzw(data, (int)Number(options, "EarlyChange", 1)), options);
                        break;
                    case "ASCII85Decode":
                    case "A85":
                        data = Ascii85(data);
                        break;
                    case "ASCIIHexDecode":
                    case "AHx":
                        data = AsciiHex(data);
                        break;
                    case "RunLengthDecode":
                    case "RL":
                        data = RunLength(data);
                        break;
                    default:
                        throw new InvalidDataException("지원하지 않는 필터: " + name.Value);
                }
            }
            return data;
        }

        private byte[] RawBytes(PdfStream stream)
        {
            byte[] d = stream.Source;
            object lengthValue = Get(stream.Dict, "Length");
            if (lengthValue is double)
            {
                int length = (int)(double)lengthValue;
                int end = stream.Start + length;
                if (length >= 0 && end <= d.Length)
                {
                    var lexer = new PdfLexer(d, end, false);
                    lexer.SkipSpace();
                    if (lexer.Pos + 9 <= d.Length && Encoding.ASCII.GetString(d, lexer.Pos, 9) == "endstream")
                    {
                        var exact = new byte[length];
                        Array.Copy(d, stream.Start, exact, 0, length);
                        return exact;
                    }
                }
            }

            int stop = stream.Start;
            while (stop + 9 <= d.Length && !(d[stop] == 'e' && Encoding.ASCII.GetString(d, stop, 9) == "endstream")) stop++;
            if (stop + 9 > d.Length) stop = d.Length;
            if (stop > stream.Start && d[stop - 1] == 10) stop--;
            if (stop > stream.Start && d[stop - 1] == 13) stop--;
            var found = new byte[stop - stream.Start];
            Array.Copy(d, stream.Start, found, 0, found.Length);
            return found;
        }

        private static byte[] Inflate(byte[] data)
        {
            using (var output = new MemoryStream())
            {
                if (data.Length <= 2) return new byte[0];
                try
                {
                    using (var input = new MemoryStream(data, 2, data.Length - 2))
                    using (var deflate = new DeflateStream(input, CompressionMode.Decompress))
                    {
                        deflate.CopyTo(output);
                    }
                }
                catch (InvalidDataException)
                {
                    // Keep whatever came out: a truncated stream still has readable text before the damage.
                    if (output.Length == 0) throw;
                }
                return output.ToArray();
            }
        }

        private byte[] Predict(byte[] data, PdfDict options)
        {
            int predictor = (int)Number(options, "Predictor", 1);
            if (predictor < 10) return data;
            int colors = (int)Number(options, "Colors", 1);
            int bits = (int)Number(options, "BitsPerComponent", 8);
            int columns = (int)Number(options, "Columns", 1);
            int pixel = Math.Max(1, colors * bits / 8);
            int row = (colors * bits * columns + 7) / 8;
            var output = new List<byte>(data.Length);
            var previous = new byte[row];
            for (int p = 0; p + 1 + row <= data.Length; p += row + 1)
            {
                int type = data[p];
                var current = new byte[row];
                for (int i = 0; i < row; i++)
                {
                    int left = i >= pixel ? current[i - pixel] : 0;
                    int up = previous[i];
                    int corner = i >= pixel ? previous[i - pixel] : 0;
                    int add = 0;
                    if (type == 1) add = left;
                    else if (type == 2) add = up;
                    else if (type == 3) add = (left + up) / 2;
                    else if (type == 4)
                    {
                        int estimate = left + up - corner;
                        int da = Math.Abs(estimate - left), db = Math.Abs(estimate - up), dc = Math.Abs(estimate - corner);
                        add = da <= db && da <= dc ? left : db <= dc ? up : corner;
                    }
                    current[i] = (byte)(data[p + 1 + i] + add);
                }
                output.AddRange(current);
                previous = current;
            }
            return output.ToArray();
        }

        private static byte[] Lzw(byte[] data, int early)
        {
            var output = new MemoryStream();
            var table = new List<byte[]>();
            Action reset = () =>
            {
                table.Clear();
                for (int i = 0; i < 256; i++) table.Add(new[] { (byte)i });
                table.Add(new byte[0]);
                table.Add(new byte[0]);
            };
            reset();
            int width = 9, bitBuffer = 0, bitCount = 0;
            byte[] previous = null;
            foreach (byte b in data)
            {
                bitBuffer = (bitBuffer << 8) | b;
                bitCount += 8;
                while (bitCount >= width)
                {
                    int code = (bitBuffer >> (bitCount - width)) & ((1 << width) - 1);
                    bitCount -= width;
                    if (code == 256) { reset(); width = 9; previous = null; continue; }
                    if (code == 257) return output.ToArray();
                    byte[] entry;
                    if (code < table.Count) entry = table[code];
                    else if (previous != null) entry = previous.Concat(new[] { previous[0] }).ToArray();
                    else return output.ToArray();
                    output.Write(entry, 0, entry.Length);
                    if (previous != null) table.Add(previous.Concat(new[] { entry[0] }).ToArray());
                    previous = entry;
                    int size = table.Count + early;
                    if (size >= 2048) width = 12;
                    else if (size >= 1024) width = 11;
                    else if (size >= 512) width = 10;
                }
            }
            return output.ToArray();
        }

        private static byte[] Ascii85(byte[] data)
        {
            var output = new MemoryStream();
            var group = new int[5];
            int count = 0;
            foreach (byte b in data)
            {
                if (b == '~') break;
                if (b == 'z' && count == 0) { output.Write(new byte[4], 0, 4); continue; }
                if (b < '!' || b > 'u') continue;
                group[count++] = b - '!';
                if (count == 5)
                {
                    uint value = 0;
                    foreach (int digit in group) value = value * 85 + (uint)digit;
                    output.WriteByte((byte)(value >> 24));
                    output.WriteByte((byte)(value >> 16));
                    output.WriteByte((byte)(value >> 8));
                    output.WriteByte((byte)value);
                    count = 0;
                }
            }
            if (count > 1)
            {
                for (int i = count; i < 5; i++) group[i] = 84;
                uint value = 0;
                foreach (int digit in group) value = value * 85 + (uint)digit;
                for (int i = 0; i < count - 1; i++) output.WriteByte((byte)(value >> (24 - i * 8)));
            }
            return output.ToArray();
        }

        private static byte[] AsciiHex(byte[] data)
        {
            var wrapped = new byte[data.Length + 2];
            wrapped[0] = (byte)'<';
            Array.Copy(data, 0, wrapped, 1, data.Length);
            wrapped[wrapped.Length - 1] = (byte)'>';
            return new PdfLexer(wrapped, 0, true).Read() as byte[] ?? new byte[0];
        }

        private static byte[] RunLength(byte[] data)
        {
            var output = new MemoryStream();
            int i = 0;
            while (i < data.Length)
            {
                int length = data[i++];
                if (length == 128) break;
                if (length < 128)
                {
                    int count = Math.Min(length + 1, data.Length - i);
                    output.Write(data, i, count);
                    i += count;
                }
                else if (i < data.Length)
                {
                    for (int k = 0; k < 257 - length; k++) output.WriteByte(data[i]);
                    i++;
                }
            }
            return output.ToArray();
        }
    }

    internal struct PdfGlyph
    {
        public string Text;
        public double Width;
        public bool IsSpace;
    }

    internal sealed class PdfFont
    {
        private const string Unknown = "�";

        private static readonly Dictionary<string, string> GlyphNames = new Dictionary<string, string>
        {
            {"space"," "},{"exclam","!"},{"quotedbl","\""},{"numbersign","#"},{"dollar","$"},{"percent","%"},{"ampersand","&"},
            {"quotesingle","'"},{"parenleft","("},{"parenright",")"},{"asterisk","*"},{"plus","+"},{"comma",","},{"hyphen","-"},
            {"period","."},{"slash","/"},{"zero","0"},{"one","1"},{"two","2"},{"three","3"},{"four","4"},{"five","5"},{"six","6"},
            {"seven","7"},{"eight","8"},{"nine","9"},{"colon",":"},{"semicolon",";"},{"less","<"},{"equal","="},{"greater",">"},
            {"question","?"},{"at","@"},{"bracketleft","["},{"backslash","\\"},{"bracketright","]"},{"asciicircum","^"},
            {"underscore","_"},{"grave","`"},{"braceleft","{"},{"bar","|"},{"braceright","}"},{"asciitilde","~"},{"bullet","•"},
            {"endash","–"},{"emdash","—"},{"quoteleft","‘"},{"quoteright","’"},{"quotedblleft","“"},
            {"quotedblright","”"},{"ellipsis","…"},{"fi","fi"},{"fl","fl"},{"periodcentered","·"},{"nbspace"," "},
            {"copyright","©"},{"registered","®"},{"trademark","™"},{"degree","°"},{"minus","−"},
            {"multiply","×"},{"divide","÷"},{"Euro","€"}
        };

        private readonly bool _composite;
        private readonly Dictionary<int, string> _toUnicode;
        private readonly List<int[]> _codeSpaces = new List<int[]>();
        private readonly string[] _simple = new string[256];
        private readonly Dictionary<int, double> _widths = new Dictionary<int, double>();
        private readonly double _defaultWidth;
        private readonly Encoding _multibyte;
        private readonly bool _utf16;
        private readonly bool _symbol;

        public readonly double Scale;
        public readonly bool Bold, Italic;

        public PdfFont(PdfFile file, PdfDict font)
        {
            string subtype = file.NameOf(font, "Subtype");
            string baseFont = file.NameOf(font, "BaseFont") ?? "";
            _composite = subtype == "Type0";
            Scale = 0.001;
            if (subtype == "Type3")
            {
                List<object> matrix = file.ArrayOf(font, "FontMatrix");
                if (matrix != null && matrix.Count > 0 && file.Get(matrix[0]) is double) Scale = (double)file.Get(matrix[0]);
            }

            PdfDict descendant = null;
            if (_composite)
            {
                List<object> list = file.ArrayOf(font, "DescendantFonts");
                if (list != null && list.Count > 0) descendant = file.Get(list[0]) as PdfDict;
            }
            PdfDict descriptor = file.DictOf(descendant ?? font, "FontDescriptor");

            int flags = (int)file.Number(descriptor, "Flags", 0);
            double weight = file.Number(descriptor, "FontWeight", 0);
            Bold = Regex.IsMatch(baseFont, @"bold|black|heavy|demi|[-,]B(?![a-z])|[a-z]Bd", RegexOptions.None) || Regex.IsMatch(baseFont, "bold|black|heavy", RegexOptions.IgnoreCase)
                || weight >= 600 || (flags & (1 << 18)) != 0;
            Italic = Regex.IsMatch(baseFont, "italic|oblique", RegexOptions.IgnoreCase) || file.Number(descriptor, "ItalicAngle", 0) != 0;
            _symbol = Regex.IsMatch(baseFont, "symbol|wingdings|dingbats|webdings", RegexOptions.IgnoreCase);

            var toUnicode = file.Get(font, "ToUnicode") as PdfStream;
            if (toUnicode != null)
            {
                try
                {
                    _toUnicode = new Dictionary<int, string>();
                    ReadCMap(file.Decode(toUnicode));
                }
                catch (InvalidDataException ex)
                {
                    _toUnicode = null;
                    Log.Write("PDF 글꼴 표를 읽지 못함: " + ex.Message);
                }
            }

            if (_composite)
            {
                string encoding = file.NameOf(font, "Encoding") ?? "";
                if (Regex.IsMatch(encoding, "UCS2|UTF16")) _utf16 = true;
                else if (Regex.IsMatch(encoding, "^(KSC|UHC)")) _multibyte = Encoding.GetEncoding(949);
                else if (Regex.IsMatch(encoding, "^(GB)")) _multibyte = Encoding.GetEncoding(936);
                else if (Regex.IsMatch(encoding, "^(B5|ETen|ETHK|HK|CNS)")) _multibyte = Encoding.GetEncoding(950);
                else if (Regex.IsMatch(encoding, "RKSJ|^EUC-|^H$|^V$")) _multibyte = Encoding.GetEncoding(932);
                _defaultWidth = file.Number(descendant, "DW", 1000);
                ReadCompositeWidths(file, descendant);
            }
            else
            {
                ReadSimpleEncoding(file, font, baseFont);
                List<object> widths = file.ArrayOf(font, "Widths");
                int first = (int)file.Number(font, "FirstChar", 0);
                if (widths != null)
                {
                    for (int i = 0; i < widths.Count; i++)
                    {
                        object w = file.Get(widths[i]);
                        if (w is double) _widths[first + i] = (double)w;
                    }
                }
                _defaultWidth = file.Number(descriptor, "MissingWidth", widths == null ? (baseFont.Contains("Courier") ? 600 : 500) : 0);
            }
        }

        private void ReadCompositeWidths(PdfFile file, PdfDict descendant)
        {
            List<object> w = file.ArrayOf(descendant, "W");
            if (w == null) return;
            int i = 0;
            while (i < w.Count)
            {
                object first = file.Get(w[i++]);
                if (!(first is double) || i >= w.Count) break;
                object second = file.Get(w[i++]);
                var list = second as List<object>;
                if (list != null)
                {
                    for (int k = 0; k < list.Count; k++)
                    {
                        object value = file.Get(list[k]);
                        if (value is double) _widths[(int)(double)first + k] = (double)value;
                    }
                }
                else if (second is double && i < w.Count)
                {
                    object value = file.Get(w[i++]);
                    if (value is double)
                    {
                        int last = Math.Min((int)(double)second, (int)(double)first + 65535);
                        for (int code = (int)(double)first; code <= last; code++) _widths[code] = (double)value;
                    }
                }
            }
        }

        private void ReadSimpleEncoding(PdfFile file, PdfDict font, string baseFont)
        {
            object encoding = file.Get(font, "Encoding");
            var dict = encoding as PdfDict;
            string baseName = encoding is PdfName ? ((PdfName)encoding).Value : file.NameOf(dict, "BaseEncoding");
            Encoding table = baseName == "MacRomanEncoding" ? Encoding.GetEncoding(10000) : Encoding.GetEncoding(1252);
            bool pictures = Regex.IsMatch(baseFont, "wingdings|dingbats|webdings", RegexOptions.IgnoreCase);
            for (int i = 0; i < 256; i++)
            {
                if (pictures) _simple[i] = "•";
                else if (i < 32) _simple[i] = Unknown;
                else _simple[i] = table.GetString(new[] { (byte)i });
            }
            if (_symbol && !pictures) _simple[0xB7] = "•";

            List<object> differences = file.ArrayOf(dict, "Differences");
            if (differences == null) return;
            int code = 0;
            foreach (object raw in differences)
            {
                object item = file.Get(raw);
                if (item is double) code = (int)(double)item;
                else if (item is PdfName)
                {
                    if (code >= 0 && code < 256) _simple[code] = GlyphText(((PdfName)item).Value);
                    code++;
                }
            }
        }

        private static string GlyphText(string name)
        {
            string text;
            if (GlyphNames.TryGetValue(name, out text)) return text;
            if (name.Length == 1) return name;
            int dot = name.IndexOf('.');
            if (dot > 0) return GlyphText(name.Substring(0, dot));
            Match hex = Regex.Match(name, "^(?:uni|u)([0-9A-Fa-f]{4,6})$");
            if (hex.Success)
            {
                int code = int.Parse(hex.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                if (code <= 0x10FFFF && (code < 0xD800 || code > 0xDFFF)) return char.ConvertFromUtf32(code);
            }
            return Unknown;
        }

        private void ReadCMap(byte[] data)
        {
            var lexer = new PdfLexer(data, 0, true);
            var args = new List<object>();
            string section = null;
            while (true)
            {
                object token = lexer.Read();
                if (token == PdfLexer.Eof) break;
                var op = token as PdfOp;
                if (op == null)
                {
                    args.Add(token);
                    continue;
                }
                switch (op.Value)
                {
                    case "begincodespacerange":
                    case "beginbfchar":
                    case "beginbfrange":
                        section = op.Value;
                        args.Clear();
                        break;
                    case "endcodespacerange":
                        for (int i = 0; i + 1 < args.Count; i += 2)
                        {
                            var low = args[i] as byte[];
                            var high = args[i + 1] as byte[];
                            if (low != null && high != null) _codeSpaces.Add(new[] { low.Length, CodeOf(low), CodeOf(high) });
                        }
                        section = null;
                        args.Clear();
                        break;
                    case "endbfchar":
                        for (int i = 0; i + 1 < args.Count; i += 2)
                        {
                            var code = args[i] as byte[];
                            var text = args[i + 1] as byte[];
                            if (code != null && text != null) _toUnicode[CodeOf(code)] = Utf16(text);
                        }
                        section = null;
                        args.Clear();
                        break;
                    case "endbfrange":
                        for (int i = 0; i + 2 < args.Count; i += 3) ReadRange(args[i] as byte[], args[i + 1] as byte[], args[i + 2]);
                        section = null;
                        args.Clear();
                        break;
                    default:
                        if (section == null) args.Clear();
                        break;
                }
            }
        }

        private void ReadRange(byte[] low, byte[] high, object target)
        {
            if (low == null || high == null) return;
            int first = CodeOf(low), last = Math.Min(CodeOf(high), CodeOf(low) + 65535);
            var list = target as List<object>;
            if (list != null)
            {
                for (int k = 0; k < list.Count && first + k <= last; k++)
                {
                    var text = list[k] as byte[];
                    if (text != null) _toUnicode[first + k] = Utf16(text);
                }
                return;
            }
            var start = target as byte[];
            if (start == null || start.Length == 0) return;
            for (int code = first; code <= last; code++)
            {
                var text = (byte[])start.Clone();
                int add = code - first;
                for (int p = text.Length - 1; p >= 0 && add > 0; p--)
                {
                    add += text[p];
                    text[p] = (byte)add;
                    add >>= 8;
                }
                _toUnicode[code] = Utf16(text);
            }
        }

        private static int CodeOf(byte[] bytes)
        {
            int code = 0;
            foreach (byte b in bytes) code = (code << 8) | b;
            return code;
        }

        private static string Utf16(byte[] bytes)
        {
            return Encoding.BigEndianUnicode.GetString(bytes);
        }

        public List<PdfGlyph> Decode(byte[] text)
        {
            var glyphs = new List<PdfGlyph>(text.Length);
            int i = 0;
            while (i < text.Length)
            {
                if (!_composite)
                {
                    int code = text[i++];
                    string mapped = null;
                    if (_toUnicode != null) _toUnicode.TryGetValue(code, out mapped);
                    if (mapped == null) mapped = _simple[code];
                    glyphs.Add(Glyph(mapped, code, code == 32));
                }
                else if (_utf16)
                {
                    int length = Math.Min(2, text.Length - i);
                    glyphs.Add(Glyph(Encoding.BigEndianUnicode.GetString(text, i, length), -1, false));
                    i += length;
                }
                else if (_multibyte != null)
                {
                    int length = text[i] >= 0x81 && i + 1 < text.Length ? 2 : 1;
                    glyphs.Add(Glyph(_multibyte.GetString(text, i, length), -1, length == 1 && text[i] == 32));
                    i += length;
                }
                else
                {
                    int length = CodeLength(text, i);
                    int code = 0;
                    for (int k = 0; k < length; k++) code = (code << 8) | text[i + k];
                    i += length;
                    string mapped = null;
                    if (_toUnicode != null) _toUnicode.TryGetValue(code, out mapped);
                    glyphs.Add(Glyph(mapped ?? Unknown, code, length == 1 && code == 32));
                }
            }
            return glyphs;
        }

        private int CodeLength(byte[] text, int position)
        {
            foreach (int[] space in _codeSpaces)
            {
                int length = space[0];
                if (position + length > text.Length) continue;
                int code = 0;
                for (int k = 0; k < length; k++) code = (code << 8) | text[position + k];
                if (code >= space[1] && code <= space[2]) return length;
            }
            return Math.Min(2, text.Length - position);
        }

        private PdfGlyph Glyph(string text, int code, bool isSpace)
        {
            double width;
            if (code < 0 || !_widths.TryGetValue(code, out width)) width = _defaultWidth;
            return new PdfGlyph { Text = Clean(text), Width = width, IsSpace = isSpace };
        }

        private string Clean(string text)
        {
            if (text.Length == 1)
            {
                char c = text[0];
                if (c >= 0xE000 && c <= 0xF8FF) return _symbol ? "•" : Unknown;
                if (c < 32) return c == 9 ? " " : "";
                if (c >= 0xFB00 && c <= 0xFB06) return text.Normalize(NormalizationForm.FormKC);
            }
            return text;
        }
    }

    internal sealed class PdfTextItem
    {
        public double X, Y, EndX, Size;
        public string Text;
        public bool Bold, Italic;
    }

    internal sealed class PdfLineMark
    {
        public bool Horizontal;
        public double Position, From, To;
    }

    internal sealed class PdfImageMark
    {
        public int Id;
        public double X0, Y0, X1, Y1;
    }

    internal sealed class PdfPageContent
    {
        public double Width, Height;
        public List<PdfTextItem> Text = new List<PdfTextItem>();
        public List<PdfTextItem> Sideways = new List<PdfTextItem>();
        public List<PdfLineMark> Lines = new List<PdfLineMark>();
        public List<PdfImageMark> Images = new List<PdfImageMark>();
        public List<double[]> Dots = new List<double[]>();
        public bool Ocr;
    }

    // Plays a page's drawing commands and records where text, lines and images land. Coordinates
    // in the result have the origin at the top left of the page as displayed, in points.
    internal sealed class PdfPageReader
    {
        private struct Matrix
        {
            public double A, B, C, D, E, F;

            public Matrix(double a, double b, double c, double d, double e, double f)
            {
                A = a; B = b; C = c; D = d; E = e; F = f;
            }

            public Matrix Then(Matrix m)
            {
                return new Matrix(
                    A * m.A + B * m.C, A * m.B + B * m.D,
                    C * m.A + D * m.C, C * m.B + D * m.D,
                    E * m.A + F * m.C + m.E, E * m.B + F * m.D + m.F);
            }

            public double X(double x, double y) { return A * x + C * y + E; }
            public double Y(double x, double y) { return B * x + D * y + F; }
        }

        private sealed class State
        {
            public Matrix Ctm;
            public PdfFont Font;
            public double Size, CharSpace, WordSpace, Scale = 1, Leading, Rise;
            public int RenderMode;

            public State Copy() { return (State)MemberwiseClone(); }
        }

        private sealed class SubPath
        {
            public double X0 = double.MaxValue, Y0 = double.MaxValue, X1 = double.MinValue, Y1 = double.MinValue;
            public readonly List<double[]> Edges = new List<double[]>();

            public void Include(double x, double y)
            {
                X0 = Math.Min(X0, x); Y0 = Math.Min(Y0, y);
                X1 = Math.Max(X1, x); Y1 = Math.Max(Y1, y);
            }
        }

        private const int MaxFormDepth = 8;

        private readonly PdfFile _file;
        private readonly Dictionary<PdfDict, PdfFont> _fonts = new Dictionary<PdfDict, PdfFont>();
        private PdfPageContent _page;
        private HashSet<string> _seen;
        private int _nextImageId;
        private readonly Dictionary<object, int> _imageIds = new Dictionary<object, int>();

        public PdfPageReader(PdfFile file)
        {
            _file = file;
        }

        public PdfPageContent Read(PdfPageRef page)
        {
            _page = new PdfPageContent { Width = page.Width, Height = page.Height };
            _seen = new HashSet<string>();
            var state = new State { Ctm = BaseMatrix(page) };
            var content = new MemoryStream();
            object contents = _file.Get(page.Dict, "Contents");
            var list = contents as List<object> ?? new List<object> { contents };
            foreach (object part in list)
            {
                var stream = _file.Get(part) as PdfStream;
                if (stream == null) continue;
                byte[] data = _file.Decode(stream);
                content.Write(data, 0, data.Length);
                content.WriteByte(10);
            }
            Execute(content.ToArray(), page.Resources, state, 0);
            return _page;
        }

        private static Matrix BaseMatrix(PdfPageRef p)
        {
            switch (p.Rotate)
            {
                case 90: return new Matrix(0, 1, 1, 0, -p.Y0, -p.X0);
                case 180: return new Matrix(-1, 0, 0, 1, p.X1, -p.Y0);
                case 270: return new Matrix(0, -1, -1, 0, p.Y1, p.X1);
                default: return new Matrix(1, 0, 0, -1, -p.X0, p.Y1);
            }
        }

        private static double Num(List<object> args, int index)
        {
            return index >= 0 && index < args.Count && args[index] is double ? (double)args[index] : 0;
        }

        private void Execute(byte[] data, PdfDict resources, State state, int depth)
        {
            var lexer = new PdfLexer(data, 0, true);
            var args = new List<object>();
            var stack = new Stack<State>();
            var marked = new Stack<bool>();
            int artifact = 0;

            Matrix tm = new Matrix(1, 0, 0, 1, 0, 0), tlm = tm;
            var paths = new List<SubPath>();
            SubPath path = null;
            double curX = 0, curY = 0;

            while (true)
            {
                object token = lexer.Read();
                if (token == PdfLexer.Eof) break;
                var op = token as PdfOp;
                if (op == null)
                {
                    args.Add(token);
                    if (args.Count > 512) args.RemoveAt(0);
                    continue;
                }

                int n = args.Count;
                switch (op.Value)
                {
                    case "q": stack.Push(state.Copy()); break;
                    case "Q": if (stack.Count > 0) state = stack.Pop(); break;
                    case "cm":
                        state.Ctm = new Matrix(Num(args, n - 6), Num(args, n - 5), Num(args, n - 4), Num(args, n - 3), Num(args, n - 2), Num(args, n - 1)).Then(state.Ctm);
                        break;

                    case "BT": tm = tlm = new Matrix(1, 0, 0, 1, 0, 0); break;
                    case "Tf":
                        state.Size = Num(args, n - 1);
                        if (n >= 2 && args[n - 2] is PdfName) state.Font = FontOf(resources, ((PdfName)args[n - 2]).Value);
                        break;
                    case "Tc": state.CharSpace = Num(args, n - 1); break;
                    case "Tw": state.WordSpace = Num(args, n - 1); break;
                    case "Tz": state.Scale = Num(args, n - 1) / 100; break;
                    case "TL": state.Leading = Num(args, n - 1); break;
                    case "Ts": state.Rise = Num(args, n - 1); break;
                    case "Tr": state.RenderMode = (int)Num(args, n - 1); break;
                    case "Td":
                        tm = tlm = new Matrix(1, 0, 0, 1, Num(args, n - 2), Num(args, n - 1)).Then(tlm);
                        break;
                    case "TD":
                        state.Leading = -Num(args, n - 1);
                        tm = tlm = new Matrix(1, 0, 0, 1, Num(args, n - 2), Num(args, n - 1)).Then(tlm);
                        break;
                    case "Tm":
                        tm = tlm = new Matrix(Num(args, n - 6), Num(args, n - 5), Num(args, n - 4), Num(args, n - 3), Num(args, n - 2), Num(args, n - 1));
                        break;
                    case "T*":
                        tm = tlm = new Matrix(1, 0, 0, 1, 0, -state.Leading).Then(tlm);
                        break;
                    case "Tj":
                        if (n >= 1 && args[n - 1] is byte[]) tm = Show((byte[])args[n - 1], tm, state, artifact > 0);
                        break;
                    case "'":
                        tm = tlm = new Matrix(1, 0, 0, 1, 0, -state.Leading).Then(tlm);
                        if (n >= 1 && args[n - 1] is byte[]) tm = Show((byte[])args[n - 1], tm, state, artifact > 0);
                        break;
                    case "\"":
                        state.WordSpace = Num(args, n - 3);
                        state.CharSpace = Num(args, n - 2);
                        tm = tlm = new Matrix(1, 0, 0, 1, 0, -state.Leading).Then(tlm);
                        if (n >= 1 && args[n - 1] is byte[]) tm = Show((byte[])args[n - 1], tm, state, artifact > 0);
                        break;
                    case "TJ":
                        var array = n >= 1 ? args[n - 1] as List<object> : null;
                        if (array == null) break;
                        foreach (object item in array)
                        {
                            if (item is byte[]) tm = Show((byte[])item, tm, state, artifact > 0);
                            else if (item is double)
                                tm = new Matrix(1, 0, 0, 1, -(double)item / 1000 * state.Size * state.Scale, 0).Then(tm);
                        }
                        break;

                    case "BMC":
                        marked.Push(false);
                        break;
                    case "BDC":
                        bool isArtifact = n >= 2 && args[n - 2] is PdfName && ((PdfName)args[n - 2]).Value == "Artifact";
                        marked.Push(isArtifact);
                        if (isArtifact) artifact++;
                        break;
                    case "EMC":
                        if (marked.Count > 0 && marked.Pop()) artifact--;
                        break;

                    case "m":
                        path = new SubPath();
                        paths.Add(path);
                        curX = state.Ctm.X(Num(args, n - 2), Num(args, n - 1));
                        curY = state.Ctm.Y(Num(args, n - 2), Num(args, n - 1));
                        path.Include(curX, curY);
                        break;
                    case "l":
                        if (path == null) break;
                        double lx = state.Ctm.X(Num(args, n - 2), Num(args, n - 1));
                        double ly = state.Ctm.Y(Num(args, n - 2), Num(args, n - 1));
                        path.Edges.Add(new[] { curX, curY, lx, ly });
                        path.Include(lx, ly);
                        curX = lx;
                        curY = ly;
                        break;
                    case "c":
                    case "v":
                    case "y":
                        if (path == null) break;
                        curX = state.Ctm.X(Num(args, n - 2), Num(args, n - 1));
                        curY = state.Ctm.Y(Num(args, n - 2), Num(args, n - 1));
                        path.Include(curX, curY);
                        break;
                    case "re":
                        double rx = Num(args, n - 4), ry = Num(args, n - 3), rw = Num(args, n - 2), rh = Num(args, n - 1);
                        var box = new SubPath();
                        var corners = new[] { rx, ry, rx + rw, ry, rx + rw, ry + rh, rx, ry + rh };
                        for (int i = 0; i < 4; i++)
                        {
                            double x0 = state.Ctm.X(corners[i * 2], corners[i * 2 + 1]), y0 = state.Ctm.Y(corners[i * 2], corners[i * 2 + 1]);
                            int next = (i + 1) % 4;
                            double x1 = state.Ctm.X(corners[next * 2], corners[next * 2 + 1]), y1 = state.Ctm.Y(corners[next * 2], corners[next * 2 + 1]);
                            box.Edges.Add(new[] { x0, y0, x1, y1 });
                            box.Include(x0, y0);
                        }
                        paths.Add(box);
                        path = null;
                        break;
                    case "S":
                    case "s":
                    case "B":
                    case "B*":
                    case "b":
                    case "b*":
                        CommitStroke(paths);
                        paths.Clear();
                        path = null;
                        break;
                    case "f":
                    case "F":
                    case "f*":
                        CommitFill(paths);
                        paths.Clear();
                        path = null;
                        break;
                    case "n":
                        paths.Clear();
                        path = null;
                        break;

                    case "Do":
                        if (n >= 1 && args[n - 1] is PdfName) DrawObject(((PdfName)args[n - 1]).Value, resources, state, depth);
                        break;
                    case "BI":
                        while (true)
                        {
                            object part = lexer.Read();
                            if (part == PdfLexer.Eof) break;
                            var partOp = part as PdfOp;
                            if (partOp != null && partOp.Value == "ID") break;
                        }
                        lexer.SkipInlineImage();
                        MarkImage(state, -1);
                        break;
                }
                args.Clear();
            }
        }

        private PdfFont FontOf(PdfDict resources, string name)
        {
            PdfDict fonts = _file.DictOf(resources, "Font");
            PdfDict dict = _file.DictOf(fonts, name);
            if (dict == null) return null;
            PdfFont font;
            if (!_fonts.TryGetValue(dict, out font))
            {
                font = new PdfFont(_file, dict);
                _fonts[dict] = font;
            }
            return font;
        }

        private Matrix Show(byte[] bytes, Matrix tm, State s, bool hidden)
        {
            if (s.Font == null) return tm;
            var text = new StringBuilder();
            double advance = 0;
            foreach (PdfGlyph glyph in s.Font.Decode(bytes))
            {
                text.Append(glyph.Text);
                advance += (glyph.Width * s.Font.Scale * s.Size + s.CharSpace + (glyph.IsSpace ? s.WordSpace : 0)) * s.Scale;
            }

            Matrix start = new Matrix(s.Size * s.Scale, 0, 0, s.Size, 0, s.Rise).Then(tm).Then(s.Ctm);
            Matrix next = new Matrix(1, 0, 0, 1, advance, 0).Then(tm);
            Matrix end = new Matrix(1, 0, 0, 1, 0, s.Rise).Then(next).Then(s.Ctm);

            string value = text.ToString();
            if (hidden || value.Trim().Length == 0) return next;

            double x = start.E, y = start.F;
            double size = Math.Sqrt(start.C * start.C + start.D * start.D);
            if (size < 2 || x < -5 || x > _page.Width + 5 || y < -5 || y > _page.Height + 5) return next;

            // Fake bold is drawn by printing the same text twice almost on top of itself.
            string key = Math.Round(x / 1.5) + "|" + Math.Round(y / 1.5) + "|" + value;
            bool fakeBold = !_seen.Add(key);
            if (fakeBold)
            {
                foreach (PdfTextItem existing in _page.Text)
                {
                    if (existing.Text == value && Math.Abs(existing.X - x) < 2 && Math.Abs(existing.Y - y) < 2) existing.Bold = true;
                }
                return next;
            }

            var item = new PdfTextItem
            {
                X = x,
                Y = y,
                EndX = end.E,
                Size = size,
                Text = value.Normalize(NormalizationForm.FormC),
                Bold = s.Font.Bold || s.RenderMode == 2,
                // Fonts without an italic style get slanted by the matrix instead.
                Italic = s.Font.Italic || Math.Abs(start.C) > Math.Abs(start.D) * 0.1
            };
            bool horizontal = Math.Abs(end.E - x) >= Math.Abs(end.F - y) && Math.Abs(start.A) >= Math.Abs(start.B) && item.EndX >= item.X;
            if (horizontal) _page.Text.Add(item);
            else _page.Sideways.Add(item);
            return next;
        }

        private void DrawObject(string name, PdfDict resources, State state, int depth)
        {
            PdfDict objects = _file.DictOf(resources, "XObject");
            var stream = _file.Get(objects, name) as PdfStream;
            if (stream == null) return;
            string subtype = _file.NameOf(stream.Dict, "Subtype");
            if (subtype == "Image")
            {
                int id;
                if (!_imageIds.TryGetValue(stream, out id))
                {
                    id = ++_nextImageId;
                    _imageIds[stream] = id;
                }
                MarkImage(state, id);
            }
            else if (subtype == "Form" && depth < MaxFormDepth)
            {
                State inner = state.Copy();
                List<object> matrix = _file.ArrayOf(stream.Dict, "Matrix");
                if (matrix != null && matrix.Count >= 6)
                {
                    var m = matrix.Select(v => _file.Get(v)).Select(v => v is double ? (double)v : 0).ToArray();
                    inner.Ctm = new Matrix(m[0], m[1], m[2], m[3], m[4], m[5]).Then(inner.Ctm);
                }
                Execute(_file.Decode(stream), _file.DictOf(stream.Dict, "Resources") ?? resources, inner, depth + 1);
            }
        }

        private void MarkImage(State s, int id)
        {
            double[] xs = { s.Ctm.X(0, 0), s.Ctm.X(1, 0), s.Ctm.X(0, 1), s.Ctm.X(1, 1) };
            double[] ys = { s.Ctm.Y(0, 0), s.Ctm.Y(1, 0), s.Ctm.Y(0, 1), s.Ctm.Y(1, 1) };
            _page.Images.Add(new PdfImageMark { Id = id, X0 = xs.Min(), Y0 = ys.Min(), X1 = xs.Max(), Y1 = ys.Max() });
        }

        private void CommitStroke(List<SubPath> paths)
        {
            foreach (SubPath path in paths)
            {
                double width = path.X1 - path.X0, height = path.Y1 - path.Y0;
                if (path.Edges.Count == 0 && width >= 2 && width <= 9 && Math.Abs(width - height) < 1.5)
                    _page.Dots.Add(new[] { path.X0, path.Y0, path.X1, path.Y1 });
                foreach (double[] e in path.Edges) AddLine(e[0], e[1], e[2], e[3]);
            }
        }

        // Many programs draw table borders as very thin filled rectangles.
        private void CommitFill(List<SubPath> paths)
        {
            const double Thin = 2.5;
            foreach (SubPath path in paths)
            {
                double width = path.X1 - path.X0, height = path.Y1 - path.Y0;
                if (width >= 2 && width <= 9 && Math.Abs(width - height) < 1.5)
                    _page.Dots.Add(new[] { path.X0, path.Y0, path.X1, path.Y1 });
                else if (height <= Thin && width > Thin) AddLine(path.X0, (path.Y0 + path.Y1) / 2, path.X1, (path.Y0 + path.Y1) / 2);
                else if (width <= Thin && height > Thin) AddLine((path.X0 + path.X1) / 2, path.Y0, (path.X0 + path.X1) / 2, path.Y1);
            }
        }

        private void AddLine(double x0, double y0, double x1, double y1)
        {
            const double Straight = 1, Shortest = 3;
            if (Math.Abs(y1 - y0) <= Straight && Math.Abs(x1 - x0) >= Shortest)
                _page.Lines.Add(new PdfLineMark { Horizontal = true, Position = (y0 + y1) / 2, From = Math.Min(x0, x1), To = Math.Max(x0, x1) });
            else if (Math.Abs(x1 - x0) <= Straight && Math.Abs(y1 - y0) >= Shortest)
                _page.Lines.Add(new PdfLineMark { Horizontal = false, Position = (x0 + x1) / 2, From = Math.Min(y0, y1), To = Math.Max(y0, y1) });
        }
    }
}
