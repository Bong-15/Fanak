using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace Fanak
{
    // Hangul files keep their format under several extensions, so the first bytes decide how to read one.
    internal static class HangulFile
    {
        public static readonly string[] Exts = { "hwp", "hwt", "hwpx", "hwtx", "owpml", "hml", "hwpml" };

        private static readonly byte[] CompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        public static void Read(string source, MdDoc doc, MdImages images)
        {
            try
            {
                byte[] data = ReadShared(source);
                if (data.Length >= 8 && CompoundSignature.SequenceEqual(data.Take(8)))
                {
                    new HwpReader(new CompoundFile(data), doc, images).Run();
                }
                else if (data.Length >= 17 && Encoding.ASCII.GetString(data, 0, 17) == "HWP Document File")
                {
                    new Hwp3Reader(data, doc, images).Run();
                }
                else if (data.Length >= 2 && data[0] == 'P' && data[1] == 'K')
                {
                    using (var package = new DocPackage(source)) new HwpxReader(package, doc, images).Run();
                }
                else
                {
                    new HmlReader(data, doc, images).Run();
                }
            }
            catch (InvalidDataException ex) { throw Unreadable(ex); }
            catch (XmlException ex) { throw Unreadable(ex); }
            catch (EndOfStreamException ex) { throw Unreadable(ex); }
            catch (IndexOutOfRangeException ex) { throw Unreadable(ex); }
            catch (ArgumentException ex) { throw Unreadable(ex); }
            catch (OverflowException ex) { throw Unreadable(ex); }
            catch (FormatException ex) { throw Unreadable(ex); }
        }

        // Hangul keeps the files it has open locked for writing, and that must not stop them being read.
        public static byte[] ReadShared(string path)
        {
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var output = new MemoryStream())
            {
                input.CopyTo(output);
                return output.ToArray();
            }
        }

        // Hangul itself opens any bytes as plain text, which would make a PDF full of garbage.
        public static bool LooksValid(string path)
        {
            var head = new byte[20];
            int length;
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) length = input.Read(head, 0, head.Length);
            if (length >= 8 && CompoundSignature.SequenceEqual(head.Take(8))) return true;
            if (length >= 17 && Encoding.ASCII.GetString(head, 0, 17) == "HWP Document File") return true;
            if (length >= 2 && head[0] == 'P' && head[1] == 'K') return true;
            int text = length >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF ? 3 : 0;
            bool wide = length >= 2 && (head[0] == 0xFF && head[1] == 0xFE || head[0] == 0xFE && head[1] == 0xFF);
            return wide || (length > text && head[text] == '<');
        }

        private static Exception Unreadable(Exception inner)
        {
            return new InvalidOperationException("한글 문서를 읽을 수 없습니다. 손상됐거나 암호가 걸린 문서일 수 있습니다.", inner);
        }

        // Returns 0 when the style is not an outline level.
        public static int OutlineLevel(string name, string englishName)
        {
            Match outline = Regex.Match((name ?? "") + "|" + (englishName ?? ""), @"(개요|Outline)\s*([1-7])", RegexOptions.IgnoreCase);
            return outline.Success ? int.Parse(outline.Groups[2].Value, CultureInfo.InvariantCulture) : 0;
        }

        public static byte[] Inflate(byte[] data)
        {
            using (var input = new MemoryStream(data))
            using (var inflater = new DeflateStream(input, CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                inflater.CopyTo(output);
                return output.ToArray();
            }
        }

        // Saved flags can't always be trusted for pictures, and a wrong guess only fails here.
        public static byte[] InflateOrKeep(byte[] data)
        {
            try { return Inflate(data); }
            catch (InvalidDataException) { return data; }
        }
    }

    // OLE compound file: a small file system inside the .hwp. Nothing in .NET reads it without Office.
    internal sealed class CompoundFile
    {
        private const uint FirstSpecialSector = 0xFFFFFFFA;

        private sealed class Entry
        {
            public string Name;
            public int Type;
            public uint Left, Right, Child, Start;
            public long Size;
        }

        private readonly byte[] _data;
        private readonly int _sectorSize;
        private readonly int _miniSectorSize;
        private readonly uint _miniCutoff;
        private readonly uint[] _fat;
        private readonly uint[] _miniFat;
        private readonly byte[] _miniStream;
        private readonly Dictionary<string, Entry> _streams = new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);

        public CompoundFile(byte[] data)
        {
            _data = data;
            _sectorSize = 1 << BitConverter.ToUInt16(data, 30);
            _miniSectorSize = 1 << BitConverter.ToUInt16(data, 32);
            _miniCutoff = BitConverter.ToUInt32(data, 56);
            _fat = ReadFat();

            List<Entry> entries = ReadDirectory(BitConverter.ToUInt32(data, 48));
            if (entries.Count == 0) throw new InvalidDataException("empty directory");
            _miniStream = ReadChain(entries[0].Start, entries[0].Size);
            _miniFat = ToUInts(ReadChain(BitConverter.ToUInt32(data, 60), BitConverter.ToUInt32(data, 64) * (long)_sectorSize));
            AddStreams(entries, entries[0].Child, "", new HashSet<uint>());
        }

        public IEnumerable<string> Names
        {
            get { return _streams.Keys; }
        }

        // Returns null when there is no such stream.
        public byte[] Read(string name)
        {
            Entry entry;
            if (!_streams.TryGetValue(name, out entry)) return null;
            return entry.Size < _miniCutoff ? ReadMiniChain(entry.Start, entry.Size) : ReadChain(entry.Start, entry.Size);
        }

        private uint[] ReadFat()
        {
            int fatSectors = (int)BitConverter.ToUInt32(_data, 44);
            var sectors = new List<uint>();
            for (int i = 0; i < 109 && sectors.Count < fatSectors; i++) sectors.Add(BitConverter.ToUInt32(_data, 76 + i * 4));

            // Big files list the rest of their FAT sectors in a chain of DIFAT sectors.
            uint next = BitConverter.ToUInt32(_data, 68);
            int perSector = _sectorSize / 4 - 1;
            for (int guard = 0; next < FirstSpecialSector && sectors.Count < fatSectors && guard < _data.Length / _sectorSize; guard++)
            {
                int offset = Offset(next);
                for (int i = 0; i < perSector && sectors.Count < fatSectors; i++) sectors.Add(BitConverter.ToUInt32(_data, offset + i * 4));
                next = BitConverter.ToUInt32(_data, offset + perSector * 4);
            }

            var fat = new uint[sectors.Count * (_sectorSize / 4)];
            for (int s = 0; s < sectors.Count; s++)
            {
                Buffer.BlockCopy(_data, Offset(sectors[s]), fat, s * _sectorSize, _sectorSize);
            }
            return fat;
        }

        private int Offset(uint sector)
        {
            long offset = (sector + 1L) * _sectorSize;
            if (offset + _sectorSize > _data.Length) throw new InvalidDataException("sector out of range");
            return (int)offset;
        }

        private List<Entry> ReadDirectory(uint start)
        {
            byte[] directory = ReadChain(start, -1);
            var entries = new List<Entry>();
            for (int at = 0; at + 128 <= directory.Length; at += 128)
            {
                int nameBytes = BitConverter.ToUInt16(directory, at + 64);
                entries.Add(new Entry
                {
                    Name = nameBytes >= 2 && nameBytes <= 64 ? Encoding.Unicode.GetString(directory, at, nameBytes - 2) : "",
                    Type = directory[at + 66],
                    Left = BitConverter.ToUInt32(directory, at + 68),
                    Right = BitConverter.ToUInt32(directory, at + 72),
                    Child = BitConverter.ToUInt32(directory, at + 76),
                    Start = BitConverter.ToUInt32(directory, at + 116),
                    Size = BitConverter.ToUInt32(directory, at + 120)
                });
            }
            return entries;
        }

        private void AddStreams(List<Entry> entries, uint id, string prefix, HashSet<uint> seen)
        {
            if (id >= entries.Count || !seen.Add(id)) return;
            Entry entry = entries[(int)id];
            AddStreams(entries, entry.Left, prefix, seen);
            if (entry.Type == 2) _streams[prefix + entry.Name] = entry;
            else if (entry.Type == 1) AddStreams(entries, entry.Child, prefix + entry.Name + "/", seen);
            AddStreams(entries, entry.Right, prefix, seen);
        }

        private byte[] ReadChain(uint start, long size)
        {
            using (var output = new MemoryStream())
            {
                uint sector = start;
                for (int guard = 0; sector < FirstSpecialSector && guard <= _fat.Length; guard++)
                {
                    output.Write(_data, Offset(sector), _sectorSize);
                    if (sector >= _fat.Length) throw new InvalidDataException("sector out of range");
                    sector = _fat[sector];
                }
                return Trim(output, size);
            }
        }

        private byte[] ReadMiniChain(uint start, long size)
        {
            using (var output = new MemoryStream())
            {
                uint sector = start;
                for (int guard = 0; sector < FirstSpecialSector && guard <= _miniFat.Length; guard++)
                {
                    long offset = sector * (long)_miniSectorSize;
                    if (offset + _miniSectorSize > _miniStream.Length || sector >= _miniFat.Length) throw new InvalidDataException("mini sector out of range");
                    output.Write(_miniStream, (int)offset, _miniSectorSize);
                    sector = _miniFat[sector];
                }
                return Trim(output, size);
            }
        }

        private static byte[] Trim(MemoryStream stream, long size)
        {
            byte[] all = stream.ToArray();
            if (size < 0 || size >= all.Length) return all;
            var result = new byte[size];
            Buffer.BlockCopy(all, 0, result, 0, (int)size);
            return result;
        }

        private static uint[] ToUInts(byte[] data)
        {
            var result = new uint[data.Length / 4];
            Buffer.BlockCopy(data, 0, result, 0, result.Length * 4);
            return result;
        }
    }

    // HWP 5.0, the binary format of .hwp and .hwt since Hangul 2002.
    // Output follows HwpxReader, because a document should read the same whichever way it was saved.
    internal sealed class HwpReader
    {
        private const int TagBinData = 0x12;
        private const int TagCharShape = 0x15;
        private const int TagParaShape = 0x19;
        private const int TagStyle = 0x1A;
        private const int TagParaHeader = 0x42;
        private const int TagParaText = 0x43;
        private const int TagParaCharShape = 0x44;
        private const int TagCtrlHeader = 0x47;
        private const int TagListHeader = 0x48;
        private const int TagTable = 0x4D;
        private const int TagPicture = 0x55;
        private const int TagEquation = 0x58;

        private const int HeadingNumber = 2;
        private const int HeadingBullet = 3;

        private static readonly uint CtrlTable = CtrlId("tbl ");
        private static readonly uint CtrlShape = CtrlId("gso ");
        private static readonly uint CtrlEquation = CtrlId("eqed");
        private static readonly uint CtrlFootnote = CtrlId("fn  ");
        private static readonly uint CtrlEndnote = CtrlId("en  ");

        private sealed class Record
        {
            public int Tag, Level;
            public byte[] Data;
            public readonly List<Record> Kids = new List<Record>();
        }

        private sealed class CharFormat
        {
            public bool Bold, Italic, Strike;
        }

        private sealed class BinaryItem
        {
            public string Stream;
            public int CompressMode;
        }

        private readonly CompoundFile _file;
        private readonly MdDoc _doc;
        private readonly MdImages _images;
        private readonly List<CharFormat> _charFormats = new List<CharFormat>();
        private readonly List<int> _paraHeadings = new List<int>();
        private readonly List<int> _styleHeadings = new List<int>();
        private readonly List<BinaryItem> _binaries = new List<BinaryItem>();
        private readonly Dictionary<int, int> _counters = new Dictionary<int, int>();
        private readonly List<List<Record>> _notes = new List<List<Record>>();
        private bool _compressed;

        public HwpReader(CompoundFile file, MdDoc doc, MdImages images)
        {
            _file = file;
            _doc = doc;
            _images = images;
        }

        private static uint CtrlId(string id)
        {
            return (uint)id[0] << 24 | (uint)id[1] << 16 | (uint)id[2] << 8 | id[3];
        }

        public void Run()
        {
            ReadFileHeader();
            ReadDocInfo();

            var sections = _file.Names
                .Select(n => Regex.Match(n, @"^BodyText/Section(\d+)$", RegexOptions.IgnoreCase))
                .Where(m => m.Success)
                .OrderBy(m => int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture))
                .Select(m => m.Value).ToList();
            if (sections.Count == 0) throw new InvalidOperationException("한글 문서 내용을 찾을 수 없습니다.");

            foreach (string section in sections)
            {
                foreach (Record record in ReadTree(section))
                {
                    if (record.Tag == TagParaHeader) Paragraph(record);
                }
            }
            for (int i = 0; i < _notes.Count; i++) _doc.AddFootnote(i + 1, CellText(_notes[i]));
        }

        private void ReadFileHeader()
        {
            byte[] header = _file.Read("FileHeader");
            if (header == null || header.Length < 40 || Encoding.ASCII.GetString(header, 0, 17) != "HWP Document File")
                throw new InvalidOperationException("한글 문서가 아니거나 손상된 파일입니다.");
            uint flags = BitConverter.ToUInt32(header, 36);
            _compressed = (flags & 1) != 0;
            if ((flags & 2) != 0) throw new InvalidOperationException("암호가 걸린 문서입니다. 한글에서 암호를 해제하고 저장한 뒤 변환해 주세요.");
            if ((flags & 4) != 0) throw new InvalidOperationException("배포용 문서는 복사가 제한되어 있어 변환할 수 없습니다.");
            if ((flags & 0x110) != 0) throw new InvalidOperationException("보안이 적용된 문서라 변환할 수 없습니다.");
        }

        private void ReadDocInfo()
        {
            byte[] info = ReadStream("DocInfo");
            if (info == null) throw new InvalidOperationException("한글 문서 내용을 찾을 수 없습니다.");
            foreach (Record record in ParseRecords(info))
            {
                switch (record.Tag)
                {
                    case TagBinData: AddBinary(record.Data); break;
                    case TagCharShape:
                        uint attributes = BitConverter.ToUInt32(record.Data, 46);
                        _charFormats.Add(new CharFormat { Italic = (attributes & 1) != 0, Bold = (attributes & 2) != 0, Strike = (attributes >> 18 & 7) != 0 });
                        break;
                    case TagParaShape: _paraHeadings.Add((int)(BitConverter.ToUInt32(record.Data, 0) >> 23 & 3)); break;
                    case TagStyle: _styleHeadings.Add(StyleHeading(record.Data)); break;
                }
            }
        }

        private void AddBinary(byte[] data)
        {
            int attributes = BitConverter.ToUInt16(data, 0);
            var item = new BinaryItem { CompressMode = attributes >> 4 & 3 };
            if ((attributes & 15) == 1)
            {
                int id = BitConverter.ToUInt16(data, 2);
                int extLength = BitConverter.ToUInt16(data, 4);
                string ext = Encoding.Unicode.GetString(data, 6, extLength * 2);
                item.Stream = "BinData/BIN" + id.ToString("X4", CultureInfo.InvariantCulture) + "." + ext;
            }
            _binaries.Add(item);
        }

        private static int StyleHeading(byte[] data)
        {
            int at = 0;
            string name = ReadWideString(data, ref at);
            string english = ReadWideString(data, ref at);
            return HangulFile.OutlineLevel(name, english);
        }

        private static string ReadWideString(byte[] data, ref int at)
        {
            int length = BitConverter.ToUInt16(data, at);
            string text = Encoding.Unicode.GetString(data, at + 2, length * 2);
            at += 2 + length * 2;
            return text;
        }

        private byte[] ReadStream(string name)
        {
            byte[] data = _file.Read(name);
            if (data == null) return null;
            return _compressed ? HangulFile.Inflate(data) : data;
        }

        private static List<Record> ParseRecords(byte[] data)
        {
            var records = new List<Record>();
            int at = 0;
            while (at + 4 <= data.Length)
            {
                uint header = BitConverter.ToUInt32(data, at);
                at += 4;
                int size = (int)(header >> 20);
                if (size == 0xFFF)
                {
                    size = BitConverter.ToInt32(data, at);
                    at += 4;
                }
                if (size < 0 || at + size > data.Length) throw new InvalidDataException("record out of range");
                var body = new byte[size];
                Buffer.BlockCopy(data, at, body, 0, size);
                at += size;
                records.Add(new Record { Tag = (int)(header & 0x3FF), Level = (int)(header >> 10 & 0x3FF), Data = body });
            }
            return records;
        }

        // A record belongs to the nearest earlier record with a lower level.
        private List<Record> ReadTree(string stream)
        {
            var roots = new List<Record>();
            var path = new Stack<Record>();
            foreach (Record record in ParseRecords(ReadStream(stream)))
            {
                while (path.Count > 0 && path.Peek().Level >= record.Level) path.Pop();
                if (path.Count > 0) path.Peek().Kids.Add(record);
                else roots.Add(record);
                path.Push(record);
            }
            return roots;
        }

        private void Paragraph(Record paragraph)
        {
            var inline = new MdInline();
            var after = new List<Record>();
            ReadText(paragraph, inline, after);

            int shape = BitConverter.ToUInt16(paragraph.Data, 8);
            int style = paragraph.Data[10];
            int headingType = shape < _paraHeadings.Count ? _paraHeadings[shape] : 0;
            int outline = style < _styleHeadings.Count ? _styleHeadings[style] : 0;

            if (outline > 0)
            {
                _doc.AddHeading(outline, inline);
            }
            else if (headingType == HeadingBullet)
            {
                _doc.AddListItem(0, false, 0, inline);
            }
            else if (headingType == HeadingNumber && !inline.IsEmpty)
            {
                int count;
                _counters.TryGetValue(shape, out count);
                _counters[shape] = count + 1;
                _doc.AddListItem(0, true, count + 1, inline);
            }
            else
            {
                _doc.AddParagraph(inline);
            }

            foreach (Record extra in after)
            {
                if (extra.Tag == TagCtrlHeader) Table(extra);
                else foreach (Record inner in extra.Kids) Paragraph(inner);
            }
        }

        private void ReadText(Record paragraph, MdInline output, List<Record> after)
        {
            Record text = paragraph.Kids.FirstOrDefault(k => k.Tag == TagParaText);
            Record shapes = paragraph.Kids.FirstOrDefault(k => k.Tag == TagParaCharShape);
            var controls = new Queue<Record>(paragraph.Kids.Where(k => k.Tag == TagCtrlHeader));

            if (text != null)
            {
                int nextShape = 0;
                CharFormat format = new CharFormat();
                var pending = new StringBuilder();
                int length = text.Data.Length / 2;
                for (int i = 0; i < length; )
                {
                    while (shapes != null && nextShape + 8 <= shapes.Data.Length && BitConverter.ToUInt32(shapes.Data, nextShape) <= i)
                    {
                        Flush(pending, output, format);
                        format = Format(BitConverter.ToUInt32(shapes.Data, nextShape + 4));
                        nextShape += 8;
                    }

                    char c = BitConverter.ToChar(text.Data, i * 2);
                    if (c >= 32)
                    {
                        pending.Append(c);
                        i++;
                        continue;
                    }

                    switch ((int)c)
                    {
                        case 10:
                            Flush(pending, output, format);
                            output.AddBreak();
                            i++;
                            break;
                        case 24: pending.Append('-'); i++; break;
                        case 30:
                        case 31: pending.Append(' '); i++; break;
                        case 0:
                        case 13: i++; break;
                        case 9: pending.Append(' '); i += 8; break;
                        case 4:
                        case 5:
                        case 6:
                        case 7:
                        case 8:
                        case 19:
                        case 20: i += 8; break;
                        default:
                            // The other codes stand for controls, whose details are the next CTRL_HEADER.
                            Flush(pending, output, format);
                            if (controls.Count > 0) Control(controls.Dequeue(), output, after);
                            i += 8;
                            break;
                    }
                }
                Flush(pending, output, format);
            }

            while (controls.Count > 0) Control(controls.Dequeue(), output, after);
        }

        private static void Flush(StringBuilder pending, MdInline output, CharFormat format)
        {
            if (pending.Length == 0) return;
            output.Add(pending.ToString(), format.Bold, format.Italic, format.Strike);
            pending.Clear();
        }

        private CharFormat Format(uint id)
        {
            return id < _charFormats.Count ? _charFormats[(int)id] : new CharFormat();
        }

        private void Control(Record control, MdInline output, List<Record> after)
        {
            if (control.Data.Length < 4) return;
            uint id = BitConverter.ToUInt32(control.Data, 0);
            if (id == CtrlTable)
            {
                after.Add(control);
            }
            else if (id == CtrlFootnote || id == CtrlEndnote)
            {
                _notes.Add(ListParagraphs(control.Kids, 0));
                output.AddRaw("[^" + _notes.Count + "]");
            }
            else if (id == CtrlEquation)
            {
                Record equation = control.Kids.FirstOrDefault(k => k.Tag == TagEquation);
                if (equation != null && equation.Data.Length >= 6)
                {
                    int at = 4;
                    output.Add(ReadWideString(equation.Data, ref at));
                }
            }
            else if (id == CtrlShape)
            {
                Shape(control, output, after);
            }
        }

        // Shapes nest in groups, and each may hold a picture or a list of paragraphs.
        private void Shape(Record parent, MdInline output, List<Record> after)
        {
            for (int i = 0; i < parent.Kids.Count; i++)
            {
                Record kid = parent.Kids[i];
                if (kid.Tag == TagPicture)
                {
                    Picture(kid, output);
                }
                else if (kid.Tag == TagListHeader)
                {
                    var list = new Record { Tag = TagListHeader };
                    list.Kids.AddRange(ListParagraphs(parent.Kids, i));
                    after.Add(list);
                }
                else
                {
                    Shape(kid, output, after);
                }
            }
        }

        // The paragraphs of a list come right after its LIST_HEADER, as its siblings.
        private static List<Record> ListParagraphs(List<Record> records, int from)
        {
            int start = from;
            while (start < records.Count && records[start].Tag != TagListHeader) start++;
            var paragraphs = new List<Record>();
            for (int i = start + 1; i < records.Count && records[i].Tag == TagParaHeader; i++) paragraphs.Add(records[i]);
            return paragraphs;
        }

        private void Picture(Record picture, MdInline output)
        {
            const int BinItemOffset = 71;
            if (picture.Data.Length < BinItemOffset + 2) return;
            int index = BitConverter.ToUInt16(picture.Data, BinItemOffset) - 1;
            if (index < 0 || index >= _binaries.Count || _binaries[index].Stream == null) return;

            BinaryItem item = _binaries[index];
            byte[] data = _file.Read(item.Stream);
            if (data == null) return;
            bool compressed = item.CompressMode == 1 || (item.CompressMode == 0 && _compressed);
            if (compressed) data = HangulFile.InflateOrKeep(data);
            string link = _images.SaveEmbedded(data, item.Stream);
            if (link != null) output.AddRaw(MdImages.Markdown(null, link));
        }

        private void Table(Record control)
        {
            var rows = new List<List<string>>();
            var cells = new List<List<Record>>();
            int table = control.Kids.FindIndex(k => k.Tag == TagTable);
            for (int i = table + 1; table >= 0 && i < control.Kids.Count; i++)
            {
                Record kid = control.Kids[i];
                if (kid.Tag != TagListHeader || kid.Data.Length < 12) continue;
                int column = BitConverter.ToUInt16(kid.Data, 8);
                int rowIndex = BitConverter.ToUInt16(kid.Data, 10);
                List<Record> paragraphs = ListParagraphs(control.Kids, i);
                while (rows.Count <= rowIndex) rows.Add(new List<string>());
                while (rows[rowIndex].Count <= column) rows[rowIndex].Add("");
                rows[rowIndex][column] = CellText(paragraphs);
                cells.Add(paragraphs);
            }

            // A one-cell table is how Hangul usually draws a box around text.
            if (rows.All(r => r.Count <= 1))
            {
                foreach (List<Record> cell in cells)
                {
                    foreach (Record paragraph in cell) Paragraph(paragraph);
                }
                return;
            }
            _doc.AddTable(rows);
        }

        private string CellText(List<Record> paragraphs)
        {
            var lines = new List<string>();
            foreach (Record paragraph in paragraphs)
            {
                var inline = new MdInline();
                var after = new List<Record>();
                ReadText(paragraph, inline, after);
                string text = inline.Render(false);
                if (text.Length > 0) lines.Add(text);
                foreach (Record extra in after)
                {
                    if (extra.Tag == TagCtrlHeader)
                    {
                        for (int i = 0; i < extra.Kids.Count; i++)
                        {
                            if (extra.Kids[i].Tag != TagListHeader) continue;
                            string inner = CellText(ListParagraphs(extra.Kids, i)).Replace("  \n", " ");
                            if (inner.Length > 0) lines.Add(inner);
                        }
                    }
                    else
                    {
                        string inner = CellText(extra.Kids);
                        if (inner.Length > 0) lines.Add(inner);
                    }
                }
            }
            return string.Join("  \n", lines);
        }
    }

    // HWPML (.hml, .hwpml): one XML file holding the whole document, pictures included.
    internal sealed class HmlReader
    {
        private sealed class CharFormat
        {
            public bool Bold, Italic, Strike;
        }

        private readonly XElement _root;
        private readonly MdDoc _doc;
        private readonly MdImages _images;
        private readonly Dictionary<string, CharFormat> _charFormats = new Dictionary<string, CharFormat>();
        private readonly Dictionary<string, int> _styleHeadings = new Dictionary<string, int>();
        private readonly Dictionary<string, string> _paraHeadingTypes = new Dictionary<string, string>();
        private readonly Dictionary<string, XElement> _binaries = new Dictionary<string, XElement>();
        private readonly Dictionary<string, string> _binaryExts = new Dictionary<string, string>();
        private readonly Dictionary<string, int> _counters = new Dictionary<string, int>();
        private readonly List<XElement> _notes = new List<XElement>();

        public HmlReader(byte[] data, MdDoc doc, MdImages images)
        {
            _doc = doc;
            _images = images;
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using (var stream = new MemoryStream(data))
            using (XmlReader reader = XmlReader.Create(stream, settings))
            {
                _root = XDocument.Load(reader).Root;
            }
            if (X.Local(_root) != "HWPML") throw new InvalidOperationException("한글 문서가 아니거나 손상된 파일입니다.");
        }

        public void Run()
        {
            ReadHead();
            foreach (XElement section in X.Below(X.Kid(_root, "BODY"), "SECTION")) WalkBlocks(section);
            for (int i = 0; i < _notes.Count; i++) _doc.AddFootnote(i + 1, CellText(_notes[i]));
        }

        private void ReadHead()
        {
            XElement head = X.Kid(_root, "HEAD");
            foreach (XElement format in X.Below(head, "CHARSHAPE"))
            {
                XElement strike = X.Kid(format, "STRIKEOUT");
                _charFormats[X.Attr(format, "Id") ?? ""] = new CharFormat
                {
                    Bold = X.Kid(format, "BOLD") != null,
                    Italic = X.Kid(format, "ITALIC") != null,
                    Strike = strike != null && (X.Attr(strike, "Type") ?? "") != "None"
                };
            }
            foreach (XElement style in X.Below(head, "STYLE"))
            {
                int level = HangulFile.OutlineLevel(X.Attr(style, "Name"), X.Attr(style, "EngName"));
                if (level > 0) _styleHeadings[X.Attr(style, "Id") ?? ""] = level;
            }
            foreach (XElement para in X.Below(head, "PARASHAPE")) _paraHeadingTypes[X.Attr(para, "Id") ?? ""] = X.Attr(para, "HeadingType") ?? "None";
            foreach (XElement item in X.Below(head, "BINITEM"))
            {
                string id = X.Attr(item, "BinData");
                if (id != null) _binaryExts[id] = X.Attr(item, "Format") ?? "";
            }
            foreach (XElement data in X.Below(X.Kid(_root, "TAIL"), "BINDATA")) _binaries[X.Attr(data, "Id") ?? ""] = data;
        }

        private void WalkBlocks(XElement parent)
        {
            foreach (XElement child in parent.Elements())
            {
                string kind = X.Local(child);
                if (kind == "P") Paragraph(child);
                else if (kind == "PARALIST") WalkBlocks(child);
            }
        }

        private void Paragraph(XElement p)
        {
            var inline = new MdInline();
            var after = new List<XElement>();
            foreach (XElement run in X.Kids(p, "TEXT")) Run(run, inline, after);

            int heading;
            string headingType;
            string shape = X.Attr(p, "ParaShape") ?? "";
            _paraHeadingTypes.TryGetValue(shape, out headingType);
            if (_styleHeadings.TryGetValue(X.Attr(p, "Style") ?? "", out heading))
            {
                _doc.AddHeading(heading, inline);
            }
            else if (headingType == "Bullet")
            {
                _doc.AddListItem(0, false, 0, inline);
            }
            else if (headingType == "Number" && !inline.IsEmpty)
            {
                int count;
                _counters.TryGetValue(shape, out count);
                _counters[shape] = count + 1;
                _doc.AddListItem(0, true, count + 1, inline);
            }
            else
            {
                _doc.AddParagraph(inline);
            }

            foreach (XElement extra in after)
            {
                if (X.Local(extra) == "TABLE") Table(extra);
                else WalkBlocks(extra);
            }
        }

        private void Run(XElement run, MdInline output, List<XElement> after)
        {
            CharFormat format;
            if (!_charFormats.TryGetValue(X.Attr(run, "CharShape") ?? "", out format)) format = new CharFormat();
            foreach (XElement child in run.Elements())
            {
                switch (X.Local(child))
                {
                    case "CHAR": Text(child, output, format); break;
                    case "TABLE": after.Add(child); break;
                    case "PICTURE": Picture(child, output); break;
                    case "EQUATION":
                        XElement script = X.Kid(child, "SCRIPT");
                        if (script != null) output.Add(script.Value);
                        break;
                    case "FOOTNOTE":
                    case "ENDNOTE":
                        _notes.Add(X.Kid(child, "PARALIST") ?? child);
                        output.AddRaw("[^" + _notes.Count + "]");
                        break;
                    case "RECTANGLE":
                    case "ELLIPSE":
                    case "POLYGON":
                    case "CURVE":
                    case "ARC":
                    case "CONTAINER":
                        foreach (XElement list in X.Below(child, "PARALIST")) after.Add(list);
                        foreach (XElement picture in X.Below(child, "PICTURE")) Picture(picture, output);
                        break;
                }
            }
        }

        private static void Text(XElement text, MdInline output, CharFormat format)
        {
            foreach (XNode node in text.Nodes())
            {
                var plain = node as XText;
                if (plain != null)
                {
                    output.Add(plain.Value, format.Bold, format.Italic, format.Strike);
                    continue;
                }
                var element = node as XElement;
                if (element == null) continue;
                switch (X.Local(element))
                {
                    case "TAB":
                    case "NBSPACE":
                    case "FWSPACE":
                        output.Add(" ", format.Bold, format.Italic, format.Strike);
                        break;
                    case "LINEBREAK": output.AddBreak(); break;
                    case "HYPHEN":
                    case "HYPEN":
                        output.Add("-", format.Bold, format.Italic, format.Strike);
                        break;
                }
            }
        }

        private void Picture(XElement picture, MdInline output)
        {
            XElement image = X.Kid(picture, "IMAGE");
            string id = image == null ? null : X.Attr(image, "BinItem");
            XElement stored;
            if (id == null || !_binaries.TryGetValue(id, out stored)) return;

            byte[] data = Convert.FromBase64String(stored.Value);
            if ((X.Attr(stored, "Compress") ?? "").Equals("true", StringComparison.OrdinalIgnoreCase)) data = HangulFile.InflateOrKeep(data);
            string ext;
            _binaryExts.TryGetValue(id, out ext);
            string link = _images.SaveEmbedded(data, "picture." + (ext ?? "").ToLowerInvariant());
            if (link != null) output.AddRaw(MdImages.Markdown(null, link));
        }

        private void Table(XElement table)
        {
            var rows = new List<List<string>>();
            var cells = new List<XElement>();
            foreach (XElement row in X.Kids(table, "ROW"))
            {
                foreach (XElement cell in X.Kids(row, "CELL"))
                {
                    int rowIndex = X.IntAttr(cell, "RowAddr", rows.Count);
                    int column = X.IntAttr(cell, "ColAddr", 0);
                    while (rows.Count <= rowIndex) rows.Add(new List<string>());
                    while (rows[rowIndex].Count <= column) rows[rowIndex].Add("");
                    rows[rowIndex][column] = CellText(X.Kid(cell, "PARALIST") ?? cell);
                    cells.Add(cell);
                }
            }

            // A one-cell table is how Hangul usually draws a box around text.
            if (rows.All(r => r.Count <= 1))
            {
                foreach (XElement cell in cells) WalkBlocks(X.Kid(cell, "PARALIST") ?? cell);
                return;
            }
            _doc.AddTable(rows);
        }

        private string CellText(XElement container)
        {
            var lines = new List<string>();
            foreach (XElement p in X.Kids(container, "P"))
            {
                var inline = new MdInline();
                var after = new List<XElement>();
                foreach (XElement run in X.Kids(p, "TEXT")) Run(run, inline, after);
                string text = inline.Render(false);
                if (text.Length > 0) lines.Add(text);
                foreach (XElement extra in after)
                {
                    if (X.Local(extra) == "TABLE")
                    {
                        foreach (XElement cell in X.Below(extra, "CELL"))
                        {
                            string inner = CellText(X.Kid(cell, "PARALIST") ?? cell).Replace("  \n", " ");
                            if (inner.Length > 0) lines.Add(inner);
                        }
                    }
                    else
                    {
                        string inner = CellText(extra);
                        if (inner.Length > 0) lines.Add(inner);
                    }
                }
            }
            return string.Join("  \n", lines);
        }
    }

    // HWP 3.x, the format of Hangul 3.0 to 97. A paragraph's controls are followed by their own paragraph lists,
    // so the file can only be walked from the start, and embedded pictures are stored after all of the text.
    internal sealed class Hwp3Reader
    {
        private const int HeaderSize = 30 + 128 + 1008;
        private const int ParaInfoSize = 43;
        private const int ParaShapeSize = 187;
        private const int CharShapeSize = 31;
        private const int LineInfoSize = 14;
        private const int StyleSize = 238;

        private static readonly Dictionary<int, char> Symbols = LoadSymbols();
        private static readonly int[] JohabVowels = { -1, -1, -1, 0, 1, 2, 3, 4, -1, -1, 5, 6, 7, 8, 9, 10, -1, -1, 11, 12, 13, 14, 15, 16, -1, -1, 17, 18, 19, 20 };

        private sealed class CharFormat
        {
            public bool Bold, Italic;
        }

        private enum PieceKind { Text, Note, Table, Box, Picture, Equation }

        private sealed class Cell
        {
            public int X, Y;
            public List<Para> Paragraphs;
        }

        private sealed class Piece
        {
            public PieceKind Kind;
            public string Text;
            public CharFormat Format;
            public List<Para> Paragraphs;
            public List<Cell> Cells;
        }

        private sealed class Para
        {
            public int Style;
            public readonly List<Piece> Pieces = new List<Piece>();
        }

        private readonly byte[] _file;
        private readonly MdDoc _doc;
        private readonly MdImages _images;
        private readonly List<int> _styleHeadings = new List<int>();
        private readonly Dictionary<string, KeyValuePair<string, byte[]>> _pictures = new Dictionary<string, KeyValuePair<string, byte[]>>(StringComparer.OrdinalIgnoreCase);
        private readonly List<List<Para>> _notes = new List<List<Para>>();
        private byte[] _data;
        private int _at;

        public Hwp3Reader(byte[] file, MdDoc doc, MdImages images)
        {
            _file = file;
            _doc = doc;
            _images = images;
        }

        private static Dictionary<int, char> LoadSymbols()
        {
            var symbols = new Dictionary<int, char>();
            string table = Hwp3Codes.Symbols;
            for (int i = 0; i + 8 <= table.Length; i += 8)
            {
                symbols[int.Parse(table.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture)] =
                    (char)int.Parse(table.Substring(i + 4, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            }
            return symbols;
        }

        public void Run()
        {
            if (_file.Length < HeaderSize || Encoding.ASCII.GetString(_file, 0, 23) != "HWP Document File V3.00")
                throw new InvalidOperationException("한글 3.0보다 오래된 형식입니다. 한글에서 다시 저장한 뒤 변환해 주세요.");

            const int Info = 30;
            if (BitConverter.ToUInt16(_file, Info + 96) != 0)
                throw new InvalidOperationException("암호가 걸린 문서입니다. 한글에서 암호를 해제하고 저장한 뒤 변환해 주세요.");
            bool compressed = _file[Info + 124] != 0;
            int start = HeaderSize + BitConverter.ToUInt16(_file, Info + 126);

            var body = new byte[_file.Length - start];
            Buffer.BlockCopy(_file, start, body, 0, body.Length);
            // Block #2 follows the compressed stream uncompressed, and the inflater just stops before it.
            _data = compressed ? HangulFile.Inflate(body) : body;

            SkipFonts();
            ReadStyles();
            List<Para> paragraphs = ReadList();
            ReadExtraBlocks();

            foreach (Para paragraph in paragraphs) Paragraph(paragraph);
            // Hangul writes the note's number as a control, which leaves only its closing bracket as text.
            for (int i = 0; i < _notes.Count; i++) _doc.AddFootnote(i + 1, Regex.Replace(CellText(_notes[i]), @"^\)\s*", ""));
        }

        private void SkipFonts()
        {
            for (int language = 0; language < 7; language++) _at += 2 + 40 * Word(_at);
        }

        private void ReadStyles()
        {
            int count = Word(_at);
            _at += 2;
            Encoding johab = Encoding.GetEncoding(1361);
            for (int i = 0; i < count; i++)
            {
                string name = johab.GetString(_data, _at, 20);
                int end = name.IndexOf('\0');
                _styleHeadings.Add(HangulFile.OutlineLevel(end >= 0 ? name.Substring(0, end) : name, null));
                _at += StyleSize;
            }
        }

        private int Word(int at)
        {
            return BitConverter.ToUInt16(_data, at);
        }

        // A list ends with a paragraph of zero characters.
        private List<Para> ReadList()
        {
            var list = new List<Para>();
            while (true)
            {
                bool reuseShape = _data[_at] != 0;
                int characters = Word(_at + 1);
                int lines = Word(_at + 3);
                bool hasShapes = _data[_at + 5] != 0;
                int style = _data[_at + 11];
                CharFormat format = ReadCharFormat(_at + 12);
                _at += ParaInfoSize;
                if (characters == 0) return list;

                if (!reuseShape) _at += ParaShapeSize;
                _at += lines * LineInfoSize;

                var formats = new CharFormat[characters];
                for (int i = 0; i < characters; i++)
                {
                    if (hasShapes)
                    {
                        if (_data[_at++] != 1)
                        {
                            format = ReadCharFormat(_at);
                            _at += CharShapeSize;
                        }
                    }
                    formats[i] = format;
                }

                var paragraph = new Para { Style = style };
                ReadCharacters(paragraph, formats);
                list.Add(paragraph);
            }
        }

        private CharFormat ReadCharFormat(int at)
        {
            byte attributes = _data[at + 26];
            return new CharFormat { Italic = (attributes & 1) != 0, Bold = (attributes & 2) != 0 };
        }

        private void ReadCharacters(Para paragraph, CharFormat[] formats)
        {
            var pending = new StringBuilder();
            CharFormat pendingFormat = null;
            int index = 0;
            while (true)
            {
                int code = Word(_at);
                CharFormat format = formats[Math.Min(index, formats.Length - 1)];
                if (code == 13)
                {
                    _at += 2;
                    FlushText(paragraph, pending, pendingFormat);
                    return;
                }

                if (code >= 32)
                {
                    if (pendingFormat != format && pending.Length > 0) FlushText(paragraph, pending, pendingFormat);
                    pendingFormat = format;
                    AppendCharacter(pending, code);
                    _at += 2;
                    index++;
                    continue;
                }

                FlushText(paragraph, pending, pendingFormat);
                int start = _at;
                Control(code, paragraph, format);
                index += (_at - start) / 2;
            }
        }

        private static void FlushText(Para paragraph, StringBuilder pending, CharFormat format)
        {
            if (pending.Length == 0) return;
            paragraph.Pieces.Add(new Piece { Kind = PieceKind.Text, Text = pending.ToString(), Format = format });
            pending.Clear();
        }

        private void Control(int code, Para paragraph, CharFormat format)
        {
            switch (code)
            {
                case 9: AddSpace(paragraph, format, 8); break;
                case 30:
                case 31: AddSpace(paragraph, format, 4); break;
                case 24:
                    paragraph.Pieces.Add(new Piece { Kind = PieceKind.Text, Text = "-", Format = format });
                    _at += 6;
                    break;
                case 7: _at += 84; break;
                case 8: _at += 96; break;
                case 18:
                case 19:
                case 20:
                case 21: _at += 8; break;
                case 22: _at += 24; break;
                case 23: _at += 10; break;
                case 25: _at += 6; break;
                case 26: _at += 246; break;
                case 28: _at += 64; break;
                case 10: ReadBox(paragraph); break;
                case 11: ReadPicture(paragraph); break;
                case 14: _at += 8 + 84; break;
                case 15:
                    _at += 8 + 8;
                    ReadList();
                    break;
                case 16:
                    _at += 8 + 10;
                    ReadList();
                    break;
                case 17:
                    _at += 8 + 14;
                    paragraph.Pieces.Add(new Piece { Kind = PieceKind.Note, Paragraphs = ReadList() });
                    break;
                default:
                    // Field codes, bookmarks, cross references and reserved codes all state their own length.
                    _at += 8 + checked((int)BitConverter.ToUInt32(_data, _at + 2));
                    break;
            }
        }

        private void AddSpace(Para paragraph, CharFormat format, int size)
        {
            paragraph.Pieces.Add(new Piece { Kind = PieceKind.Text, Text = " ", Format = format });
            _at += size;
        }

        private static void AppendCharacter(StringBuilder output, int code)
        {
            char symbol;
            if (code >= 0x8000)
            {
                char hangul = Johab(code);
                if (hangul != '\0') output.Append(hangul);
            }
            else if (Symbols.TryGetValue(code, out symbol))
            {
                output.Append(symbol);
            }
            else if (code >= 0x4000 && code < 0x4000 + 4888)
            {
                // Hanja keep the KS X 1001 order.
                int n = code - 0x4000;
                output.Append(Encoding.GetEncoding(949).GetString(new[] { (byte)(0xCA + n / 94), (byte)(0xA1 + n % 94) }));
            }
            else
            {
                output.Append((char)code);
            }
        }

        // Five bits each for the first, middle and last sounds of a syllable.
        private static char Johab(int code)
        {
            int first = (code >> 10 & 31) - 2;
            int middle = code >> 5 & 31;
            int last = code & 31;
            int vowel = middle < JohabVowels.Length ? JohabVowels[middle] : -1;
            int final = last == 1 ? 0 : last >= 2 && last <= 17 ? last - 1 : last >= 19 && last <= 29 ? last - 2 : -1;
            if (first < 0 || first > 18 || vowel < 0 || final < 0) return '\0';
            return (char)(0xAC00 + first * 588 + vowel * 28 + final);
        }

        private void ReadBox(Para paragraph)
        {
            int info = _at + 8;
            int kind = Word(info + 78);
            int count = Word(info + 80);
            int cellInfo = info + 84;
            _at = cellInfo + count * 27;

            var cells = new List<Cell>();
            for (int i = 0; i < count; i++)
            {
                int at = cellInfo + i * 27;
                cells.Add(new Cell { X = Word(at + 4), Y = Word(at + 6), Paragraphs = ReadList() });
            }
            ReadList();

            const int KindTable = 0;
            const int KindEquation = 2;
            if (kind == KindEquation)
            {
                string script = cells.Count == 0 ? "" : CellText(cells[0].Paragraphs);
                paragraph.Pieces.Add(new Piece { Kind = PieceKind.Equation, Text = script.Replace("  \n", " ") });
            }
            else if (kind == KindTable)
            {
                paragraph.Pieces.Add(new Piece { Kind = PieceKind.Table, Cells = cells });
            }
            else
            {
                paragraph.Pieces.Add(new Piece { Kind = PieceKind.Box, Paragraphs = cells.SelectMany(c => c.Paragraphs).ToList() });
            }
        }

        private void ReadPicture(Para paragraph)
        {
            int info = _at + 8;
            int extra = checked((int)BitConverter.ToUInt32(_data, info));
            const int KindEmbedded = 2;
            if (_data[info + 74] == KindEmbedded)
                paragraph.Pieces.Add(new Piece { Kind = PieceKind.Picture, Text = AsciiZ(info + 83, 256) });
            _at = info + 348 + extra;
            ReadList();
        }

        private string AsciiZ(int at, int length)
        {
            int end = at;
            while (end < at + length && _data[end] != 0) end++;
            return Encoding.GetEncoding(949).GetString(_data, at, end - at);
        }

        private void ReadExtraBlocks()
        {
            const int BlockPicture = 1;
            while (_at + 8 <= _data.Length)
            {
                uint id = BitConverter.ToUInt32(_data, _at);
                int length = checked((int)BitConverter.ToUInt32(_data, _at + 4));
                if (id == 0 && length == 0) return;
                if (id == BlockPicture && length >= 32)
                {
                    string name = AsciiZ(_at + 8, 16);
                    string format = AsciiZ(_at + 24, 16);
                    var picture = new byte[length - 32];
                    Buffer.BlockCopy(_data, _at + 40, picture, 0, picture.Length);
                    _pictures[name] = new KeyValuePair<string, byte[]>(format, picture);
                }
                _at += 8 + length;
            }
        }

        private void Paragraph(Para paragraph)
        {
            var inline = new MdInline();
            var after = new List<Piece>();
            Render(paragraph, inline, after);

            int outline = paragraph.Style < _styleHeadings.Count ? _styleHeadings[paragraph.Style] : 0;
            if (outline > 0) _doc.AddHeading(outline, inline);
            else _doc.AddParagraph(inline);

            foreach (Piece extra in after)
            {
                if (extra.Kind == PieceKind.Table) Table(extra.Cells);
                else foreach (Para inner in extra.Paragraphs) Paragraph(inner);
            }
        }

        private void Render(Para paragraph, MdInline output, List<Piece> after)
        {
            foreach (Piece piece in paragraph.Pieces)
            {
                switch (piece.Kind)
                {
                    case PieceKind.Text: output.Add(piece.Text, piece.Format.Bold, piece.Format.Italic, false); break;
                    case PieceKind.Equation: output.Add(piece.Text); break;
                    case PieceKind.Note:
                        _notes.Add(piece.Paragraphs);
                        output.AddRaw("[^" + _notes.Count + "]");
                        break;
                    case PieceKind.Picture: Picture(piece.Text, output); break;
                    default: after.Add(piece); break;
                }
            }
        }

        private void Picture(string name, MdInline output)
        {
            KeyValuePair<string, byte[]> picture;
            if (!_pictures.TryGetValue(name, out picture)) return;
            string link = _images.SaveEmbedded(picture.Value, name + "." + picture.Key);
            if (link != null) output.AddRaw(MdImages.Markdown(null, link));
        }

        // Cells only store where they sit, so rows and columns come from the distinct positions.
        private void Table(List<Cell> cells)
        {
            List<int> columns = cells.Select(c => c.X).Distinct().OrderBy(x => x).ToList();
            List<int> rowTops = cells.Select(c => c.Y).Distinct().OrderBy(y => y).ToList();
            if (columns.Count <= 1)
            {
                foreach (Cell cell in cells.OrderBy(c => c.Y))
                {
                    foreach (Para paragraph in cell.Paragraphs) Paragraph(paragraph);
                }
                return;
            }

            // Merged cells leave gaps in the grid.
            var rows = rowTops.Select(y => Enumerable.Repeat("", columns.Count).ToList()).ToList();
            foreach (Cell cell in cells) rows[rowTops.IndexOf(cell.Y)][columns.IndexOf(cell.X)] = CellText(cell.Paragraphs);
            _doc.AddTable(rows);
        }

        private string CellText(List<Para> paragraphs)
        {
            var lines = new List<string>();
            foreach (Para paragraph in paragraphs)
            {
                var inline = new MdInline();
                var after = new List<Piece>();
                Render(paragraph, inline, after);
                string text = inline.Render(false);
                if (text.Length > 0) lines.Add(text);
                foreach (Piece extra in after)
                {
                    IEnumerable<List<Para>> lists = extra.Kind == PieceKind.Table ? extra.Cells.Select(c => c.Paragraphs) : new[] { extra.Paragraphs };
                    foreach (List<Para> list in lists)
                    {
                        string inner = CellText(list).Replace("  \n", " ");
                        if (inner.Length > 0) lines.Add(inner);
                    }
                }
            }
            return string.Join("  \n", lines);
        }
    }
}
