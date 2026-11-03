using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace Fanak
{
    // Word, PowerPoint, Excel and Hancom's .hwpx are all zip files of XML, so none of those programs is needed.
    internal sealed class DocPackage : IDisposable
    {
        public sealed class Relation
        {
            public string Type, Target;
            public bool External;
        }

        private readonly FileStream _stream;
        private readonly ZipArchive _zip;
        private readonly Dictionary<string, ZipArchiveEntry> _entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);

        public DocPackage(string path)
        {
            // Sharing allows reading a document that is open in Word at the same time.
            _stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            try
            {
                _zip = new ZipArchive(_stream, ZipArchiveMode.Read);
            }
            catch
            {
                _stream.Dispose();
                throw;
            }
            foreach (ZipArchiveEntry entry in _zip.Entries) _entries[entry.FullName.Replace('\\', '/')] = entry;
        }

        public void Dispose()
        {
            _zip.Dispose();
            _stream.Dispose();
        }

        public IEnumerable<string> Names { get { return _entries.Keys; } }

        public byte[] Bytes(string name)
        {
            ZipArchiveEntry entry;
            if (!_entries.TryGetValue(name, out entry)) return null;
            using (Stream input = entry.Open())
            using (var output = new MemoryStream())
            {
                input.CopyTo(output);
                return output.ToArray();
            }
        }

        public XElement Xml(string name)
        {
            ZipArchiveEntry entry;
            if (!_entries.TryGetValue(name, out entry)) return null;
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using (Stream input = entry.Open())
            using (XmlReader reader = XmlReader.Create(input, settings))
            {
                return XDocument.Load(reader, LoadOptions.PreserveWhitespace).Root;
            }
        }

        public Dictionary<string, Relation> Relations(string part)
        {
            var result = new Dictionary<string, Relation>();
            int slash = part.LastIndexOf('/');
            string folder = slash < 0 ? "" : part.Substring(0, slash + 1);
            string file = slash < 0 ? part : part.Substring(slash + 1);
            XElement xml = Xml(folder + "_rels/" + file + ".rels");
            if (xml == null) return result;
            foreach (XElement rel in xml.Elements())
            {
                bool external = X.Attr(rel, "TargetMode") == "External";
                string target = X.Attr(rel, "Target") ?? "";
                result[X.Attr(rel, "Id") ?? ""] = new Relation
                {
                    Type = X.Attr(rel, "Type") ?? "",
                    External = external,
                    Target = external ? target : Resolve(folder, target)
                };
            }
            return result;
        }

        private static string Resolve(string folder, string target)
        {
            string combined = target.StartsWith("/") ? target.TrimStart('/') : folder + target;
            var parts = new List<string>();
            foreach (string part in combined.Split('/'))
            {
                if (part == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
                else if (part != "." && part.Length > 0) parts.Add(part);
            }
            return string.Join("/", parts);
        }
    }

    // Element names are matched without namespaces, since each format is read by its own class anyway.
    internal static class X
    {
        public static string Local(XElement e) { return e.Name.LocalName; }

        public static IEnumerable<XElement> Kids(XElement e, string name)
        {
            return e == null ? Enumerable.Empty<XElement>() : e.Elements().Where(k => k.Name.LocalName == name);
        }

        public static XElement Kid(XElement e, string name) { return Kids(e, name).FirstOrDefault(); }

        public static IEnumerable<XElement> Below(XElement e, string name)
        {
            return e == null ? Enumerable.Empty<XElement>() : e.Descendants().Where(k => k.Name.LocalName == name);
        }

        public static string Attr(XElement e, string name)
        {
            if (e == null) return null;
            XAttribute attribute = e.Attributes().FirstOrDefault(a => a.Name.LocalName == name);
            return attribute == null ? null : attribute.Value;
        }

        // A relationship id is r:id, which some elements have next to a plain id meaning something else.
        public static string RelId(XElement e)
        {
            if (e == null) return null;
            XAttribute attribute = e.Attributes().FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None);
            return attribute == null ? null : attribute.Value;
        }

        public static int IntAttr(XElement e, string name, int fallback)
        {
            int value;
            return int.TryParse(Attr(e, name), NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : fallback;
        }

        // <w:b/> means on, and <w:b w:val="0"/> means off.
        public static bool Toggle(XElement e)
        {
            if (e == null) return false;
            string value = Attr(e, "val");
            return value != "0" && value != "false" && value != "none";
        }

        // Newer drawing features come with a fallback copy of the same content; reading both would double it.
        public static XElement ChosenBranch(XElement alternate)
        {
            return Kid(alternate, "Choice") ?? Kid(alternate, "Fallback");
        }
    }

    internal sealed class DocxReader
    {
        private sealed class Style
        {
            public int Heading;
            public string BasedOn, NumId;
            public int Level;
            public bool HasLevel, Bold, Italic;
        }

        private sealed class Format
        {
            public bool Bold, Italic, Strike;
        }

        private readonly DocPackage _package;
        private readonly MdDoc _doc;
        private readonly MdImages _images;
        private readonly Dictionary<string, Style> _styles = new Dictionary<string, Style>();
        private readonly Dictionary<string, string> _numToAbstract = new Dictionary<string, string>();
        private readonly Dictionary<string, Dictionary<int, XElement>> _abstractLevels = new Dictionary<string, Dictionary<int, XElement>>();
        private readonly Dictionary<string, int[]> _counters = new Dictionary<string, int[]>();
        private readonly Dictionary<string, int> _noteNumbers = new Dictionary<string, int>();
        private Dictionary<string, DocPackage.Relation> _rels;

        public DocxReader(DocPackage package, MdDoc doc, MdImages images)
        {
            _package = package;
            _doc = doc;
            _images = images;
        }

        public void Run()
        {
            XElement document = _package.Xml("word/document.xml");
            if (document == null) throw new InvalidOperationException("Word 문서 내용을 찾을 수 없습니다.");
            _rels = _package.Relations("word/document.xml");
            ReadStyles();
            ReadNumbering();
            WalkBlocks(X.Kid(document, "body"));
            WriteNotes();
        }

        private void ReadStyles()
        {
            XElement styles = _package.Xml("word/styles.xml");
            foreach (XElement style in X.Kids(styles, "style"))
            {
                string id = X.Attr(style, "styleId");
                if (id == null) continue;
                var info = new Style { BasedOn = X.Attr(X.Kid(style, "basedOn"), "val") };
                string name = (X.Attr(X.Kid(style, "name"), "val") ?? "").ToLowerInvariant();
                Match heading = Regex.Match(name, @"^(heading|제목)\s*([1-9])$");
                if (heading.Success) info.Heading = int.Parse(heading.Groups[2].Value, CultureInfo.InvariantCulture);
                else if (name == "title" || name == "제목") info.Heading = 1;

                XElement pPr = X.Kid(style, "pPr");
                XElement outline = X.Kid(pPr, "outlineLvl");
                if (info.Heading == 0 && outline != null)
                {
                    int level = X.IntAttr(outline, "val", 9);
                    if (level < 9) info.Heading = level + 1;
                }
                XElement numPr = X.Kid(pPr, "numPr");
                if (numPr != null)
                {
                    info.NumId = X.Attr(X.Kid(numPr, "numId"), "val");
                    XElement ilvl = X.Kid(numPr, "ilvl");
                    if (ilvl != null) { info.Level = X.IntAttr(ilvl, "val", 0); info.HasLevel = true; }
                }
                XElement rPr = X.Kid(style, "rPr");
                info.Bold = X.Toggle(X.Kid(rPr, "b"));
                info.Italic = X.Toggle(X.Kid(rPr, "i"));
                _styles[id] = info;
            }
        }

        private void ReadNumbering()
        {
            XElement numbering = _package.Xml("word/numbering.xml");
            foreach (XElement abstractNum in X.Kids(numbering, "abstractNum"))
            {
                string id = X.Attr(abstractNum, "abstractNumId");
                if (id == null) continue;
                var levels = new Dictionary<int, XElement>();
                foreach (XElement level in X.Kids(abstractNum, "lvl")) levels[X.IntAttr(level, "ilvl", 0)] = level;
                _abstractLevels[id] = levels;
            }
            foreach (XElement num in X.Kids(numbering, "num"))
            {
                string id = X.Attr(num, "numId");
                string target = X.Attr(X.Kid(num, "abstractNumId"), "val");
                if (id != null && target != null) _numToAbstract[id] = target;
            }
        }

        private Style StyleChain(string id, Func<Style, bool> wanted)
        {
            for (int depth = 0; id != null && depth < 16; depth++)
            {
                Style style;
                if (!_styles.TryGetValue(id, out style)) return null;
                if (wanted(style)) return style;
                id = style.BasedOn;
            }
            return null;
        }

        private void WalkBlocks(XElement parent)
        {
            if (parent == null) return;
            foreach (XElement child in parent.Elements())
            {
                switch (X.Local(child))
                {
                    case "p": Paragraph(child); break;
                    case "tbl": Table(child); break;
                    case "sdt": WalkBlocks(X.Kid(child, "sdtContent")); break;
                    case "customXml":
                    case "ins":
                    case "moveTo": WalkBlocks(child); break;
                    case "AlternateContent": WalkBlocks(X.ChosenBranch(child)); break;
                }
            }
        }

        private void Paragraph(XElement p)
        {
            XElement pPr = X.Kid(p, "pPr");
            string styleId = X.Attr(X.Kid(pPr, "pStyle"), "val");

            var boxes = new List<XElement>();
            var inline = new MdInline();
            WalkInline(p, inline, new Format(), boxes);

            int heading = 0;
            XElement outline = X.Kid(pPr, "outlineLvl");
            if (outline != null && X.IntAttr(outline, "val", 9) < 9) heading = X.IntAttr(outline, "val", 9) + 1;
            else
            {
                Style headingStyle = StyleChain(styleId, s => s.Heading > 0);
                if (headingStyle != null) heading = headingStyle.Heading;
            }

            if (heading > 0)
            {
                _doc.AddHeading(heading, inline);
            }
            else
            {
                string numId;
                int level;
                ReadNumberingOf(pPr, styleId, out numId, out level);
                string format = numId == null ? null : LevelFormat(numId, level);
                if (format == null || format == "none") _doc.AddParagraph(inline);
                else if (format == "bullet") _doc.AddListItem(level, false, 0, inline);
                else _doc.AddListItem(level, true, NextNumber(numId, level), inline);
            }

            foreach (XElement box in boxes) WalkBlocks(box);
        }

        private void ReadNumberingOf(XElement pPr, string styleId, out string numId, out int level)
        {
            XElement numPr = X.Kid(pPr, "numPr");
            numId = X.Attr(X.Kid(numPr, "numId"), "val");
            level = X.IntAttr(X.Kid(numPr, "ilvl"), "val", 0);
            if (numPr == null)
            {
                Style style = StyleChain(styleId, s => s.NumId != null);
                if (style != null)
                {
                    numId = style.NumId;
                    level = style.Level;
                }
            }
            if (numId == "0") numId = null;
        }

        private string LevelFormat(string numId, int level)
        {
            string abstractId;
            Dictionary<int, XElement> levels;
            XElement definition;
            if (!_numToAbstract.TryGetValue(numId, out abstractId)) return null;
            if (!_abstractLevels.TryGetValue(abstractId, out levels)) return null;
            if (!levels.TryGetValue(level, out definition)) return null;
            return X.Attr(X.Kid(definition, "numFmt"), "val");
        }

        private int NextNumber(string numId, int level)
        {
            int[] counts;
            if (!_counters.TryGetValue(numId, out counts))
            {
                counts = new int[10];
                _counters[numId] = counts;
            }
            int index = Math.Max(0, Math.Min(9, level));
            for (int i = index + 1; i < counts.Length; i++) counts[i] = 0;
            if (counts[index] == 0)
            {
                Dictionary<int, XElement> levels;
                XElement definition;
                string abstractId;
                int start = 1;
                if (_numToAbstract.TryGetValue(numId, out abstractId) && _abstractLevels.TryGetValue(abstractId, out levels) && levels.TryGetValue(index, out definition))
                    start = X.IntAttr(X.Kid(definition, "start"), "val", 1);
                counts[index] = start;
            }
            else
            {
                counts[index]++;
            }
            return counts[index];
        }

        private void WalkInline(XElement parent, MdInline output, Format format, List<XElement> boxes)
        {
            foreach (XElement child in parent.Elements())
            {
                switch (X.Local(child))
                {
                    case "r": Run(child, output, boxes); break;
                    case "hyperlink": Hyperlink(child, output, format, boxes); break;
                    case "ins":
                    case "moveTo":
                    case "smartTag":
                    case "fldSimple":
                    case "customXml":
                        WalkInline(child, output, format, boxes);
                        break;
                    case "sdt":
                        XElement content = X.Kid(child, "sdtContent");
                        if (content != null) WalkInline(content, output, format, boxes);
                        break;
                    case "oMath":
                    case "oMathPara":
                        output.Add(string.Concat(X.Below(child, "t").Select(t => t.Value)));
                        break;
                    case "AlternateContent":
                        XElement chosen = X.ChosenBranch(child);
                        if (chosen != null) WalkInline(chosen, output, format, boxes);
                        break;
                }
            }
        }

        private void Hyperlink(XElement link, MdInline output, Format format, List<XElement> boxes)
        {
            DocPackage.Relation rel;
            string id = X.RelId(link);
            if (id == null || !_rels.TryGetValue(id, out rel) || !rel.External)
            {
                WalkInline(link, output, format, boxes);
                return;
            }
            var text = new MdInline();
            WalkInline(link, text, format, boxes);
            string label = text.Render(false);
            if (label.Length > 0) output.AddRaw("[" + label + "](" + rel.Target.Replace(" ", "%20").Replace(")", "%29") + ")");
        }

        private void Run(XElement run, MdInline output, List<XElement> boxes)
        {
            XElement rPr = X.Kid(run, "rPr");
            string styleId = X.Attr(X.Kid(rPr, "rStyle"), "val");
            var format = new Format
            {
                Bold = rPr != null && X.Kid(rPr, "b") != null ? X.Toggle(X.Kid(rPr, "b")) : StyleChain(styleId, s => s.Bold) != null,
                Italic = rPr != null && X.Kid(rPr, "i") != null ? X.Toggle(X.Kid(rPr, "i")) : StyleChain(styleId, s => s.Italic) != null,
                Strike = X.Toggle(X.Kid(rPr, "strike")) || X.Toggle(X.Kid(rPr, "dstrike"))
            };
            RunContent(run, output, format, boxes);
        }

        private void RunContent(XElement parent, MdInline output, Format format, List<XElement> boxes)
        {
            foreach (XElement child in parent.Elements())
            {
                switch (X.Local(child))
                {
                    case "t": output.Add(child.Value, format.Bold, format.Italic, format.Strike); break;
                    case "tab": output.Add(" ", format.Bold, format.Italic, format.Strike); break;
                    case "noBreakHyphen": output.Add("-", format.Bold, format.Italic, format.Strike); break;
                    case "br":
                        if (X.Attr(child, "type") != "page" && X.Attr(child, "type") != "column") output.AddBreak();
                        break;
                    case "cr": output.AddBreak(); break;
                    case "footnoteReference":
                    case "endnoteReference":
                        output.AddRaw(NoteReference(child));
                        break;
                    case "drawing":
                    case "pict":
                    case "object":
                        Picture(child, output);
                        boxes.AddRange(X.Below(child, "txbxContent"));
                        break;
                    case "AlternateContent":
                        XElement chosen = X.ChosenBranch(child);
                        if (chosen != null) RunContent(chosen, output, format, boxes);
                        break;
                }
            }
        }

        private void Picture(XElement holder, MdInline output)
        {
            string alt = X.Below(holder, "docPr").Select(d => X.Attr(d, "descr")).FirstOrDefault();
            IEnumerable<string> ids = X.Below(holder, "blip").Select(b => X.Attr(b, "embed"))
                .Concat(X.Below(holder, "imagedata").Select(X.RelId));
            foreach (string id in ids.Where(i => i != null))
            {
                DocPackage.Relation rel;
                if (!_rels.TryGetValue(id, out rel)) continue;
                if (rel.External)
                {
                    output.AddRaw(MdImages.Markdown(alt, rel.Target.Replace(" ", "%20")));
                    continue;
                }
                byte[] data = _package.Bytes(rel.Target);
                string link = data == null ? null : _images.SaveEmbedded(data, rel.Target);
                if (link != null) output.AddRaw(MdImages.Markdown(alt, link));
            }
        }

        private string NoteReference(XElement reference)
        {
            string key = X.Local(reference) + ":" + X.Attr(reference, "id");
            int number;
            if (!_noteNumbers.TryGetValue(key, out number))
            {
                number = _noteNumbers.Count + 1;
                _noteNumbers[key] = number;
            }
            return "[^" + number + "]";
        }

        private void WriteNotes()
        {
            if (_noteNumbers.Count == 0) return;
            XElement footnotes = _package.Xml("word/footnotes.xml");
            XElement endnotes = _package.Xml("word/endnotes.xml");
            foreach (KeyValuePair<string, int> note in _noteNumbers.OrderBy(n => n.Value))
            {
                string[] key = note.Key.Split(':');
                XElement source = key[0] == "footnoteReference" ? footnotes : endnotes;
                XElement body = X.Kids(source, key[0] == "footnoteReference" ? "footnote" : "endnote").FirstOrDefault(n => X.Attr(n, "id") == key[1]);
                if (body == null) continue;
                _doc.AddFootnote(note.Value, CellText(body));
            }
        }

        private void Table(XElement table)
        {
            var rows = new List<List<string>>();
            var plainCells = new List<XElement>();
            foreach (XElement row in X.Kids(table, "tr"))
            {
                var cells = new List<string>();
                foreach (XElement cell in RowCells(row))
                {
                    XElement cellProps = X.Kid(cell, "tcPr");
                    XElement merge = X.Kid(cellProps, "vMerge");
                    bool continued = merge != null && X.Attr(merge, "val") != "restart";
                    cells.Add(continued ? "" : CellText(cell));
                    plainCells.Add(cell);
                    int span = X.IntAttr(X.Kid(cellProps, "gridSpan"), "val", 1);
                    for (int i = 1; i < span; i++) cells.Add("");
                }
                rows.Add(cells);
            }

            // A table with a single column is a box around text, so it reads better as plain text.
            if (rows.Count > 0 && rows.All(r => r.Count <= 1))
            {
                foreach (XElement cell in plainCells) WalkBlocks(cell);
                return;
            }
            _doc.AddTable(rows);
        }

        private static IEnumerable<XElement> RowCells(XElement row)
        {
            foreach (XElement child in row.Elements())
            {
                if (X.Local(child) == "tc") yield return child;
                else if (X.Local(child) == "sdt")
                {
                    foreach (XElement cell in X.Kids(X.Kid(child, "sdtContent"), "tc")) yield return cell;
                }
            }
        }

        // One table cell has to fit on one Markdown line, so lists and nested tables are flattened.
        private string CellText(XElement cell)
        {
            var lines = new List<string>();
            CollectCellLines(cell, lines);
            return string.Join("  \n", lines);
        }

        private void CollectCellLines(XElement parent, List<string> lines)
        {
            foreach (XElement child in parent.Elements())
            {
                switch (X.Local(child))
                {
                    case "p":
                        var inline = new MdInline();
                        var boxes = new List<XElement>();
                        WalkInline(child, inline, new Format(), boxes);
                        string text = inline.Render(false);
                        if (text.Length > 0)
                        {
                            string numId;
                            int level;
                            ReadNumberingOf(X.Kid(child, "pPr"), X.Attr(X.Kid(X.Kid(child, "pPr"), "pStyle"), "val"), out numId, out level);
                            string format = numId == null ? null : LevelFormat(numId, level);
                            if (format == "bullet") text = "• " + text;
                            else if (format != null && format != "none") text = NextNumber(numId, level) + ". " + text;
                            lines.Add(text);
                        }
                        foreach (XElement box in boxes) CollectCellLines(box, lines);
                        break;
                    case "tbl":
                        foreach (XElement row in X.Kids(child, "tr"))
                        {
                            var parts = new List<string>();
                            foreach (XElement inner in RowCells(row)) parts.Add(CellText(inner).Replace("  \n", " "));
                            string joined = string.Join(" / ", parts.Where(s => s.Length > 0));
                            if (joined.Length > 0) lines.Add(joined);
                        }
                        break;
                    case "sdt":
                        XElement content = X.Kid(child, "sdtContent");
                        if (content != null) CollectCellLines(content, lines);
                        break;
                }
            }
        }
    }

    internal sealed class PptxReader
    {
        private readonly DocPackage _package;
        private readonly MdDoc _doc;
        private readonly MdImages _images;
        private Dictionary<string, DocPackage.Relation> _rels;

        public PptxReader(DocPackage package, MdDoc doc, MdImages images)
        {
            _package = package;
            _doc = doc;
            _images = images;
        }

        public void Run()
        {
            XElement presentation = _package.Xml("ppt/presentation.xml");
            if (presentation == null) throw new InvalidOperationException("PowerPoint 문서 내용을 찾을 수 없습니다.");
            Dictionary<string, DocPackage.Relation> rels = _package.Relations("ppt/presentation.xml");

            int number = 0;
            foreach (XElement id in X.Below(X.Kid(presentation, "sldIdLst"), "sldId"))
            {
                DocPackage.Relation rel;
                if (!rels.TryGetValue(X.RelId(id) ?? "", out rel)) continue;
                number++;
                XElement slide = _package.Xml(rel.Target);
                if (slide != null) Slide(slide, rel.Target, number);
            }
        }

        private void Slide(XElement slide, string part, int number)
        {
            _rels = _package.Relations(part);
            XElement tree = X.Below(slide, "spTree").FirstOrDefault();

            var shapes = new List<XElement>();
            Flatten(tree, shapes);

            XElement title = shapes.FirstOrDefault(s => X.Local(s) == "sp" && IsTitle(s));
            var heading = new MdInline();
            if (title != null) heading.Add(string.Join(" ", ParagraphTexts(title)));
            if (heading.PlainText.Trim().Length == 0) heading.Add("슬라이드 " + number);
            _doc.AddHeading(2, heading);

            foreach (XElement shape in shapes)
            {
                if (shape == title) continue;
                switch (X.Local(shape))
                {
                    case "sp": TextShape(shape); break;
                    case "pic": Picture(shape); break;
                    case "graphicFrame":
                        XElement table = X.Below(shape, "tbl").FirstOrDefault();
                        if (table != null) Table(table);
                        break;
                }
            }
            Notes();
        }

        private static void Flatten(XElement parent, List<XElement> shapes)
        {
            foreach (XElement child in parent.Elements())
            {
                switch (X.Local(child))
                {
                    case "sp":
                    case "pic":
                    case "graphicFrame":
                        shapes.Add(child);
                        break;
                    case "grpSp":
                        Flatten(child, shapes);
                        break;
                    case "AlternateContent":
                        XElement chosen = X.ChosenBranch(child);
                        if (chosen != null) Flatten(chosen, shapes);
                        break;
                }
            }
        }

        private static XElement Placeholder(XElement shape)
        {
            return X.Below(X.Kid(shape, "nvSpPr"), "ph").FirstOrDefault();
        }

        private static bool IsTitle(XElement shape)
        {
            string type = X.Attr(Placeholder(shape), "type");
            return type == "title" || type == "ctrTitle";
        }

        private static IEnumerable<string> ParagraphTexts(XElement shape)
        {
            return X.Below(X.Kid(shape, "txBody"), "p")
                .Select(p => string.Concat(X.Below(p, "t").Select(t => t.Value)).Trim())
                .Where(t => t.Length > 0);
        }

        private void TextShape(XElement shape)
        {
            XElement placeholder = Placeholder(shape);
            string type = X.Attr(placeholder, "type");
            if (type == "sldNum" || type == "dt" || type == "ftr" || type == "hdr") return;
            bool bodyBullets = placeholder != null && type != "subTitle";

            int counter = 0;
            foreach (XElement p in X.Kids(X.Kid(shape, "txBody"), "p"))
            {
                MdInline text = Inline(p);
                XElement props = X.Kid(p, "pPr");
                int level = X.IntAttr(props, "lvl", 0);
                bool numbered = X.Kid(props, "buAutoNum") != null;
                bool bullet = X.Kid(props, "buChar") != null || (bodyBullets && X.Kid(props, "buNone") == null);
                if (text.IsEmpty) continue;
                if (numbered) _doc.AddListItem(level, true, ++counter, text);
                else if (bullet) _doc.AddListItem(level, false, 0, text);
                else _doc.AddParagraph(text);
                if (!numbered) counter = 0;
            }
        }

        private MdInline Inline(XElement paragraph)
        {
            var output = new MdInline();
            foreach (XElement child in paragraph.Elements())
            {
                string kind = X.Local(child);
                if (kind == "br")
                {
                    output.AddBreak();
                }
                else if (kind == "r" || kind == "fld")
                {
                    XElement props = X.Kid(child, "rPr");
                    string text = string.Concat(X.Kids(child, "t").Select(t => t.Value));
                    DocPackage.Relation rel;
                    string link = X.RelId(X.Kid(props, "hlinkClick"));
                    if (link != null && _rels.TryGetValue(link, out rel) && rel.External && text.Trim().Length > 0)
                    {
                        output.AddRaw("[" + MdInline.Escape(text.Trim()) + "](" + rel.Target.Replace(" ", "%20").Replace(")", "%29") + ")");
                        continue;
                    }
                    output.Add(text, X.Attr(props, "b") == "1", X.Attr(props, "i") == "1", (X.Attr(props, "strike") ?? "noStrike") != "noStrike");
                }
            }
            return output;
        }

        private void Picture(XElement shape)
        {
            string id = X.Below(shape, "blip").Select(b => X.Attr(b, "embed")).FirstOrDefault();
            DocPackage.Relation rel;
            if (id == null || !_rels.TryGetValue(id, out rel) || rel.External) return;
            byte[] data = _package.Bytes(rel.Target);
            string link = data == null ? null : _images.SaveEmbedded(data, rel.Target);
            string alt = X.Below(X.Kid(shape, "nvPicPr"), "cNvPr").Select(c => X.Attr(c, "descr")).FirstOrDefault();
            if (link != null) _doc.AddRaw(MdImages.Markdown(alt, link));
        }

        private void Table(XElement table)
        {
            var rows = new List<List<string>>();
            foreach (XElement row in X.Kids(table, "tr"))
            {
                var cells = new List<string>();
                foreach (XElement cell in X.Kids(row, "tc"))
                {
                    bool merged = X.Attr(cell, "hMerge") == "1" || X.Attr(cell, "vMerge") == "1";
                    var parts = merged ? new List<string>() :
                        X.Below(X.Kid(cell, "txBody"), "p").Select(p => Inline(p).Render(false)).Where(t => t.Length > 0).ToList();
                    cells.Add(string.Join("  \n", parts));
                }
                rows.Add(cells);
            }
            _doc.AddTable(rows);
        }

        private void Notes()
        {
            DocPackage.Relation rel = _rels.Values.FirstOrDefault(r => r.Type.EndsWith("/notesSlide"));
            if (rel == null) return;
            XElement notes = _package.Xml(rel.Target);
            XElement body = X.Below(notes, "sp").FirstOrDefault(s => X.Attr(Placeholder(s), "type") == "body");
            if (body == null) return;
            var lines = ParagraphTexts(body).Select(MdInline.Escape).ToList();
            if (lines.Count > 0) _doc.AddQuote("발표자 노트: " + string.Join("  \n", lines));
        }
    }

    internal sealed class XlsxReader
    {
        private readonly DocPackage _package;
        private readonly MdDoc _doc;
        private readonly List<string> _strings = new List<string>();
        private readonly List<int> _cellFormats = new List<int>();
        private readonly Dictionary<int, string> _customFormats = new Dictionary<int, string>();

        public XlsxReader(DocPackage package, MdDoc doc)
        {
            _package = package;
            _doc = doc;
        }

        public void Run()
        {
            XElement workbook = _package.Xml("xl/workbook.xml");
            if (workbook == null) throw new InvalidOperationException("Excel 문서 내용을 찾을 수 없습니다.");
            Dictionary<string, DocPackage.Relation> rels = _package.Relations("xl/workbook.xml");
            ReadStrings();
            ReadFormats();

            foreach (XElement sheet in X.Kids(X.Kid(workbook, "sheets"), "sheet"))
            {
                if ((X.Attr(sheet, "state") ?? "visible") != "visible") continue;
                DocPackage.Relation rel;
                if (!rels.TryGetValue(X.RelId(sheet) ?? "", out rel)) continue;
                XElement data = X.Kid(_package.Xml(rel.Target), "sheetData");
                List<List<string>> rows = ReadRows(data);
                if (rows.Count == 0) continue;
                var title = new MdInline();
                title.Add(X.Attr(sheet, "name") ?? "");
                _doc.AddHeading(2, title);
                _doc.AddTable(rows);
            }
        }

        private void ReadStrings()
        {
            XElement table = _package.Xml("xl/sharedStrings.xml");
            foreach (XElement item in X.Kids(table, "si")) _strings.Add(ItemText(item));
        }

        // Skips the pronunciation guides, which sit next to the text as <rPh>.
        private static string ItemText(XElement item)
        {
            string direct = string.Concat(X.Kids(item, "t").Select(t => t.Value));
            string runs = string.Concat(X.Kids(item, "r").SelectMany(r => X.Kids(r, "t")).Select(t => t.Value));
            return direct + runs;
        }

        private void ReadFormats()
        {
            XElement styles = _package.Xml("xl/styles.xml");
            foreach (XElement format in X.Kids(X.Kid(styles, "numFmts"), "numFmt"))
                _customFormats[X.IntAttr(format, "numFmtId", -1)] = X.Attr(format, "formatCode") ?? "";
            foreach (XElement xf in X.Kids(X.Kid(styles, "cellXfs"), "xf"))
                _cellFormats.Add(X.IntAttr(xf, "numFmtId", 0));
        }

        private List<List<string>> ReadRows(XElement data)
        {
            var rows = new List<List<string>>();
            foreach (XElement row in X.Kids(data, "row"))
            {
                var cells = new List<string>();
                foreach (XElement cell in X.Kids(row, "c"))
                {
                    int column = ColumnOf(X.Attr(cell, "r"), cells.Count);
                    while (cells.Count < column) cells.Add("");
                    cells.Add(MdInline.Escape(CellValue(cell)).Replace("\r", "").Replace("\n", "  \n"));
                }
                if (cells.Any(c => c.Trim().Length > 0)) rows.Add(cells);
            }
            return rows;
        }

        private static int ColumnOf(string reference, int fallback)
        {
            if (string.IsNullOrEmpty(reference)) return fallback;
            int column = 0;
            foreach (char c in reference)
            {
                if (c < 'A' || c > 'Z') break;
                column = column * 26 + (c - 'A' + 1);
            }
            return column > 0 ? column - 1 : fallback;
        }

        private string CellValue(XElement cell)
        {
            string type = X.Attr(cell, "t");
            string raw = (X.Kid(cell, "v") ?? new XElement("v")).Value;
            switch (type)
            {
                case "s":
                    int index;
                    return int.TryParse(raw, out index) && index >= 0 && index < _strings.Count ? _strings[index] : "";
                case "inlineStr": return ItemText(X.Kid(cell, "is") ?? new XElement("is"));
                case "str":
                case "e": return raw;
                case "b": return raw == "1" ? "TRUE" : "FALSE";
            }

            double number;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) return raw;
            int style = X.IntAttr(cell, "s", 0);
            int formatId = style >= 0 && style < _cellFormats.Count ? _cellFormats[style] : 0;
            return FormatNumber(number, formatId);
        }

        private string FormatNumber(double number, int formatId)
        {
            string code;
            _customFormats.TryGetValue(formatId, out code);
            string stripped = code == null ? "" : Regex.Replace(code, "\"[^\"]*\"|\\[[^\\]]*\\]|\\\\.", "");

            bool percent = formatId == 9 || formatId == 10 || stripped.Contains("%");
            if (percent) return (number * 100).ToString("0.##", CultureInfo.InvariantCulture) + "%";

            bool date = (formatId >= 14 && formatId <= 22) || (formatId >= 45 && formatId <= 47) || Regex.IsMatch(stripped, "[ymdhsYMDHS]");
            if (date && number >= 0 && number < 2958466)
            {
                DateTime value = DateTime.FromOADate(number);
                if (number < 1) return value.ToString("HH:mm", CultureInfo.InvariantCulture);
                return value.ToString(number == Math.Floor(number) ? "yyyy-MM-dd" : "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
            }
            return number.ToString("G15", CultureInfo.InvariantCulture);
        }
    }

    internal sealed class HwpxReader
    {
        private sealed class CharFormat
        {
            public bool Bold, Italic, Strike;
        }

        private readonly DocPackage _package;
        private readonly MdDoc _doc;
        private readonly MdImages _images;
        private readonly Dictionary<string, CharFormat> _charFormats = new Dictionary<string, CharFormat>();
        private readonly Dictionary<string, int> _styleHeadings = new Dictionary<string, int>();
        private readonly Dictionary<string, string> _paraHeadingTypes = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _binaries = new Dictionary<string, string>();
        private readonly Dictionary<string, int> _counters = new Dictionary<string, int>();
        private readonly List<XElement> _notes = new List<XElement>();

        public HwpxReader(DocPackage package, MdDoc doc, MdImages images)
        {
            _package = package;
            _doc = doc;
            _images = images;
        }

        public void Run()
        {
            ReadHeader();
            List<string> sections = ReadManifest();
            if (sections.Count == 0) throw new InvalidOperationException("한글 문서 내용을 찾을 수 없습니다.");
            foreach (string section in sections)
            {
                XElement root = _package.Xml(section);
                if (root != null) WalkBlocks(root);
            }
            for (int i = 0; i < _notes.Count; i++) _doc.AddFootnote(i + 1, CellText(_notes[i]));
        }

        private void ReadHeader()
        {
            XElement header = _package.Xml("Contents/header.xml");
            foreach (XElement format in X.Below(header, "charPr"))
            {
                XElement strike = X.Kid(format, "strikeout");
                _charFormats[X.Attr(format, "id") ?? ""] = new CharFormat
                {
                    Bold = X.Kid(format, "bold") != null,
                    Italic = X.Kid(format, "italic") != null,
                    Strike = strike != null && (X.Attr(strike, "shape") ?? "NONE") != "NONE" && (X.Attr(strike, "shape") ?? "") != "3D"
                };
            }
            foreach (XElement style in X.Below(header, "style"))
            {
                int outline = HangulFile.OutlineLevel(X.Attr(style, "name"), X.Attr(style, "engName"));
                if (outline > 0) _styleHeadings[X.Attr(style, "id") ?? ""] = outline;
            }
            foreach (XElement para in X.Below(header, "paraPr"))
            {
                XElement heading = X.Below(para, "heading").FirstOrDefault();
                if (heading != null) _paraHeadingTypes[X.Attr(para, "id") ?? ""] = X.Attr(heading, "type") ?? "NONE";
            }
        }

        private List<string> ReadManifest()
        {
            var sections = new List<string>();
            XElement manifest = _package.Xml("Contents/content.hpf");
            var items = new Dictionary<string, string>();
            foreach (XElement item in X.Kids(X.Kid(manifest, "manifest"), "item"))
            {
                string id = X.Attr(item, "id");
                string href = X.Attr(item, "href");
                if (id == null || href == null) continue;
                items[id] = href;
                if (href.StartsWith("BinData/", StringComparison.OrdinalIgnoreCase)) _binaries[id] = href;
            }
            foreach (XElement reference in X.Kids(X.Kid(manifest, "spine"), "itemref"))
            {
                string href;
                if (items.TryGetValue(X.Attr(reference, "idref") ?? "", out href) && Regex.IsMatch(href, @"section\d+\.xml$", RegexOptions.IgnoreCase))
                    sections.Add(href);
            }
            if (sections.Count == 0)
            {
                sections = _package.Names.Where(n => Regex.IsMatch(n, @"^Contents/section\d+\.xml$", RegexOptions.IgnoreCase))
                    .OrderBy(n => int.Parse(Regex.Match(n, @"\d+").Value, CultureInfo.InvariantCulture)).ToList();
            }
            return sections;
        }

        private void WalkBlocks(XElement parent)
        {
            foreach (XElement child in parent.Elements())
            {
                string kind = X.Local(child);
                if (kind == "p") Paragraph(child);
                else if (kind == "subList" || kind == "sec") WalkBlocks(child);
            }
        }

        private void Paragraph(XElement p)
        {
            var inline = new MdInline();
            var after = new List<XElement>();
            foreach (XElement run in X.Kids(p, "run")) Run(run, inline, after);

            int heading;
            string headingType;
            _paraHeadingTypes.TryGetValue(X.Attr(p, "paraPrIDRef") ?? "", out headingType);
            if (_styleHeadings.TryGetValue(X.Attr(p, "styleIDRef") ?? "", out heading))
            {
                _doc.AddHeading(heading, inline);
            }
            else if (headingType == "BULLET")
            {
                _doc.AddListItem(0, false, 0, inline);
            }
            else if (headingType == "NUMBER" && !inline.IsEmpty)
            {
                string key = X.Attr(p, "paraPrIDRef") ?? "";
                int count;
                _counters.TryGetValue(key, out count);
                _counters[key] = count + 1;
                _doc.AddListItem(0, true, count + 1, inline);
            }
            else
            {
                _doc.AddParagraph(inline);
            }

            foreach (XElement extra in after)
            {
                if (X.Local(extra) == "tbl") Table(extra);
                else WalkBlocks(extra);
            }
        }

        private void Run(XElement run, MdInline output, List<XElement> after)
        {
            CharFormat format;
            if (!_charFormats.TryGetValue(X.Attr(run, "charPrIDRef") ?? "", out format)) format = new CharFormat();
            RunContent(run, output, format, after);
        }

        private void RunContent(XElement parent, MdInline output, CharFormat format, List<XElement> after)
        {
            foreach (XElement child in parent.Elements())
            {
                switch (X.Local(child))
                {
                    case "t": Text(child, output, format); break;
                    case "tbl": after.Add(child); break;
                    case "pic": Picture(child, output); break;
                    case "equation":
                        XElement script = X.Kid(child, "script");
                        if (script != null) output.Add(script.Value);
                        break;
                    case "ctrl":
                        foreach (XElement control in child.Elements())
                        {
                            if (X.Local(control) != "footNote" && X.Local(control) != "endNote") continue;
                            _notes.Add(X.Kid(control, "subList") ?? control);
                            output.AddRaw("[^" + _notes.Count + "]");
                        }
                        break;
                    case "switch":
                        XElement chosen = X.Kid(child, "case") ?? X.Kid(child, "default");
                        if (chosen != null) RunContent(chosen, output, format, after);
                        break;
                    case "rect":
                    case "ellipse":
                    case "polygon":
                    case "curve":
                    case "container":
                        foreach (XElement list in X.Below(child, "subList")) after.Add(list);
                        foreach (XElement picture in X.Below(child, "pic")) Picture(picture, output);
                        break;
                }
            }
        }

        private static void Text(XElement t, MdInline output, CharFormat format)
        {
            foreach (XNode node in t.Nodes())
            {
                var text = node as XText;
                if (text != null)
                {
                    output.Add(text.Value, format.Bold, format.Italic, format.Strike);
                    continue;
                }
                var element = node as XElement;
                if (element == null) continue;
                switch (X.Local(element))
                {
                    case "tab":
                    case "nbSpace":
                    case "fwSpace":
                        output.Add(" ", format.Bold, format.Italic, format.Strike);
                        break;
                    case "lineBreak": output.AddBreak(); break;
                    case "hyphen": output.Add("-", format.Bold, format.Italic, format.Strike); break;
                }
            }
        }

        private void Picture(XElement picture, MdInline output)
        {
            string id = X.Below(picture, "img").Select(i => X.Attr(i, "binaryItemIDRef")).FirstOrDefault();
            string href;
            if (id == null || !_binaries.TryGetValue(id, out href)) return;
            byte[] data = _package.Bytes(href);
            string link = data == null ? null : _images.SaveEmbedded(data, href);
            if (link != null) output.AddRaw(MdImages.Markdown(null, link));
        }

        private void Table(XElement table)
        {
            var rows = new List<List<string>>();
            var cells = new List<XElement>();
            foreach (XElement row in X.Kids(table, "tr"))
            {
                foreach (XElement cell in X.Kids(row, "tc"))
                {
                    XElement address = X.Kid(cell, "cellAddr");
                    int rowIndex = X.IntAttr(address, "rowAddr", rows.Count);
                    int column = X.IntAttr(address, "colAddr", 0);
                    while (rows.Count <= rowIndex) rows.Add(new List<string>());
                    while (rows[rowIndex].Count <= column) rows[rowIndex].Add("");
                    rows[rowIndex][column] = CellText(X.Kid(cell, "subList") ?? cell);
                    cells.Add(cell);
                }
            }

            // A one-cell table is how Hangul usually draws a box around text.
            if (rows.All(r => r.Count <= 1))
            {
                foreach (XElement cell in cells) WalkBlocks(X.Kid(cell, "subList") ?? cell);
                return;
            }
            _doc.AddTable(rows);
        }

        private string CellText(XElement container)
        {
            var lines = new List<string>();
            foreach (XElement p in X.Kids(container, "p"))
            {
                var inline = new MdInline();
                var after = new List<XElement>();
                foreach (XElement run in X.Kids(p, "run")) Run(run, inline, after);
                string text = inline.Render(false);
                if (text.Length > 0) lines.Add(text);
                foreach (XElement extra in after)
                {
                    if (X.Local(extra) == "tbl")
                    {
                        foreach (XElement cell in X.Below(extra, "tc"))
                        {
                            string inner = CellText(X.Kid(cell, "subList") ?? cell).Replace("  \n", " ");
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
}
