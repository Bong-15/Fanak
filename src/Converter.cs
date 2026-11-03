using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Windows.Data.Pdf;
using Windows.Foundation;
using Windows.Storage;
using Windows.Storage.Streams;

namespace Fanak
{
    internal enum FileKind { Unknown, Image, Pdf, Document }

    internal sealed class MenuPlan
    {
        public List<string> Formats = new List<string>();
        public string Message;
    }

    internal sealed class ConvertRequest
    {
        public string[] Files;
        public string Format;
        public int? Dpi;
        public int? PdfDpi;
        public bool Resample;
        public Int32Rect? Crop;
        public CancellationToken Cancel;
        public Action<string, int, int> FileStarted;
    }

    // Everything here uses components that ship with Windows: WIC for images, Windows.Data.Pdf for reading PDFs.
    internal static class Converter
    {
        public const string CropAction = "CROP";
        public const string DpiAction = "DPI";
        public const string KeepFormat = "ORIGINAL";
        public const int MinDpi = 1;
        public const int MaxDpi = 2400;
        public const int DefaultPdfDpi = 300;

        private static readonly HashSet<string> ImageExts = new HashSet<string>(
            new[] { "png", "jpg", "jpeg", "jfif", "bmp", "tif", "tiff", "gif", "webp", "heic", "heif", "ico" },
            StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> DocumentExts = new HashSet<string>(
            new[] { "docx", "docm", "pptx", "pptm", "xlsx", "xlsm" }.Concat(HangulFile.Exts).Concat(TextDocs.DelimitedExts).Concat(TextDocs.WordOpensExts),
            StringComparer.OrdinalIgnoreCase);

        // These are binary formats that only their own programs can read. Saving as the newer format fixes that.
        private static readonly Dictionary<string, string> OldDocumentExts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "doc", "docx" }, { "ppt", "pptx" }, { "xls", "xlsx" }
        };

        // The menu draws these in order, so formats of a kind stay side by side.
        private static readonly string[] ImageTargets = { "PNG", "JPG", "TIFF", "BMP", "GIF", "PDF", CropAction, DpiAction };
        private static readonly string[] PdfTargets = { "PNG", "JPG", "TIFF", "MD", DpiAction };
        private static readonly string[] DelimitedTargets = { "XLSX", "MD" };

        private static string Ext(string path)
        {
            return Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        }

        public static FileKind KindOf(string path)
        {
            string ext = Ext(path);
            if (ext == "pdf") return FileKind.Pdf;
            if (DocumentExts.Contains(ext)) return FileKind.Document;
            // Their own programs can still open the old formats, though the MD readers can't.
            if (OldDocumentExts.ContainsKey(ext) && (OfficePdf.CanConvert(path) || OfficePdf.CanMakeHwp(path))) return FileKind.Document;
            return ImageExts.Contains(ext) ? FileKind.Image : FileKind.Unknown;
        }

        private static string[] TargetsOf(string path)
        {
            switch (KindOf(path))
            {
                case FileKind.Image: return ImageTargets;
                case FileKind.Pdf: return PdfTargets;
            }
            if (TextDocs.DelimitedExts.Contains(Ext(path))) return DelimitedTargets;

            var targets = new List<string>();
            if (!OldDocumentExts.ContainsKey(Ext(path))) targets.Add("MD");
            if (OfficePdf.CanConvert(path)) targets.Add("PDF");
            if (OfficePdf.CanMakeDocx(path)) targets.Add("DOCX");
            if (OfficePdf.CanMakeHwp(path)) targets.Add("HWP");
            return targets.ToArray();
        }

        public static string Label(string action)
        {
            if (action == CropAction) return "자르기";
            return action;
        }

        // Formats offered are the ones every dropped file can be converted to.
        public static MenuPlan Plan(IList<string> files)
        {
            var plan = new MenuPlan();
            var kinds = files.Select(KindOf).Distinct().ToList();
            if (kinds.Contains(FileKind.Unknown))
            {
                string newer = null;
                foreach (string file in files)
                {
                    if (OldDocumentExts.TryGetValue(Ext(file), out newer)) break;
                }
                plan.Message = newer != null ? "." + newer + "로 저장한 뒤\n변환해 주세요" : "지원하지 않는\n파일 형식";
                return plan;
            }

            IEnumerable<string> formats = null;
            foreach (string file in files)
            {
                string[] targets = TargetsOf(file);
                formats = formats == null ? targets : formats.Intersect(targets);
            }
            plan.Formats = formats.ToList();
            // Images merge into one PDF and documents each make their own, so a mix has no clear meaning.
            if (kinds.Count > 1) plan.Formats.Remove("PDF");
            if (plan.Formats.Count == 0)
            {
                plan.Message = "함께 변환할 수\n없는 파일 조합";
                return plan;
            }

            if (files.Count == 1)
            {
                string own = NormalizeFormat(Ext(files[0]));
                plan.Formats.RemoveAll(f => f == own);
            }
            else
            {
                plan.Formats.Remove(CropAction);
            }
            return plan;
        }

        private static string NormalizeFormat(string ext)
        {
            switch (ext)
            {
                case "png": return "PNG";
                case "jpg":
                case "jpeg":
                case "jfif": return "JPG";
                case "bmp": return "BMP";
                case "tif":
                case "tiff": return "TIFF";
                case "gif": return "GIF";
                default: return null;
            }
        }

        // Runs one request and returns the files it created. Failures are per file, except that a
        // merged PDF fails as a whole rather than silently leaving a page out.
        public static List<string> Run(ConvertRequest request, Action<string, Exception> onFailure)
        {
            var outputs = new List<string>();
            string format = request.Format.ToUpperInvariant();
            if (request.Dpi.HasValue && (request.Dpi < MinDpi || request.Dpi > MaxDpi))
                throw new ArgumentOutOfRangeException("Dpi", "DPI는 " + MinDpi + "~" + MaxDpi + " 사이여야 합니다.");

            if (format == "PDF" && request.Files.All(f => KindOf(f) == FileKind.Image))
            {
                try
                {
                    if (request.FileStarted != null) request.FileStarted(request.Files[0], 1, 1);
                    outputs.Add(ImagesToPdf(request.Files));
                }
                catch (Exception ex)
                {
                    onFailure(request.Files[0], ex);
                }
                return outputs;
            }

            for (int index = 0; index < request.Files.Length && !request.Cancel.IsCancellationRequested; index++)
            {
                string file = request.Files[index];
                try
                {
                    if (request.FileStarted != null) request.FileStarted(file, index + 1, request.Files.Length);
                    if (!File.Exists(file)) throw new FileNotFoundException("파일을 찾을 수 없습니다.", file);
                    FileKind kind = KindOf(file);
                    if (kind == FileKind.Unknown) throw new NotSupportedException("지원하지 않는 파일 형식입니다.");
                    if (format == "MD")
                    {
                        if (kind == FileKind.Image) throw new NotSupportedException("이미지는 MD로 바꿀 수 없습니다.");
                        outputs.Add(MarkdownConverter.Convert(file));
                    }
                    else if (format == "PDF")
                    {
                        if (kind != FileKind.Document || !OfficePdf.CanConvert(file))
                            throw new NotSupportedException(kind == FileKind.Document ? "이 파일은 PDF로 바꿀 수 없습니다. 파일을 열 프로그램이 설치되어 있어야 합니다." : "이미지와 함께 PDF로 묶을 수 없는 파일입니다.");
                        outputs.Add(OfficePdf.Convert(file));
                    }
                    else if (format == "XLSX" && TextDocs.DelimitedExts.Contains(Ext(file))) outputs.Add(TextDocs.ToXlsx(file));
                    else if (format == "DOCX" && OfficePdf.CanMakeDocx(file)) outputs.Add(OfficePdf.ToDocx(file));
                    else if (format == "HWP" && OfficePdf.CanMakeHwp(file)) outputs.Add(OfficePdf.ToHwp(file));
                    else if (kind == FileKind.Document) throw new NotSupportedException("이 문서는 " + format + " 형식으로 바꿀 수 없습니다.");
                    else if (kind == FileKind.Pdf) outputs.AddRange(PdfToImages(file, format, request.Dpi ?? request.PdfDpi ?? DefaultPdfDpi, request.Cancel));
                    else outputs.Add(ConvertImage(file, format, request.Dpi, request.Resample, request.Crop));
                }
                catch (Exception ex)
                {
                    // Stopping a program on request makes it fail, which isn't a failure to report.
                    if (request.Cancel.IsCancellationRequested) Log.Write("변환 취소: " + file);
                    else onFailure(file, ex);
                }
            }
            return outputs;
        }

        internal static string UniquePath(string source, string suffix, string ext)
        {
            string dir = Path.GetDirectoryName(source);
            string stem = Path.GetFileNameWithoutExtension(source) + suffix;
            string candidate = Path.Combine(dir, stem + "." + ext);
            for (int n = 1; File.Exists(candidate); n++)
                candidate = Path.Combine(dir, stem + " (" + n + ")." + ext);
            return candidate;
        }

        internal static void TryDelete(string path)
        {
            try { File.Delete(path); }
            catch (IOException ex) { Log.Write("임시 결과 삭제 실패: " + path + " - " + ex.Message); }
            catch (UnauthorizedAccessException ex) { Log.Write("임시 결과 삭제 실패: " + path + " - " + ex.Message); }
        }

        // ---- Images ----

        private static string ConvertImage(string source, string format, int? dpi, bool resample, Int32Rect? crop)
        {
            // Formats Windows can read but not write, and GIF for DPI changes since GIF has no DPI field, become PNG.
            string target = format;
            if (target == KeepFormat)
            {
                target = NormalizeFormat(Ext(source));
                if (target == null || (target == "GIF" && dpi.HasValue)) target = "PNG";
            }
            if (target == "GIF" && dpi.HasValue) throw new NotSupportedException("GIF에는 DPI를 저장할 수 없습니다.");

            // Re-encoding a JPEG loses quality, and a DPI change does not need it.
            if (dpi.HasValue && !resample && !crop.HasValue && target == "JPG" && NormalizeFormat(Ext(source)) == "JPG")
            {
                byte[] patched = JpegDpi.Patch(File.ReadAllBytes(source), dpi.Value);
                if (patched != null)
                {
                    string copy = UniquePath(source, "_" + dpi.Value + "dpi", Ext(source));
                    File.WriteAllBytes(copy, patched);
                    VerifySavedDpi(copy, dpi.Value);
                    return copy;
                }
            }

            BitmapSource image = LoadUpright(source);
            string suffix = "";
            if (crop.HasValue)
            {
                image = new CroppedBitmap(image, ClampCrop(crop.Value, image));
                suffix += "_crop";
            }
            if (dpi.HasValue)
            {
                image = ChangeDpi(image, dpi.Value, resample);
                suffix += "_" + dpi.Value + "dpi";
            }

            string output = UniquePath(source, suffix, target.ToLowerInvariant());
            WriteImage(image, output, target);
            if (dpi.HasValue) VerifySavedDpi(output, dpi.Value);
            return output;
        }

        // Reads the finished file back, so a success message means the value really is in the file.
        private static void VerifySavedDpi(string path, int dpi)
        {
            int width, height;
            double saved;
            bool readable = TryReadImageInfo(path, out width, out height, out saved);
            if (readable && Math.Abs(saved - dpi) < 0.5) return;
            TryDelete(path);
            throw new InvalidOperationException(readable
                ? "저장된 파일의 DPI가 " + saved.ToString("0.#", CultureInfo.InvariantCulture) + "로 확인되어 결과를 지웠습니다. (요청 " + dpi + ")"
                : "저장된 파일을 다시 읽을 수 없어 결과를 지웠습니다.");
        }

        public static bool TryReadImageInfo(string path, out int width, out int height, out double dpi)
        {
            width = 0;
            height = 0;
            dpi = 0;
            try
            {
                using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    BitmapFrame frame = BitmapDecoder.Create(input, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None).Frames[0];
                    width = frame.PixelWidth;
                    height = frame.PixelHeight;
                    dpi = frame.DpiX > 0 ? frame.DpiX : 96;
                    return true;
                }
            }
            catch (Exception ex)
            {
                if (ex is NotSupportedException || ex is FileFormatException || ex is IOException || ex is ArgumentException || ex is UnauthorizedAccessException) return false;
                throw;
            }
        }

        private static Int32Rect ClampCrop(Int32Rect crop, BitmapSource image)
        {
            int x = Math.Max(0, crop.X);
            int y = Math.Max(0, crop.Y);
            int width = Math.Min(crop.Width, image.PixelWidth - x);
            int height = Math.Min(crop.Height, image.PixelHeight - y);
            if (width < 1 || height < 1) throw new ArgumentException("자를 영역이 이미지 밖에 있습니다.");
            return new Int32Rect(x, y, width, height);
        }

        private static BitmapSource ChangeDpi(BitmapSource image, int dpi, bool resample)
        {
            if (resample)
            {
                double scaleX = dpi / (image.DpiX > 0 ? image.DpiX : 96);
                double scaleY = dpi / (image.DpiY > 0 ? image.DpiY : 96);
                image = new TransformedBitmap(image, new ScaleTransform(scaleX, scaleY));
            }
            return WithDpi(image, dpi);
        }

        private static BitmapSource WithDpi(BitmapSource image, double dpi)
        {
            int stride = (image.PixelWidth * image.Format.BitsPerPixel + 7) / 8;
            var pixels = new byte[stride * image.PixelHeight];
            image.CopyPixels(pixels, stride, 0);
            return BitmapSource.Create(image.PixelWidth, image.PixelHeight, dpi, dpi, image.Format, image.Palette, pixels, stride);
        }

        private static void WriteImage(BitmapSource image, string output, string format)
        {
            if (format == "JPG" || format == "BMP") image = FlattenOnWhite(image);
            BitmapEncoder encoder = CreateEncoder(format);
            encoder.Frames.Add(BitmapFrame.Create(image));
            try
            {
                using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write))
                {
                    encoder.Save(stream);
                }
            }
            catch
            {
                TryDelete(output);
                throw;
            }
        }

        private static BitmapFrame LoadFrame(string source)
        {
            try
            {
                using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    BitmapDecoder decoder = BitmapDecoder.Create(input, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                    return decoder.Frames[0];
                }
            }
            catch (Exception ex)
            {
                if (ex is NotSupportedException || ex is FileFormatException || ex is ArgumentException)
                    throw new NotSupportedException("이미지를 읽을 수 없습니다. 파일이 손상됐거나 Windows 코덱이 없습니다.", ex);
                throw;
            }
        }

        public static BitmapSource LoadUpright(string source)
        {
            BitmapFrame frame = LoadFrame(source);
            BitmapSource image = ApplyOrientation(frame, ReadOrientation(frame));
            if (image.CanFreeze) image.Freeze();
            return image;
        }

        private static BitmapEncoder CreateEncoder(string format)
        {
            switch (format)
            {
                case "PNG": return new PngBitmapEncoder();
                case "JPG": return new JpegBitmapEncoder { QualityLevel = 90 };
                case "BMP": return new BmpBitmapEncoder();
                case "TIFF": return new TiffBitmapEncoder();
                case "GIF": return new GifBitmapEncoder();
                default: throw new NotSupportedException(format + " 형식은 지원하지 않습니다.");
            }
        }

        private static int ReadOrientation(BitmapFrame frame)
        {
            var metadata = frame.Metadata as BitmapMetadata;
            if (metadata == null) return 1;
            foreach (string query in new[] { "/app1/ifd/{ushort=274}", "/ifd/{ushort=274}" })
            {
                try
                {
                    object value = metadata.GetQuery(query);
                    if (value is ushort) return (ushort)value;
                }
                catch (Exception ex)
                {
                    if (!(ex is NotSupportedException || ex is ArgumentException || ex is InvalidOperationException)) throw;
                }
            }
            return 1;
        }

        // Output files carry no EXIF, so the rotation has to be baked into the pixels.
        private static BitmapSource ApplyOrientation(BitmapSource image, int orientation)
        {
            var transform = new TransformGroup();
            switch (orientation)
            {
                case 2: transform.Children.Add(new ScaleTransform(-1, 1)); break;
                case 3: transform.Children.Add(new RotateTransform(180)); break;
                case 4: transform.Children.Add(new ScaleTransform(1, -1)); break;
                case 5: transform.Children.Add(new ScaleTransform(-1, 1)); transform.Children.Add(new RotateTransform(270)); break;
                case 6: transform.Children.Add(new RotateTransform(90)); break;
                case 7: transform.Children.Add(new ScaleTransform(-1, 1)); transform.Children.Add(new RotateTransform(90)); break;
                case 8: transform.Children.Add(new RotateTransform(270)); break;
                default: return image;
            }
            return new TransformedBitmap(image, transform);
        }

        // JPG, BMP and the PDF pages have no alpha channel here; without this, transparent areas turn black.
        private static BitmapSource FlattenOnWhite(BitmapSource image)
        {
            var bgra = new FormatConvertedBitmap(image, PixelFormats.Bgra32, null, 0);
            int stride = bgra.PixelWidth * 4;
            var pixels = new byte[stride * bgra.PixelHeight];
            bgra.CopyPixels(pixels, stride, 0);
            for (int i = 0; i < pixels.Length; i += 4)
            {
                int alpha = pixels[i + 3];
                if (alpha == 255) continue;
                for (int c = 0; c < 3; c++)
                    pixels[i + c] = (byte)((pixels[i + c] * alpha + 255 * (255 - alpha)) / 255);
                pixels[i + 3] = 255;
            }
            double dpiX = image.DpiX > 0 ? image.DpiX : 96;
            double dpiY = image.DpiY > 0 ? image.DpiY : 96;
            return BitmapSource.Create(bgra.PixelWidth, bgra.PixelHeight, dpiX, dpiY, PixelFormats.Bgr32, null, pixels, stride);
        }

        // ---- PDF to images ----

        // Runs on a thread-pool thread because the Windows.Data.Pdf calls below block on async
        // operations, which must not happen on the STA conversion thread.
        private static List<string> PdfToImages(string source, string format, int dpi, CancellationToken cancel)
        {
            if (format == KeepFormat) format = "PNG";
            string target = format;
            try
            {
                return Task.Run(() => RenderPdf(source, target, dpi, cancel)).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                if (ex is IOException || ex is UnauthorizedAccessException || ex is NotSupportedException || ex is OperationCanceledException) throw;
                throw new InvalidOperationException("PDF를 열 수 없습니다. 파일이 손상됐거나 암호가 걸려 있을 수 있습니다. (" + ex.Message + ")", ex);
            }
        }

        private static List<string> RenderPdf(string source, string format, int dpi, CancellationToken cancel)
        {
            var outputs = new List<string>();
            try
            {
                StorageFile file = Wait(StorageFile.GetFileFromPathAsync(source));
                PdfDocument document = Wait(PdfDocument.LoadFromFileAsync(file));
                int digits = document.PageCount.ToString(CultureInfo.InvariantCulture).Length;
                for (uint i = 0; i < document.PageCount; i++)
                {
                    cancel.ThrowIfCancellationRequested();
                    BitmapSource page = RenderPage(document, i, dpi);
                    string suffix = document.PageCount == 1 ? "" : "_p" + (i + 1).ToString(CultureInfo.InvariantCulture).PadLeft(digits, '0');
                    string output = UniquePath(source, suffix, format.ToLowerInvariant());
                    WriteImage(page, output, format);
                    outputs.Add(output);
                    VerifySavedDpi(output, dpi);
                }
                return outputs;
            }
            catch
            {
                outputs.ForEach(TryDelete);
                throw;
            }
        }

        private static BitmapSource RenderPage(PdfDocument document, uint index, int dpi)
        {
            using (PdfPage page = document.GetPage(index))
            using (var memory = new InMemoryRandomAccessStream())
            {
                var options = new PdfPageRenderOptions
                {
                    DestinationWidth = (uint)Math.Max(1, Math.Round(page.Size.Width * dpi / 96)),
                    DestinationHeight = (uint)Math.Max(1, Math.Round(page.Size.Height * dpi / 96))
                };
                Wait(page.RenderToStreamAsync(memory, options));

                var png = new byte[(int)memory.Size];
                using (var reader = new DataReader(memory.GetInputStreamAt(0)))
                {
                    Wait(reader.LoadAsync((uint)png.Length));
                    reader.ReadBytes(png);
                }
                using (var stream = new MemoryStream(png))
                {
                    BitmapDecoder decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                    return WithDpi(decoder.Frames[0], dpi);
                }
            }
        }

        // The usual AsTask() helpers need a Windows SDK file this no-install build does not have.
        internal static T Wait<T>(IAsyncOperation<T> operation)
        {
            using (var done = new ManualResetEventSlim())
            {
                operation.Completed = (info, status) => done.Set();
                done.Wait();
                return operation.GetResults();
            }
        }

        internal static void Wait(IAsyncAction action)
        {
            using (var done = new ManualResetEventSlim())
            {
                action.Completed = (info, status) => done.Set();
                done.Wait();
                action.GetResults();
            }
        }

        // ---- Images to PDF ----

        // Windows has no PDF writer for images, but a PDF holding only images is simple to write directly.
        // Objects: 1 catalog, 2 page tree, then a page, its content stream and its image for each input.
        private static string ImagesToPdf(string[] files)
        {
            string[] ordered = files.OrderBy(f => f, StringComparer.CurrentCultureIgnoreCase).ToArray();
            foreach (string file in ordered)
            {
                if (!File.Exists(file)) throw new FileNotFoundException("파일을 찾을 수 없습니다.", file);
                if (KindOf(file) != FileKind.Image) throw new NotSupportedException("이미지만 PDF로 만들 수 있습니다.");
            }

            string output = UniquePath(ordered[0], "", "pdf");
            int objectCount = 2 + ordered.Length * 3;
            var offsets = new long[objectCount + 1];
            try
            {
                using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write))
                {
                    WriteAscii(stream, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
                    BeginObject(stream, offsets, 1);
                    WriteAscii(stream, "<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

                    BeginObject(stream, offsets, 2);
                    string kids = string.Join(" ", Enumerable.Range(0, ordered.Length).Select(i => (3 + i * 3) + " 0 R"));
                    WriteAscii(stream, "<< /Type /Pages /Count " + ordered.Length + " /Kids [" + kids + "] >>\nendobj\n");

                    for (int i = 0; i < ordered.Length; i++)
                        WritePdfPage(stream, offsets, 3 + i * 3, ordered[i]);

                    long xref = stream.Position;
                    var table = new StringBuilder();
                    table.Append("xref\n0 ").Append(objectCount + 1).Append("\n0000000000 65535 f \n");
                    for (int i = 1; i <= objectCount; i++)
                        table.Append(offsets[i].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
                    table.Append("trailer\n<< /Size ").Append(objectCount + 1).Append(" /Root 1 0 R >>\nstartxref\n")
                        .Append(xref.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
                    WriteAscii(stream, table.ToString());
                }
                return output;
            }
            catch
            {
                TryDelete(output);
                throw;
            }
        }

        private static void WritePdfPage(Stream stream, long[] offsets, int pageObject, string file)
        {
            const double MaxPagePoints = 14400;
            BitmapFrame frame = LoadFrame(file);
            int orientation = ReadOrientation(frame);
            BitmapSource image = ApplyOrientation(frame, orientation);

            string colorSpace, filter;
            byte[] data = TryRawJpeg(file, frame, orientation, out colorSpace);
            filter = "/DCTDecode";
            if (data == null)
            {
                data = ZlibRgb(image);
                colorSpace = "/DeviceRGB";
                filter = "/FlateDecode";
            }

            double dpiX = image.DpiX > 0 ? image.DpiX : 96;
            double dpiY = image.DpiY > 0 ? image.DpiY : 96;
            double width = image.PixelWidth * 72 / dpiX;
            double height = image.PixelHeight * 72 / dpiY;
            double fit = Math.Min(1, MaxPagePoints / Math.Max(width, height));
            string w = (width * fit).ToString("0.###", CultureInfo.InvariantCulture);
            string h = (height * fit).ToString("0.###", CultureInfo.InvariantCulture);

            BeginObject(stream, offsets, pageObject);
            WriteAscii(stream, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + w + " " + h + "] /Resources << /XObject << /Im0 "
                + (pageObject + 2) + " 0 R >> >> /Contents " + (pageObject + 1) + " 0 R >>\nendobj\n");

            string content = "q " + w + " 0 0 " + h + " 0 0 cm /Im0 Do Q\n";
            BeginObject(stream, offsets, pageObject + 1);
            WriteAscii(stream, "<< /Length " + content.Length + " >>\nstream\n" + content + "endstream\nendobj\n");

            BeginObject(stream, offsets, pageObject + 2);
            WriteAscii(stream, "<< /Type /XObject /Subtype /Image /Width " + image.PixelWidth + " /Height " + image.PixelHeight
                + " /ColorSpace " + colorSpace + " /BitsPerComponent 8 /Filter " + filter + " /Length " + data.Length + " >>\nstream\n");
            stream.Write(data, 0, data.Length);
            WriteAscii(stream, "\nendstream\nendobj\n");
        }

        // A JPEG can be embedded as is, so a photo is not compressed a second time.
        private static byte[] TryRawJpeg(string file, BitmapFrame frame, int orientation, out string colorSpace)
        {
            colorSpace = null;
            if (NormalizeFormat(Ext(file)) != "JPG" || orientation != 1) return null;
            if (frame.Format == PixelFormats.Gray8) colorSpace = "/DeviceGray";
            else if (frame.Format == PixelFormats.Bgr24 || frame.Format == PixelFormats.Bgr32) colorSpace = "/DeviceRGB";
            else return null;
            return File.ReadAllBytes(file);
        }

        private static byte[] ZlibRgb(BitmapSource image)
        {
            var rgb = new FormatConvertedBitmap(FlattenOnWhite(image), PixelFormats.Rgb24, null, 0);
            int stride = rgb.PixelWidth * 3;
            var pixels = new byte[stride * rgb.PixelHeight];
            rgb.CopyPixels(pixels, stride, 0);

            uint a = 1, b = 0;
            foreach (byte value in pixels)
            {
                a = (a + value) % 65521;
                b = (b + a) % 65521;
            }

            using (var memory = new MemoryStream())
            {
                memory.WriteByte(0x78);
                memory.WriteByte(0x9C);
                using (var deflate = new DeflateStream(memory, CompressionMode.Compress, true))
                {
                    deflate.Write(pixels, 0, pixels.Length);
                }
                uint adler = (b << 16) | a;
                memory.WriteByte((byte)(adler >> 24));
                memory.WriteByte((byte)(adler >> 16));
                memory.WriteByte((byte)(adler >> 8));
                memory.WriteByte((byte)adler);
                return memory.ToArray();
            }
        }

        private static void BeginObject(Stream stream, long[] offsets, int number)
        {
            offsets[number] = stream.Position;
            WriteAscii(stream, number + " 0 obj\n");
        }

        private static void WriteAscii(Stream stream, string text)
        {
            byte[] bytes = Encoding.GetEncoding(28591).GetBytes(text);
            stream.Write(bytes, 0, bytes.Length);
        }
    }

    // Changes a JPEG's DPI by editing only its header fields, leaving the compressed image bytes alone.
    internal static class JpegDpi
    {
        private const int InchUnit = 2;

        // Returns the edited file, or null if the file's structure is not understood.
        public static byte[] Patch(byte[] jpeg, int dpi)
        {
            if (jpeg.Length < 4 || jpeg[0] != 0xFF || jpeg[1] != 0xD8) return null;

            bool hasJfif = false;
            int pos = 2;
            while (pos + 4 <= jpeg.Length)
            {
                if (jpeg[pos] != 0xFF) return null;
                byte marker = jpeg[pos + 1];
                if (marker == 0xDA || marker == 0xD9) break;
                int length = (jpeg[pos + 2] << 8) | jpeg[pos + 3];
                if (length < 2 || pos + 2 + length > jpeg.Length) return null;

                int body = pos + 4;
                int bodyLength = length - 2;
                if (marker == 0xE0 && bodyLength >= 14 && Matches(jpeg, body, "JFIF\0"))
                {
                    jpeg[body + 7] = 1;
                    PutBigEndian16(jpeg, body + 8, dpi);
                    PutBigEndian16(jpeg, body + 10, dpi);
                    hasJfif = true;
                }
                else if (marker == 0xE1 && bodyLength >= 14 && Matches(jpeg, body, "Exif\0\0"))
                {
                    if (!PatchExif(jpeg, body + 6, bodyLength - 6, dpi)) return null;
                }
                pos += 2 + length;
            }

            return hasJfif ? jpeg : InsertJfif(jpeg, dpi);
        }

        private static bool Matches(byte[] data, int offset, string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (data[offset + i] != (byte)text[i]) return false;
            }
            return true;
        }

        private static void PutBigEndian16(byte[] data, int offset, int value)
        {
            data[offset] = (byte)(value >> 8);
            data[offset + 1] = (byte)value;
        }

        // Windows prefers the EXIF resolution over the JFIF one, so both have to agree.
        private static bool PatchExif(byte[] jpeg, int tiff, int size, int dpi)
        {
            bool little = jpeg[tiff] == (byte)'I' && jpeg[tiff + 1] == (byte)'I';
            bool big = jpeg[tiff] == (byte)'M' && jpeg[tiff + 1] == (byte)'M';
            if (!little && !big) return false;

            int ifd = ReadInt(jpeg, tiff + 4, 4, little);
            if (ifd < 8 || ifd + 2 > size) return false;
            int count = ReadInt(jpeg, tiff + ifd, 2, little);
            if (ifd + 2 + count * 12 > size) return false;

            for (int i = 0; i < count; i++)
            {
                int entry = tiff + ifd + 2 + i * 12;
                int tag = ReadInt(jpeg, entry, 2, little);
                if (tag == 282 || tag == 283)
                {
                    int value = ReadInt(jpeg, entry + 8, 4, little);
                    if (value < 8 || value + 8 > size) return false;
                    WriteInt(jpeg, tiff + value, 4, dpi, little);
                    WriteInt(jpeg, tiff + value + 4, 4, 1, little);
                }
                else if (tag == 296)
                {
                    WriteInt(jpeg, entry + 8, 2, InchUnit, little);
                }
            }
            return true;
        }

        private static int ReadInt(byte[] data, int offset, int bytes, bool little)
        {
            int value = 0;
            for (int i = 0; i < bytes; i++)
            {
                int b = data[little ? offset + bytes - 1 - i : offset + i];
                value = (value << 8) | b;
            }
            return value;
        }

        private static void WriteInt(byte[] data, int offset, int bytes, int value, bool little)
        {
            for (int i = 0; i < bytes; i++)
            {
                int shift = 8 * (little ? i : bytes - 1 - i);
                data[offset + i] = (byte)(value >> shift);
            }
        }

        private static byte[] InsertJfif(byte[] jpeg, int dpi)
        {
            byte[] segment =
            {
                0xFF, 0xE0, 0x00, 0x10, (byte)'J', (byte)'F', (byte)'I', (byte)'F', 0x00, 0x01, 0x01, 0x01,
                (byte)(dpi >> 8), (byte)dpi, (byte)(dpi >> 8), (byte)dpi, 0x00, 0x00
            };
            var result = new byte[jpeg.Length + segment.Length];
            result[0] = jpeg[0];
            result[1] = jpeg[1];
            Array.Copy(segment, 0, result, 2, segment.Length);
            Array.Copy(jpeg, 2, result, 2 + segment.Length, jpeg.Length - 2);
            return result;
        }
    }

    // Kept in the user's profile rather than the program folder, because it lists file paths
    // and the folder is meant to be copied to other people.
    internal static class Log
    {
        private const long MaxLogBytes = 1024 * 1024;
        private static readonly object Gate = new object();

        public static string FilePath
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Fanak", "fanak.log");
            }
        }

        public static void Write(string message)
        {
            lock (Gate)
            {
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                    if (File.Exists(FilePath) && new FileInfo(FilePath).Length > MaxLogBytes) File.Delete(FilePath);
                    File.AppendAllText(FilePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + message + Environment.NewLine, Encoding.UTF8);
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }
}
