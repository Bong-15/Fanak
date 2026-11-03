using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Fanak
{
    // Only a document's own program lays it out faithfully, so this drives the installed Word, PowerPoint, Excel or
    // Hangul without a window, for PDF and for moving between Hangul and Word. It is the one feature that needs
    // anything beyond Windows; the menu hides it without them.
    internal static class OfficePdf
    {
        // Long enough for a large document, since the only cost of waiting is on a hidden dialog.
        private const int TimeoutMinutes = 10;
        private const int StopPollMilliseconds = 100;
        private const string HangulProcess = "Hwp";
        private const string WordProgId = "Word.Application";
        private const string PowerPointProgId = "PowerPoint.Application";
        private const string ExcelProgId = "Excel.Application";
        private const string HangulProgId = "HWPFrame.HwpObject";
        private const string WordProcess = "WINWORD";
        private const string ExcelProcess = "EXCEL";

        private static readonly string[] WordExts = { "doc", "docx", "docm" };
        private static readonly string[] SlideExts = { "ppt", "pptx", "pptm" };
        private static readonly string[] SheetExts = { "xls", "xlsx", "xlsm" };
        private static readonly byte[] CompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        private static readonly Dictionary<string, bool> Installed = new Dictionary<string, bool>();
        private static string _activeProcess;
        private static HashSet<int> _activeBefore;
        private static volatile bool _stopped;

        private static string Ext(string path)
        {
            return Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        }

        // True when this file's program is installed.
        public static bool CanConvert(string path)
        {
            string ext = Ext(path);
            if (WordExts.Contains(ext) || TextDocs.WordOpensExts.Contains(ext)) return IsInstalled(WordProgId);
            if (SlideExts.Contains(ext)) return IsInstalled(PowerPointProgId);
            if (SheetExts.Contains(ext)) return IsInstalled(ExcelProgId);
            return HangulFile.Exts.Contains(ext) && IsInstalled(HangulProgId);
        }

        public static bool CanMakeDocx(string path)
        {
            return HangulFile.Exts.Contains(Ext(path)) && IsInstalled(HangulProgId);
        }

        // Hangul has no filter for the old .doc format, so Word has to turn that into a .docx first.
        public static bool CanMakeHwp(string path)
        {
            if (!WordExts.Contains(Ext(path)) || !IsInstalled(HangulProgId)) return false;
            return Ext(path) != "doc" || IsInstalled(WordProgId);
        }

        private static bool IsInstalled(string progId)
        {
            lock (Installed)
            {
                bool installed;
                if (!Installed.TryGetValue(progId, out installed))
                {
                    installed = Type.GetTypeFromProgID(progId) != null;
                    Installed[progId] = installed;
                }
                return installed;
            }
        }

        public static string Convert(string source)
        {
            string ext = Ext(source);
            bool word = WordExts.Contains(ext);
            bool hangul = HangulFile.Exts.Contains(ext);
            bool sheet = SheetExts.Contains(ext);
            bool text = TextDocs.WordOpensExts.Contains(ext);
            string processName = word || text ? WordProcess : sheet ? ExcelProcess : "POWERPNT";
            string program = word || text ? "Word가" : hangul ? "한글이" : sheet ? "Excel이" : "PowerPoint가";
            if (!CanConvert(source))
                throw new InvalidOperationException(program + " 설치되어 있어야 PDF로 바꿀 수 있습니다.");
            if (hangul) return ToFormat(source, "PDF", "pdf");
            if (text && ext != "rtf") return ConvertMarkedText(source, ext);
            if (ext == "rtf" && !StartsWith(source, Encoding.ASCII.GetBytes("{\\rtf")))
                throw new InvalidOperationException("RTF 문서가 아니거나 손상된 파일입니다.");

            string output = Converter.UniquePath(source, "", "pdf");
            Export(processName, program, output, before =>
            {
                if (word || text) ExportWord(source, output, 0, Type.Missing);
                else if (sheet) ExportSheets(source, output);
                else ExportSlides(source, output, before.Count == 0);
            });
            return output;
        }

        public static string ToDocx(string source)
        {
            return ToFormat(source, "OOXML", "docx");
        }

        public static string ToHwp(string source)
        {
            if (!StartsWith(source, CompoundSignature) && !StartsWith(source, Encoding.ASCII.GetBytes("PK")) && !StartsWith(source, Encoding.ASCII.GetBytes("{\\rtf")))
                throw new InvalidOperationException("Word 문서가 아니거나 손상된 파일입니다.");
            if (Ext(source) != "doc") return ToFormat(source, "HWP", "hwp");

            return InPrivateFolder(folder =>
            {
                string docx = Path.Combine(folder, Path.GetFileNameWithoutExtension(source) + ".docx");
                Export(WordProcess, "Word가", docx, before => SaveWordAsDocx(source, docx));
                return ConvertHangul(docx, "HWP", "hwp", source);
            });
        }

        private static string ToFormat(string source, string saveFormat, string ext)
        {
            if (HangulFile.Exts.Contains(Ext(source)) && !HangulFile.LooksValid(source))
                throw new InvalidOperationException("한글 문서가 아니거나 손상된 파일입니다.");
            return ConvertHangul(source, saveFormat, ext, source);
        }

        // Word is given a copy in a form it can't misread, written beside nothing the user owns.
        private static string ConvertMarkedText(string source, string ext)
        {
            const int FormatEncodedText = 5;
            // Word's own "web page" format setting reports every page as corrupt; left to detect the type, it opens them.
            const int FormatDetect = 0;
            const int Utf8CodePage = 65001;

            return InPrivateFolder(folder =>
            {
                string copy = Path.Combine(folder, Path.GetFileName(source));
                File.WriteAllBytes(copy, TextDocs.PrepareForWord(source, ext));
                string pdf = Path.ChangeExtension(copy, "pdf");
                Export(WordProcess, "Word가", pdf, before => ExportWord(copy, pdf, ext == "txt" ? FormatEncodedText : FormatDetect, Utf8CodePage));
                string output = Converter.UniquePath(source, "", "pdf");
                File.Move(pdf, output);
                return output;
            });
        }

        private static string InPrivateFolder(Func<string, string> work)
        {
            string folder = Path.Combine(Path.GetTempPath(), "Fanak", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            try
            {
                return work(folder);
            }
            finally
            {
                try { Directory.Delete(folder, true); }
                catch (IOException ex) { Log.Write("임시 폴더 삭제 실패: " + folder + " - " + ex.Message); }
                catch (UnauthorizedAccessException ex) { Log.Write("임시 폴더 삭제 실패: " + folder + " - " + ex.Message); }
            }
        }

        // Hangul works on a private copy. The original may be locked by an open window or sit on a network
        // drive that a program started through COM can't see, and neither should matter.
        private static string ConvertHangul(string source, string saveFormat, string ext, string outputBeside)
        {
            StopLeftoverHangul();
            // Conversion normally leaves open windows alone, but a test window once vanished during one, for no
            // reason found in five repeats. Asking is a small price against losing someone's unsaved work.
            if (ProcessIds(HangulProcess).Any())
                throw new InvalidOperationException("열려 있는 한글 창을 저장하고 닫은 뒤 다시 변환해 주세요.");
            return InPrivateFolder(folder =>
            {
                string copy = Path.Combine(folder, Path.GetFileName(source));
                File.WriteAllBytes(copy, HangulFile.ReadShared(source));
                string result = Path.ChangeExtension(copy, ext);
                Export(HangulProcess, "한글이", result, before => ExportHangul(copy, result, saveFormat, ext, before));
                string output = Converter.UniquePath(outputBeside, "", ext);
                File.Move(result, output);
                return output;
            });
        }

        private static void Export(string processName, string program, string output, Action<HashSet<int>> export)
        {
            HashSet<int> before = ProcessIds(processName);
            Exception failure = null;
            lock (Installed)
            {
                _activeProcess = processName;
                _activeBefore = before;
                _stopped = false;
            }

            // Office may wait forever on a dialog nobody can see, so it runs where it can be abandoned.
            var worker = new Thread(() =>
            {
                try
                {
                    export(before);
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();

            try
            {
                // A COM call into a program just stopped can take half a minute to fail, so nobody waits for it.
                var waited = Stopwatch.StartNew();
                bool finished = false;
                while (!finished && !_stopped && waited.Elapsed < TimeSpan.FromMinutes(TimeoutMinutes))
                    finished = worker.Join(StopPollMilliseconds);
                if (!finished && _stopped)
                {
                    WaitForStarted(processName, before);
                    Converter.TryDelete(output);
                    throw new OperationCanceledException();
                }
                if (!finished)
                {
                    StopStarted(processName, before);
                    Converter.TryDelete(output);
                    throw new TimeoutException(program + " 응답하지 않아 변환을 멈췄습니다. 암호가 걸렸거나 확인 창이 필요한 문서일 수 있습니다.");
                }

                if (failure != null)
                {
                    StopStarted(processName, before);
                    Converter.TryDelete(output);
                    Log.Write(processName + " 오류: " + failure.Message);
                    if (failure is InvalidOperationException) throw new InvalidOperationException(failure.Message, failure);
                    throw new InvalidOperationException("변환하지 못했습니다. 암호가 걸렸거나 손상된 문서일 수 있습니다.", failure);
                }

                if (!IsFinished(output))
                {
                    Converter.TryDelete(output);
                    throw new InvalidOperationException("변환한 파일이 제대로 만들어지지 않았습니다.");
                }

                // Hangul can outlive Quit, still answering COM, and the next conversion would connect to it.
                // Excel stays until its COM references are collected, which on this thread is never.
                if (processName == HangulProcess || processName == ExcelProcess) WaitForStarted(processName, before);
            }
            finally
            {
                lock (Installed) _activeProcess = null;
            }
        }

        // Called on the way out of the app, so a conversion cut short leaves no hidden program holding a file.
        public static void StopActive()
        {
            lock (Installed)
            {
                if (_activeProcess != null && StopStarted(_activeProcess, _activeBefore)) _stopped = true;
            }
        }

        private static void ExportWord(string source, string output, int openFormat, object encoding)
        {
            const int FormatPdf = 17;
            WithWordDocument(source, openFormat, encoding, document => document.ExportAsFixedFormat(OutputFileName: output, ExportFormat: FormatPdf));
        }

        private static void SaveWordAsDocx(string source, string output)
        {
            const int FormatDocx = 16;
            WithWordDocument(source, 0, Type.Missing, document => document.SaveAs2(output, FormatDocx));
        }

        private static void WithWordDocument(string source, int openFormat, object encoding, Action<dynamic> work)
        {
            const int AlertsNone = 0;
            const int AutomationSecurityForceDisable = 3;
            const int DoNotSave = 0;

            dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID(WordProgId));
            // Word normally starts a fresh process for this. Should it ever attach to one with the user's
            // documents open, then hiding it or quitting it could lose their work.
            bool ours = word.Documents.Count == 0;
            try
            {
                if (ours)
                {
                    word.Visible = false;
                    word.DisplayAlerts = AlertsNone;
                    word.AutomationSecurity = AutomationSecurityForceDisable;
                }
                // With any password given, a locked document fails here instead of asking for one on a hidden window.
                dynamic document = word.Documents.Open(
                    FileName: source, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false,
                    PasswordDocument: "fanak", Format: openFormat, Encoding: encoding, Visible: false, NoEncodingDialog: true);
                try
                {
                    work(document);
                }
                finally
                {
                    document.Close(SaveChanges: DoNotSave);
                }
            }
            finally
            {
                if (ours && word.Documents.Count == 0) word.Quit(SaveChanges: DoNotSave);
            }
        }

        private static void ExportSlides(string source, string output, bool startedByUs)
        {
            const int MsoTrue = -1;
            const int MsoFalse = 0;
            const int SaveAsPdf = 32;

            // PowerPoint keeps one process per user, so this may be the window the user is working in.
            dynamic powerPoint = Activator.CreateInstance(Type.GetTypeFromProgID(PowerPointProgId));
            try
            {
                dynamic presentation = powerPoint.Presentations.Open(source, MsoTrue, MsoFalse, MsoFalse);
                try
                {
                    presentation.SaveAs(output, SaveAsPdf);
                }
                finally
                {
                    presentation.Close();
                }
            }
            finally
            {
                if (startedByUs && powerPoint.Presentations.Count == 0) powerPoint.Quit();
            }
        }

        private static void ExportSheets(string source, string output)
        {
            const int AutomationSecurityForceDisable = 3;
            const int DontUpdateLinks = 0;
            const int TypePdf = 0;

            dynamic excel = Activator.CreateInstance(Type.GetTypeFromProgID(ExcelProgId));
            // Opening a file that is already open hands back the user's own workbook, and closing that would lose work.
            if (excel.Workbooks.Count != 0)
                throw new InvalidOperationException("다른 프로그램이 Excel을 사용 중입니다. 잠시 뒤 다시 변환해 주세요.");
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.AutomationSecurity = AutomationSecurityForceDisable;
                dynamic book = excel.Workbooks.Open(
                    Filename: source, UpdateLinks: DontUpdateLinks, ReadOnly: true, Password: "fanak",
                    IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false);
                try
                {
                    foreach (dynamic sheet in book.Worksheets) FitToWidth(sheet);
                    book.ExportAsFixedFormat(Type: TypePdf, Filename: output);
                }
                finally
                {
                    book.Close(SaveChanges: false);
                }
            }
            finally
            {
                excel.Quit();
            }
        }

        // Left alone, a wide sheet prints its columns on separate pages, which nobody wants from a PDF. This changes
        // only the open copy. A sheet already set up for printing is the author's choice, so it stays as it is.
        private static void FitToWidth(dynamic sheet)
        {
            const int Landscape = 2;
            const double PortraitWidthPoints = 500;

            try
            {
                dynamic setup = sheet.PageSetup;
                if (setup.Zoom is bool || !string.IsNullOrEmpty((string)setup.PrintArea)) return;
                if ((double)sheet.UsedRange.Width > PortraitWidthPoints) setup.Orientation = Landscape;
                setup.Zoom = false;
                setup.FitToPagesWide = 1;
                setup.FitToPagesTall = false;
            }
            catch (COMException ex)
            {
                Log.Write("시트 쪽 맞춤 실패: " + ex.Message);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
            {
                Log.Write("시트 쪽 맞춤 실패: " + ex.Message);
            }
        }

        private static void ExportHangul(string source, string output, string saveFormat, string ext, HashSet<int> before)
        {
            const int DiscardChanges = 1;
            // Hancom's bits for: OK on an OK box, OK on OK/Cancel, Yes on Yes/No/Cancel, Yes on Yes/No.
            const int AnswerOkAndYes = 0x00011011;

            dynamic hangul = Activator.CreateInstance(Type.GetTypeFromProgID(HangulProgId));
            // A window the user opened gets no COM connections, so a new process normally starts. If none did,
            // this reached a Hangul that something else is automating, and opening a file would replace its document.
            if (!ProcessIds(HangulProcess).Except(before).Any())
                throw new InvalidOperationException("다른 프로그램이 한글을 사용 중입니다. 잠시 뒤 다시 변환해 주세요.");
            try
            {
                hangul.XHwpWindows.Item(0).Visible = false;
                // Saving as .docx asks "저장을 계속할까요?" even with no visible window, and waits for an answer.
                hangul.SetMessageBoxMode(AnswerOkAndYes);
                if (!(bool)hangul.Open(source, "", "forceopen:true;suspendpassword:true;versionwarning:false"))
                    throw new InvalidOperationException("문서를 열지 못했습니다. 암호가 걸렸거나 손상된 문서일 수 있습니다.");
                if (!(bool)hangul.SaveAs(output, saveFormat, ""))
                    throw new InvalidOperationException(ext.ToUpperInvariant() + "로 저장하지 못했습니다. 저장이 제한된 문서일 수 있습니다.");
            }
            finally
            {
                hangul.Clear(DiscardChanges);
                hangul.Quit();
                Marshal.FinalReleaseComObject((object)hangul);
            }
        }

        // A hidden Hangul started for automation has no one to answer to once its client is gone, and it
        // keeps its document locked. Windows the user opened don't match, nor do any that are showing.
        private static void StopLeftoverHangul()
        {
            using (var search = new ManagementObjectSearcher("SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name='Hwp.exe'"))
            using (ManagementObjectCollection found = search.Get())
            {
                foreach (ManagementBaseObject item in found)
                {
                    string command = item["CommandLine"] as string;
                    if (command == null || command.IndexOf("-Automation", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    try
                    {
                        using (Process process = Process.GetProcessById((int)(uint)item["ProcessId"]))
                        {
                            if (process.MainWindowHandle != IntPtr.Zero) continue;
                            Log.Write("남아 있던 숨은 한글을 종료함: " + process.Id);
                            process.Kill();
                            process.WaitForExit(5000);
                        }
                    }
                    catch (ArgumentException) { }
                    catch (InvalidOperationException ex) { Log.Write("Office 종료 실패: " + ex.Message); }
                    catch (System.ComponentModel.Win32Exception ex) { Log.Write("Office 종료 실패: " + ex.Message); }
                }
            }
        }

        private static void WaitForStarted(string name, HashSet<int> before)
        {
            const int GraceMilliseconds = 5000;
            const int PollMilliseconds = 100;
            for (int waited = 0; waited < GraceMilliseconds; waited += PollMilliseconds)
            {
                if (!ProcessIds(name).Except(before).Any()) return;
                Thread.Sleep(PollMilliseconds);
            }
            StopStarted(name, before);
        }

        private static HashSet<int> ProcessIds(string name)
        {
            var ids = new HashSet<int>();
            foreach (Process process in Process.GetProcessesByName(name))
            {
                ids.Add(process.Id);
                process.Dispose();
            }
            return ids;
        }

        // Only a process that appeared during this conversion; whatever ran before belongs to the user.
        private static bool StopStarted(string name, HashSet<int> before)
        {
            bool stopped = false;
            foreach (Process process in Process.GetProcessesByName(name))
            {
                try
                {
                    if (!before.Contains(process.Id) && process.MainWindowHandle == IntPtr.Zero)
                    {
                        process.Kill();
                        stopped = true;
                    }
                }
                catch (InvalidOperationException ex) { Log.Write("Office 종료 실패: " + ex.Message); }
                catch (System.ComponentModel.Win32Exception ex) { Log.Write("Office 종료 실패: " + ex.Message); }
                finally { process.Dispose(); }
            }
            return stopped;
        }

        private static bool IsFinished(string path)
        {
            if (!File.Exists(path)) return false;
            switch (Ext(path))
            {
                case "pdf": return StartsWith(path, Encoding.ASCII.GetBytes("%PDF-"));
                case "docx": return StartsWith(path, Encoding.ASCII.GetBytes("PK"));
                case "hwp": return StartsWith(path, CompoundSignature);
                default: return false;
            }
        }

        private static bool StartsWith(string path, byte[] signature)
        {
            var head = new byte[signature.Length];
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                return input.Read(head, 0, head.Length) == head.Length && head.SequenceEqual(signature);
            }
        }
    }
}
