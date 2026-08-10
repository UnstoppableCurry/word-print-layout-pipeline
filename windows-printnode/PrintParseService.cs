// PrintParseService.cs
// 零外部依赖：用 .NET Framework 4.x 自带 csc.exe 编译。
// 编译命令见 setup.ps1。
// 功能：HTTP 服务 -> 接收 doc/docx -> 打印引擎打印到 Microsoft XPS Document Writer
//       -> 解析 XPS 文字+坐标(WPF Glyphs) -> 渲染页面 PNG(RenderTargetBitmap) -> 返回 JSON。
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Collections.Concurrent;
using System.Web.Script.Serialization;   // System.Web.Extensions.dll
using System.Windows;                    // PresentationCore / WindowsBase
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;      // ReachFramework.dll
using System.IO.Packaging;               // WindowsBase.dll

namespace PrintNode
{
    // ===================== 打印引擎抽象（可替换模块） =====================
    // TODO(Word): 换成真 Word 时，实现 WordComPrintEngine（Word.Application
    //   Documents.Open + PrintOut(OutputFileName:xpsPath, PrintToFile:true,
    //   ActivePrinter:"Microsoft XPS Document Writer")），
    //   并把 Program.CreateEngine() 里一行切换即可，其余管线不变。
    public interface IPrintEngine
    {
        string Name { get; }
        // 把 inputPath 打印成 XPS，输出到 xpsPath。实现可忽略 xpsPath 而使用固定端口文件，但返回实际产生的文件路径。
        string PrintToXps(string inputPath, string xpsPath);
    }

    // 占位引擎：WordPad 命令行打印。排版≠Word，仅用于打通管线验证。
    public class WordpadPrintEngine : IPrintEngine
    {
        // 打印机端口被 setup.ps1 固定为该文件路径（Local Port 技巧：XPS 驱动直接写文件不弹窗）
        public const string FixedPortFile = @"C:\printnode\print\output.xps";
        public const string PrinterName = "Microsoft XPS Document Writer";

        public string Name { get { return "wordpad-placeholder"; } }

        public string PrintToXps(string inputPath, string xpsPath)
        {
            string wordpad = Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Windows NT\Accessories\wordpad.exe");
            if (!File.Exists(wordpad))
                wordpad = Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Windows NT\Accessories\wordpad.exe");
            if (File.Exists(FixedPortFile)) File.Delete(FixedPortFile);

            var psi = new System.Diagnostics.ProcessStartInfo();
            psi.FileName = wordpad;
            psi.Arguments = "/pt \"" + inputPath + "\" \"" + PrinterName + "\"";
            psi.CreateNoWindow = true;
            psi.UseShellExecute = false;
            var p = System.Diagnostics.Process.Start(psi);

            // 等待 wordpad 退出 + 后台 spool 写完（文件尺寸稳定）
            var deadline = DateTime.Now.AddSeconds(180);
            long lastSize = -1; int stableCount = 0;
            while (DateTime.Now < deadline)
            {
                Thread.Sleep(1000);
                if (File.Exists(FixedPortFile))
                {
                    long sz = new FileInfo(FixedPortFile).Length;
                    if (sz > 0 && sz == lastSize)
                    {
                        stableCount++;
                        if (stableCount >= 3 && (p.HasExited)) break;
                    }
                    else { stableCount = 0; lastSize = sz; }
                }
            }
            try { if (!p.HasExited) p.Kill(); } catch { }
            if (!File.Exists(FixedPortFile) || new FileInfo(FixedPortFile).Length == 0)
                throw new Exception("wordpad print timeout: no XPS produced within 180s");
            File.Copy(FixedPortFile, xpsPath, true);
            return xpsPath;
        }
    }

    // 真 Word 引擎：COM 后期绑定（dynamic，无需 Office PIA，csc 加 /r:Microsoft.CSharp.dll）。
    // 打印机端口已被 setup.ps1 固定为 FixedPortFile（Local Port 技巧），PrintOut 直写 XPS 不弹窗。
    public class WordComPrintEngine : IPrintEngine
    {
        public const string FixedPortFile = @"C:\printnode\print\output.xps";
        public const string PrinterName = "Microsoft XPS Document Writer";

        private dynamic _word;   // Word.Application，跨作业复用

        public string Name { get { return "word"; } }

        public static bool IsAvailable()
        {
            try { return Type.GetTypeFromProgID("Word.Application") != null; }
            catch { return false; }
        }

        private dynamic GetWord()
        {
            if (_word != null) return _word;
            Type t = Type.GetTypeFromProgID("Word.Application");
            if (t == null) throw new Exception("Word.Application ProgID not found (Office not installed?)");
            dynamic w = Activator.CreateInstance(t);
            w.Visible = false;
            w.DisplayAlerts = 0;              // wdAlertsNone：禁止一切弹窗阻塞自动化
            w.ScreenUpdating = false;
            _word = w;
            return w;
        }

        public string PrintToXps(string inputPath, string xpsPath)
        {
            if (File.Exists(FixedPortFile)) File.Delete(FixedPortFile);

            dynamic word = GetWord();
            // 切到 XPS 打印机（固定文件端口）
            try { word.ActivePrinter = PrinterName; }
            catch (Exception ex) { throw new Exception("set ActivePrinter failed: " + ex.Message); }

            dynamic doc = null;
            try
            {
                // 不传 Visible:false：隐藏窗口会导致 PrintOut 报“文档窗口处于非活动状态”
                doc = word.Documents.Open(inputPath, ReadOnly: true, AddToRecentFiles: false);
                try { doc.Activate(); } catch { }
                doc.PrintOut(Background: false);   // 同步等待送入后台打印
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(false); } catch { }
                }
            }

            // 等待 spool 写完（文件尺寸稳定）
            var deadline = DateTime.Now.AddSeconds(300);
            long lastSize = -1; int stableCount = 0;
            while (DateTime.Now < deadline)
            {
                Thread.Sleep(1000);
                if (File.Exists(FixedPortFile))
                {
                    long sz = new FileInfo(FixedPortFile).Length;
                    if (sz > 0 && sz == lastSize)
                    {
                        stableCount++;
                        if (stableCount >= 3) break;
                    }
                    else { stableCount = 0; lastSize = sz; }
                }
            }
            if (!File.Exists(FixedPortFile) || new FileInfo(FixedPortFile).Length == 0)
                throw new Exception("word print timeout: no XPS produced within 300s");
            File.Copy(FixedPortFile, xpsPath, true);
            return xpsPath;
        }
    }

    // ===================== 数据结构 =====================
    public class LineInfo
    {
        public string text; public double x0; public double top; public double x1; public double bottom;
    }
    public class PageInfo
    {
        public int page; public double width; public double height;
        public string image_url; public List<LineInfo> lines = new List<LineInfo>();
    }

    // ===================== 作业（STA 线程内执行） =====================
    public class Job
    {
        public string Id;
        public string InputPath;      // 已保存的上传文件
        public string OriginalName;
        public string JobDir;
        public HttpListenerContext Ctx;
    }

    public class StaWorker
    {
        private readonly BlockingCollection<Job> _q = new BlockingCollection<Job>();
        private readonly IPrintEngine _engine;
        public StaWorker(IPrintEngine engine)
        {
            _engine = engine;
            var t = new Thread(Loop);
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);   // WPF 渲染必须 STA
            t.Start();
        }
        public void Enqueue(Job j) { _q.Add(j); }

        private void Loop()
        {
            foreach (var job in _q.GetConsumingEnumerable())
            {
                try { Process(job); }
                catch (Exception ex) { RespondError(job.Ctx, 500, Diag(ex)); }
            }
        }

        private void Process(Job job)
        {
            string xpsPath = Path.Combine(job.JobDir, "result.xps");
            _engine.PrintToXps(job.InputPath, xpsPath);
            var pages = ParseXpsPackage(xpsPath, job);

            var result = new Dictionary<string, object>();
            result["engine"] = _engine.Name;
            result["job"] = job.Id;
            result["file"] = job.OriginalName;
            result["page_count"] = pages.Count;
            result["pages"] = pages;
            RespondJson(job.Ctx, 200, result);
        }

        // ---- 手动 OPC 导航：绕过 XpsDocument.GetFixedDocumentSequence() 对
        // ---- 打印驱动直写（interleaved）XPS 返回 null 的问题 ----
        private List<PageInfo> ParseXpsPackage(string xpsPath, Job job)
        {
            var pages = new List<PageInfo>();
            // pack URI：authority 直接做包标识（无逗号转义，避免 generic parser 端口歧义）
            Uri pkgUri = new Uri("pack://pn" + job.Id + ".xps");
            Package pkg = Package.Open(xpsPath, FileMode.Open, FileAccess.Read);
            PackageStore.AddPackage(pkgUri, pkg);
            try
            {
                PackagePart fdseq = FindPart(pkg, ".fdseq");
                if (fdseq == null) throw new Exception("no .fdseq part; parts=" + ListParts(pkg));
                foreach (string fdocRef in ReadRefs(fdseq, "DocumentReference"))
                {
                    PackagePart fdoc = pkg.GetPart(ResolvePartUri(fdseq.Uri, fdocRef));
                    foreach (string pageRef in ReadRefs(fdoc, "PageContent"))
                    {
                        PackagePart fpage = pkg.GetPart(ResolvePartUri(fdoc.Uri, pageRef));
                        pages.Add(RenderPage(pkgUri, fpage, pages.Count + 1, job));
                    }
                }
                if (pages.Count == 0) throw new Exception("no pages found; parts=" + ListParts(pkg));
            }
            finally
            {
                PackageStore.RemovePackage(pkgUri);
                pkg.Close();
            }
            return pages;
        }

        private static string ListParts(Package pkg)
        {
            var sb = new StringBuilder();
            foreach (var p in pkg.GetParts()) { sb.Append(p.Uri).Append(';'); }
            return sb.ToString();
        }

        private static PackagePart FindPart(Package pkg, string suffix)
        {
            foreach (var p in pkg.GetParts())
                if (p.Uri.OriginalString.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) return p;
            return null;
        }

        private static List<string> ReadRefs(PackagePart part, string elementName)
        {
            var refs = new List<string>();
            using (var s = part.GetStream(FileMode.Open, FileAccess.Read))
            using (var xr = System.Xml.XmlReader.Create(s))
            {
                while (xr.Read())
                {
                    if (xr.NodeType == System.Xml.XmlNodeType.Element && xr.LocalName == elementName)
                    {
                        string src = xr.GetAttribute("Source");
                        if (!string.IsNullOrEmpty(src)) refs.Add(src);
                    }
                }
            }
            return refs;
        }

        private static Uri ResolvePartUri(Uri ownerPartUri, string reference)
        {
            if (reference.StartsWith("/")) return new Uri(reference, UriKind.Relative);
            return PackUriHelper.ResolvePartUri(ownerPartUri, new Uri(reference, UriKind.Relative));
        }

        // 用 XamlReader + pack URI 加载 FixedPage：字体（包内混淆 odttf）由 WPF 经 PackageStore 解析
        private PageInfo RenderPage(Uri pkgUri, PackagePart fpage, int pageNo, Job job)
        {
            var pi = new PageInfo();
            pi.page = pageNo;
            Uri partPackUri = PackUriHelper.Create(pkgUri, fpage.Uri);
            var pctx = new ParserContext();
            pctx.BaseUri = partPackUri;
            FixedPage page;
            string xaml;
            using (var s = fpage.GetStream(FileMode.Open, FileAccess.Read))
                xaml = new StreamReader(s, Encoding.UTF8).ReadToEnd();
            // Win10 XPS 驱动输出 OpenXPS 命名空间，WPF 只认 MS XPS 命名空间：直接替换
            xaml = xaml.Replace("http://schemas.openxps.org/oxps/v1.0",
                                "http://schemas.microsoft.com/xps/2005/06");
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(xaml)))
                page = (FixedPage)System.Windows.Markup.XamlReader.Load(ms, pctx);
            double w = page.Width > 0 ? page.Width : 793.7;
            double h = page.Height > 0 ? page.Height : 1122.5;
            pi.width = Math.Round(w, 1);
            pi.height = Math.Round(h, 1);
            page.Measure(new Size(w, h));
            page.Arrange(new Rect(0, 0, w, h));
            page.UpdateLayout();

            // --- 渲染 PNG（2x 便于肉眼校验） ---
            string imgName = "page" + pageNo + ".png";
            string imgPath = Path.Combine(job.JobDir, imgName);
            int pw = Math.Max(1, (int)Math.Round(w * 2));
            int ph = Math.Max(1, (int)Math.Round(h * 2));
            var rtb = new RenderTargetBitmap(pw, ph, 192, 192, PixelFormats.Pbgra32);
            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                dc.DrawRectangle(Brushes.White, null, new Rect(0, 0, w, h));
                dc.DrawRectangle(new VisualBrush(page), null, new Rect(0, 0, w, h));
            }
            rtb.Render(dv);
            var enc = new PngBitmapEncoder();
            enc.Frames.Add(BitmapFrame.Create(rtb));
            using (var fs = File.Create(imgPath)) enc.Save(fs);
            pi.image_url = "/images/" + job.Id + "/" + imgName;

            // --- 提取文字 + 坐标 ---
            ExtractGlyphs(page, page, pi.lines);
            pi.lines = MergeToLines(pi.lines);
            return pi;
        }

        public static string Diag(Exception ex)
        {
            return ex.GetType().Name + ": " + ex.Message + " | " + ex.StackTrace;
        }

        private static void ExtractGlyphs(Visual node, Visual root, List<LineInfo> acc)
        {
            var g = node as Glyphs;
            if (g != null && !string.IsNullOrEmpty(g.UnicodeString))
            {
                try
                {
                    // 统一到页面坐标系
                    Point origin = g.TransformToAncestor(root).Transform(new Point(g.OriginX, g.OriginY));
                    double em = g.FontRenderingEmSize;
                    double pageW = root is FixedPage ? ((FixedPage)root).Width : 10000;
                    double width = -1;
                    try
                    {
                        var run = g.ToGlyphRun();
                        if (run != null)
                        {
                            var geo = run.BuildGeometry();
                            if (geo != null && geo.Bounds.Width > 0)
                                width = geo.Bounds.Right;   // bounds 相对 baseline origin
                        }
                    }
                    catch { }
                    if (width <= 0 || origin.X + width > pageW + 10)
                    {
                        // 估算：全角 1.0em，半角 0.55em
                        double est = 0;
                        foreach (char ch in g.UnicodeString) est += (ch > 0x2E7F ? 1.0 : 0.55);
                        width = est * em;
                    }
                    var li = new LineInfo();
                    li.text = g.UnicodeString;
                    li.x0 = Math.Round(origin.X, 1);
                    li.x1 = Math.Round(Math.Min(origin.X + width, pageW), 1);
                    li.top = Math.Round(origin.Y - em * 0.88, 1);
                    li.bottom = Math.Round(origin.Y + em * 0.25, 1);
                    acc.Add(li);
                }
                catch { }
            }
            int count = VisualTreeHelper.GetChildrenCount(node);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(node, i) as Visual;
                if (child != null) ExtractGlyphs(child, root, acc);
            }
        }

        // 按基线聚类成行，行内按 x 排序拼接
        private static List<LineInfo> MergeToLines(List<LineInfo> glyphs)
        {
            var sorted = new List<LineInfo>(glyphs);
            sorted.Sort((a, b) => a.top != b.top ? a.top.CompareTo(b.top) : a.x0.CompareTo(b.x0));
            var lines = new List<LineInfo>();
            const double yTol = 3.0;
            foreach (var g in sorted)
            {
                LineInfo cur = lines.Count > 0 ? lines[lines.Count - 1] : null;
                double curBaseline = cur != null ? cur.bottom - (cur.bottom - cur.top) * 0.22 : 0;
                double gBaseline = g.bottom - (g.bottom - g.top) * 0.22;
                if (cur != null && Math.Abs(gBaseline - curBaseline) <= yTol)
                {
                    double gap = g.x0 - cur.x1;
                    cur.text += (gap > 2.0 ? " " : "") + g.text;
                    cur.x1 = Math.Max(cur.x1, g.x1);
                    cur.top = Math.Min(cur.top, g.top);
                    cur.bottom = Math.Max(cur.bottom, g.bottom);
                }
                else
                {
                    lines.Add(new LineInfo { text = g.text, x0 = g.x0, top = g.top, x1 = g.x1, bottom = g.bottom });
                }
            }
            return lines;
        }

        public static void RespondJson(HttpListenerContext ctx, int status, object payload)
        {
            var ser = new JavaScriptSerializer();
            ser.MaxJsonLength = 50 * 1024 * 1024;
            string json = ser.Serialize(payload);
            byte[] buf = Encoding.UTF8.GetBytes(json);
            ctx.Response.StatusCode = status;
            ctx.Response.ContentType = "application/json; charset=utf-8";
            ctx.Response.ContentLength64 = buf.Length;
            ctx.Response.OutputStream.Write(buf, 0, buf.Length);
            ctx.Response.OutputStream.Close();
        }

        public static void RespondError(HttpListenerContext ctx, int status, string msg)
        {
            try
            {
                RespondJson(ctx, status, new Dictionary<string, object> { { "error", msg } });
            }
            catch { }
        }
    }

    // ===================== HTTP 服务 =====================
    public class Program
    {
        public const string RootDir = @"C:\printnode";
        public const string JobsDir = RootDir + @"\jobs";

        public static IPrintEngine CreateEngine()
        {
            // 装了 Word 就用真 Word 引擎，否则回退占位引擎
            if (WordComPrintEngine.IsAvailable())
                return new WordComPrintEngine();
            return new WordpadPrintEngine();
        }

        public static void Main()
        {
            Directory.CreateDirectory(JobsDir);
            Directory.CreateDirectory(Path.GetDirectoryName(WordpadPrintEngine.FixedPortFile));
            var engine = CreateEngine();
            var worker = new StaWorker(engine);

            var listener = new HttpListener();
            listener.Prefixes.Add("http://+:8090/");
            listener.Start();
            Console.WriteLine("PrintParseService listening on http://+:8090/ engine=" + engine.Name);

            while (true)
            {
                HttpListenerContext ctx;
                try { ctx = listener.GetContext(); }
                catch { break; }
                ThreadPool.QueueUserWorkItem(delegate { Handle(ctx, worker, engine); });
            }
        }

        private static void Handle(HttpListenerContext ctx, StaWorker worker, IPrintEngine engine)
        {
            try
            {
                string path = ctx.Request.Url.AbsolutePath;
                if (ctx.Request.HttpMethod == "GET" && (path == "/health" || path == "/health/"))
                {
                    StaWorker.RespondJson(ctx, 200, new Dictionary<string, object>
                    {
                        { "status", "ok" },
                        { "engine", engine.Name },
                        { "printer", WordpadPrintEngine.PrinterName },
                        { "time", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") }
                    });
                    return;
                }
                if (ctx.Request.HttpMethod == "GET" && (path == "/" || path == ""))
                {
                    StaWorker.RespondJson(ctx, 200, new Dictionary<string, object>
                    {
                        { "service", "print-parse-node" },
                        { "engine", engine.Name },
                        { "endpoints", new string[] { "GET /health", "POST /api/print-parse (multipart field: file)", "GET /images/{job}/{page}.png" } }
                    });
                    return;
                }
                if (ctx.Request.HttpMethod == "GET" && path.StartsWith("/images/"))
                {
                    ServeImage(ctx, path);
                    return;
                }
                if (ctx.Request.HttpMethod == "POST" && path.StartsWith("/api/print-parse"))
                {
                    var job = SaveUpload(ctx);
                    worker.Enqueue(job);
                    return;
                }
                StaWorker.RespondError(ctx, 404, "not found: " + path);
            }
            catch (Exception ex)
            {
                StaWorker.RespondError(ctx, 500, StaWorker.Diag(ex));
            }
        }

        private static void ServeImage(HttpListenerContext ctx, string path)
        {
            // /images/{job}/{page}.png
            var parts = path.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 3 || parts[1].IndexOf("..") >= 0 || parts[2].IndexOf("..") >= 0)
            { StaWorker.RespondError(ctx, 400, "bad path"); return; }
            string fp = Path.Combine(JobsDir, parts[1] + @"\" + parts[2]);
            if (!File.Exists(fp)) { StaWorker.RespondError(ctx, 404, "image not found"); return; }
            byte[] buf = File.ReadAllBytes(fp);
            ctx.Response.StatusCode = 200;
            ctx.Response.ContentType = "image/png";
            ctx.Response.ContentLength64 = buf.Length;
            ctx.Response.OutputStream.Write(buf, 0, buf.Length);
            ctx.Response.OutputStream.Close();
        }

        // 极简 multipart 解析：取第一个带 filename 的 part
        private static Job SaveUpload(HttpListenerContext ctx)
        {
            string ctype = ctx.Request.ContentType ?? "";
            if (!ctype.StartsWith("multipart/form-data"))
                throw new Exception("expect multipart/form-data");
            string boundary = null;
            foreach (var seg in ctype.Split(';'))
            {
                var s = seg.Trim();
                if (s.StartsWith("boundary=")) boundary = s.Substring("boundary=".Length).Trim('"');
            }
            if (boundary == null) throw new Exception("no multipart boundary");

            byte[] body;
            using (var ms = new MemoryStream())
            {
                ctx.Request.InputStream.CopyTo(ms);
                body = ms.ToArray();
            }
            byte[] bnd = Encoding.ASCII.GetBytes("--" + boundary);
            byte[] hdrEnd = new byte[] { 13, 10, 13, 10 };
            int pos = 0;
            string filename = "upload.bin";
            byte[] content = null;
            while (true)
            {
                int b = IndexOf(body, bnd, pos);
                if (b < 0) break;
                int next = IndexOf(body, bnd, b + bnd.Length);
                if (next < 0) next = body.Length;
                int hs = b + bnd.Length;
                int he = IndexOf(body, hdrEnd, hs);
                if (he > 0 && he < next)
                {
                    string headers = Encoding.UTF8.GetString(body, hs, he - hs);
                    int fi = headers.IndexOf("filename=\"");
                    if (fi >= 0)
                    {
                        int fe = headers.IndexOf('"', fi + 10);
                        filename = headers.Substring(fi + 10, fe - fi - 10);
                        filename = Path.GetFileName(filename);
                        int cs = he + 4;
                        int ce = next - 2; // strip trailing CRLF
                        content = new byte[ce - cs];
                        Array.Copy(body, cs, content, 0, ce - cs);
                        break;
                    }
                }
                pos = next;
            }
            if (content == null) throw new Exception("no file part found");

            // 文件名消毒：扩展名外一律忽略（客户端中文名经 ASCII/UTF-8 头可能含非法字符）
            string ext = ".bin";
            try
            {
                string e = Path.GetExtension(filename);
                if (!string.IsNullOrEmpty(e) && e.Length <= 6 && e.IndexOfAny(Path.GetInvalidPathChars()) < 0)
                    ext = e;
            }
            catch { }

            string id = Guid.NewGuid().ToString("N").Substring(0, 12);
            string jobDir = Path.Combine(JobsDir, id);
            Directory.CreateDirectory(jobDir);
            string inputPath = Path.Combine(jobDir, "input" + ext);
            File.WriteAllBytes(inputPath, content);
            return new Job { Id = id, InputPath = inputPath, OriginalName = filename, JobDir = jobDir, Ctx = ctx };
        }

        private static int IndexOf(byte[] hay, byte[] needle, int start)
        {
            for (int i = start; i <= hay.Length - needle.Length; i++)
            {
                bool ok = true;
                for (int j = 0; j < needle.Length; j++)
                    if (hay[i + j] != needle[j]) { ok = false; break; }
                if (ok) return i;
            }
            return -1;
        }
    }
}
