using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

namespace AmbientWriter
{
    #region COM / Extensibility plumbing

    [ComVisible(true)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, object custom);
        [DispId(2)]
        void OnDisconnection(ext_DisconnectMode RemoveMode, object custom);
        [DispId(3)]
        void OnAddInsUpdate(object custom);
        [DispId(4)]
        void OnStartupComplete(object custom);
        [DispId(5)]
        void OnBeginShutdown(object custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1
    }

    #endregion

    [ComVisible(true)]
    [Guid("7C3E9A21-5B4D-4F0E-9A11-2D6C8B3F5E70")]
    [ProgId("AmbientWriter.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDTExtensibility2))]
    public class Connect : IDTExtensibility2
    {
        private const string GhostBookmark = "AWB_GHOST";
        private const string LogFile = "ambientwriter.log";

        private Word.Application _app;
        private System.Windows.Forms.Timer _timer;
        private IntPtr _hook = IntPtr.Zero;
        private HookProc _hookProc;      // keep alive: GC must not collect the delegate

        private Settings _cfg = new Settings();

        // ghost state
        private bool _ghostActive;
        private int _ghostStart;
        private int _ghostEnd;

        // idle detection
        private int _lastCaret = -1;
        private DateTime _lastChange = DateTime.MinValue;
        private DateTime _suppressUntil = DateTime.MinValue;
        private int _suppressCaret = -1;

        // async suggestion
        private readonly object _lock = new object();
        private volatile bool _hasResult;
        private string _resultText;
        private int _resultCaret;
        private volatile bool _busy;

        public Connect()
        {
            Log("ctor: " + typeof(Connect).Assembly.Location);
        }

        #region IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, object custom)
        {
            try
            {
                Log("OnConnection enter");
                _app = (Word.Application)Application;
                _cfg = Settings.Load();
                _timer = new System.Windows.Forms.Timer();
                _timer.Interval = 250;
                _timer.Tick += Tick;
                _timer.Start();
                InstallHook();
                Log("connected; config=" + _cfg.Describe());
            }
            catch (Exception ex)
            {
                Log("OnConnection failed: " + ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, object custom)
        {
            try
            {
                if (_timer != null) { _timer.Stop(); _timer.Dispose(); _timer = null; }
                UninstallHook();
                DiscardGhost();
                _app = null;
                Log("disconnected");
            }
            catch (Exception ex) { Log("OnDisconnection failed: " + ex); }
        }

        public void OnAddInsUpdate(object custom) { }

        public void OnStartupComplete(object custom) { }

        public void OnBeginShutdown(object custom) { }

        #endregion

        #region idle detect + suggestion lifecycle

        private void Tick(object sender, EventArgs e)
        {
            try
            {
                if (_app == null) return;

                if (_hasResult)
                {
                    string t;
                    int caret;
                    lock (_lock) { t = _resultText; caret = _resultCaret; _hasResult = false; }
                    _busy = false;
                    if (!string.IsNullOrEmpty(t)) ShowGhost(t, caret);
                    return;
                }

                Word.Document doc = _app.ActiveDocument;
                if (doc == null) { DiscardGhost(); return; }

                Word.Selection sel = _app.Selection;
                int caretNow = sel.Start;

                if (_ghostActive)
                {
                    if (caretNow != _ghostStart || sel.Type != Word.WdSelectionType.wdSelectionIP)
                        DiscardGhost();
                    else
                        return;
                }

                if (caretNow != _lastCaret)
                {
                    _lastCaret = caretNow;
                    _lastChange = DateTime.Now;
                    return;
                }

                if (_busy) return;
                if (!_cfg.Enabled) return;
                if (!IsWordForeground()) return;
                if (sel.Type != Word.WdSelectionType.wdSelectionIP) return;
                if (DateTime.Now < _suppressUntil && caretNow == _suppressCaret) return;
                if ((DateTime.Now - _lastChange).TotalMilliseconds < _cfg.IdleMs) return;

                int from = Math.Max(0, caretNow - _cfg.ContextChars);
                if (caretNow - from < _cfg.MinContextChars) return;

                Word.Range ctxRange = doc.Range(from, caretNow);
                string ctx = ctxRange.Text;
                if (string.IsNullOrEmpty(ctx) || ctx.Trim().Length < _cfg.MinContextChars) return;

                _busy = true;
                _lastChange = DateTime.Now;   // avoid re-firing while waiting
                RequestSuggestion(ctx, caretNow);
            }
            catch (Exception ex)
            {
                Log("Tick: " + ex.Message);
            }
        }

        private void RequestSuggestion(string ctx, int caret)
        {
            ThreadPool.QueueUserWorkItem(delegate
            {
                string text = null;
                try { text = Llm.Complete(_cfg, ctx); }
                catch (Exception ex) { Log("llm: " + ex.Message); }
                lock (_lock) { _resultText = text; _resultCaret = caret; _hasResult = true; }
            });
        }

        private void ShowGhost(string text, int caret)
        {
            try
            {
                Word.Document doc = _app.ActiveDocument;
                if (doc == null) return;
                Word.Selection sel = _app.Selection;
                if (sel.Start != caret) return;              // user moved: drop it

                string clean = Clean(text);
                if (clean.Length == 0) return;

                Word.Range r = doc.Range(caret, caret);
                r.Text = clean;
                r.Font.Color = (Word.WdColor)0x8C8C8C;       // gray
                r.Font.Italic = 0;
                doc.Bookmarks.Add(GhostBookmark, r);

                _ghostStart = caret;
                _ghostEnd = r.End;
                _ghostActive = true;

                sel.SetRange(caret, caret);
                Log("ghost shown (" + clean.Length + " chars)");
            }
            catch (Exception ex) { Log("ShowGhost: " + ex.Message); }
        }

        private void AcceptGhost()
        {
            try
            {
                Word.Document doc = _app != null ? _app.ActiveDocument : null;
                if (doc == null || !doc.Bookmarks.Exists(GhostBookmark)) { _ghostActive = false; return; }

                Word.Range g = doc.Bookmarks[GhostBookmark].Range;
                g.Font.Color = Word.WdColor.wdColorAutomatic;
                int end = g.End;
                doc.Bookmarks[GhostBookmark].Delete();
                _ghostActive = false;
                _app.Selection.SetRange(end, end);
                Log("ghost accepted");
            }
            catch (Exception ex) { Log("AcceptGhost: " + ex.Message); }
        }

        private void DiscardGhost()
        {
            try
            {
                if (!_ghostActive) return;
                Word.Document doc = _app != null ? _app.ActiveDocument : null;
                if (doc != null && doc.Bookmarks.Exists(GhostBookmark))
                {
                    Word.Range g = doc.Bookmarks[GhostBookmark].Range;
                    int start = g.Start;
                    g.Delete();
                    if (doc.Bookmarks.Exists(GhostBookmark)) doc.Bookmarks[GhostBookmark].Delete();
                    _app.Selection.SetRange(start, start);
                }
                _ghostActive = false;
                Log("ghost discarded");
            }
            catch (Exception ex) { Log("DiscardGhost: " + ex.Message); }
        }

        private static string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = s.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
            s = s.Replace("\u0007", "");                     // cell marker
            while (s.IndexOf("  ", StringComparison.Ordinal) >= 0) s = s.Replace("  ", " ");
            s = s.Trim();
            if (s.Length > 2 && s.StartsWith("\"", StringComparison.Ordinal) && s.EndsWith("\"", StringComparison.Ordinal))
                s = s.Substring(1, s.Length - 2);
            return s;
        }

        #endregion

        #region keyboard hook

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int VK_TAB = 0x09;
        private const int VK_ESCAPE = 0x1B;

        private void InstallHook()
        {
            if (_hook != IntPtr.Zero) return;
            _hookProc = HookCallback;
            _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, GetModuleHandle(null), 0);
            if (_hook == IntPtr.Zero) Log("hook failed: " + Marshal.GetLastWin32Error());
            else Log("hook installed");
        }

        private void UninstallHook()
        {
            if (_hook != IntPtr.Zero) { UnhookWindowsHookEx(_hook); _hook = IntPtr.Zero; }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vk = Marshal.ReadInt32(lParam);
                if (_ghostActive)
                {
                    if (vk == VK_TAB) { AcceptGhost(); return (IntPtr)1; }
                    if (vk == VK_ESCAPE)
                    {
                        _suppressCaret = _ghostStart;
                        _suppressUntil = DateTime.Now.AddSeconds(3);
                        DiscardGhost();
                        return (IntPtr)1;
                    }
                    // any other key: drop the ghost, let the key through
                    DiscardGhost();
                }
            }
            return CallNextHookEx(_hook, nCode, wParam, lParam);
        }

        private static bool IsWordForeground()
        {
            IntPtr h = GetForegroundWindow();
            if (h == IntPtr.Zero) return false;
            uint pid;
            GetWindowThreadProcessId(h, out pid);
            return pid == (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
        }

        #endregion

        #region logging

        private static void Log(string msg)
        {
            try
            {
                string dir = Settings.Dir;
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, LogFile),
                    DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) + "  " + msg + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch { }
        }

        #endregion
    }

    internal class Settings
    {
        public bool Enabled = true;
        public int IdleMs = 1100;
        public int ContextChars = 400;
        public int MinContextChars = 20;
        public string ApiKey = "";
        public string BaseUrl = "https://api.deepseek.com/v1";
        public string Model = "deepseek-chat";

        public static string Dir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AmbientWriter"); }
        }

        public static Settings Load()
        {
            var s = new Settings();
            try
            {
                string path = Path.Combine(Dir, "config.ini");
                if (!File.Exists(path)) return s;
                foreach (string raw in File.ReadAllLines(path, Encoding.UTF8))
                {
                    string line = raw.Trim();
                    if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;
                    int i = line.IndexOf('=');
                    if (i <= 0) continue;
                    string k = line.Substring(0, i).Trim().ToLowerInvariant();
                    string v = line.Substring(i + 1).Trim();
                    switch (k)
                    {
                        case "enabled": s.Enabled = v != "0" && v.ToLowerInvariant() != "false"; break;
                        case "idlems": s.IdleMs = Int(v, s.IdleMs); break;
                        case "contextchars": s.ContextChars = Int(v, s.ContextChars); break;
                        case "mincontextchars": s.MinContextChars = Int(v, s.MinContextChars); break;
                        case "apikey": s.ApiKey = v; break;
                        case "baseurl": s.BaseUrl = v; break;
                        case "model": s.Model = v; break;
                    }
                }
            }
            catch { }
            return s;
        }

        private static int Int(string v, int fallback)
        {
            int n;
            return int.TryParse(v, out n) && n > 0 ? n : fallback;
        }

        public string Describe()
        {
            return string.Format("enabled={0} idle={1}ms model={2} key={3}", Enabled, IdleMs, Model,
                string.IsNullOrEmpty(ApiKey) ? "<mock>" : "<set>");
        }
    }

    internal static class Llm
    {
        public static string Complete(Settings cfg, string context)
        {
            if (string.IsNullOrEmpty(cfg.ApiKey)) return Mock(context);

            var payload = new Dictionary<string, object>();
            payload["model"] = cfg.Model;
            payload["temperature"] = 0.8;
            payload["max_tokens"] = 120;
            payload["stream"] = false;
            var messages = new List<object>();
            messages.Add(new Dictionary<string, object> { { "role", "system" }, { "content", SystemPrompt } });
            messages.Add(new Dictionary<string, object> { { "role", "user" }, { "content", context } });
            payload["messages"] = messages;

            string body = new JavaScriptSerializer().Serialize(payload);
            byte[] bytes = Encoding.UTF8.GetBytes(body);

            string url = cfg.BaseUrl.TrimEnd('/') + "/chat/completions";
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/json";
            req.Accept = "application/json";
            req.Timeout = 15000;
            req.ReadWriteTimeout = 15000;
            req.Headers["Authorization"] = "Bearer " + cfg.ApiKey;
            req.ContentLength = bytes.Length;

            using (Stream s = req.GetRequestStream()) s.Write(bytes, 0, bytes.Length);
            using (var resp = (HttpWebResponse)req.GetResponse())
            using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
            {
                string json = sr.ReadToEnd();
                var obj = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(json);
                object choices;
                if (obj != null && obj.TryGetValue("choices", out choices))
                {
                    var list = choices as object[];
                    if (list != null && list.Length > 0)
                    {
                        var first = list[0] as Dictionary<string, object>;
                        object message;
                        if (first != null && first.TryGetValue("message", out message))
                        {
                            var m = message as Dictionary<string, object>;
                            object content;
                            if (m != null && m.TryGetValue("content", out content))
                                return content as string;
                        }
                    }
                }
            }
            return null;
        }

        private const string SystemPrompt =
            "你是中文写作助手。根据用户提供的上文，续写紧接着的下一句或下一段。" +
            "要求：风格、人称、术语与上文一致；只输出续写正文本身，" +
            "不要解释、不要加引号、不要重复上文。";

        private static string Mock(string context)
        {
            return "（模拟续写）这里会出现 AI 生成的下一句，按 Tab 接受，Esc 放弃。";
        }
    }
}
