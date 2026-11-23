using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace AmbientWriterWatcher
{
    /// <summary>
    /// Out-of-process companion for Word: watches the caret, asks an LLM for a
    /// continuation, shows it as grey inline text, Tab accepts / Esc discards.
    /// </summary>
    internal static class Program
    {
        private const string GhostBookmark = "AWB_GHOST";
        private static readonly string Dir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AmbientWriter");
        private static readonly string LogPath = Path.Combine(Dir, "watcher.log");

        private static Word.Application _app;
        private static uint _appPid;
        private static IntPtr _appHwnd = IntPtr.Zero;
        private static bool _ghostActive;
        private static int _ghostStart;
        private static int _ghostEnd;
        private static string _ghostText = "";
        private static string _ghostDoc = "";
        private static int _lastCaret = -1;
        private static DateTime _lastChange = DateTime.MinValue;
        private static DateTime _suppressUntil = DateTime.MinValue;
        private static int _suppressCaret = -1;
        private static volatile bool _busy;
        private static volatile bool _hasResult;
        private static string _resultText;
        private static int _resultCaret;
        private static readonly object _lock = new object();
        private static Settings _cfg = new Settings();
        private static System.Windows.Forms.Timer _timer;
        private static int _pendingAccept;
        private static int _pendingEsc;
        private static int _pendingDiscard;
        private static readonly string BridgePath = Path.Combine(Dir, "bridge.json");
        private static readonly string StatusPath = Path.Combine(Dir, "status.json");
        private static DateTime _bridgeStamp = DateTime.MinValue;
        private static DateTime _lastTickLog = DateTime.MinValue;

        [STAThread]
        private static void Main(string[] args)
        {
            Directory.CreateDirectory(Dir);
            _cfg = Settings.Load();
            Log("start; " + _cfg.Describe());

            Hooks.Install();

            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 250;
            _timer.Tick += Tick;
            _timer.Start();

            // self check: AmbientWriterWatcher.exe --selftest calls the model once and logs it
            if (args != null && Array.IndexOf(args, "--selftest") >= 0)
            {
                PollBridge();
                string reply = null;
                string err = null;
                try { reply = Llm.Complete(_cfg, "今天的会议讨论了三个议题，分别是"); }
                catch (Exception ex) { err = ex.Message; }
                Log(err != null ? "selftest ERR: " + err : "selftest OK -> " + (reply == null ? "<null>" : reply));
                Application.Exit();
                return;
            }

            // scripting hook: AmbientWriterWatcher.exe --toggle flips the switch and exits
            if (args != null && Array.IndexOf(args, "--toggle") >= 0)
            {
                SetEnabled(!_cfg.Enabled);
                Application.Exit();
                return;
            }

            var ctx = new ApplicationContext();
            Application.Run(ctx);
        }

        #region bridge (shared config with the GPT Plus task pane)

        private static void PollBridge()
        {
            try
            {
                if (!File.Exists(BridgePath)) return;
                DateTime stamp = File.GetLastWriteTimeUtc(BridgePath);
                if (stamp == _bridgeStamp) return;
                _bridgeStamp = stamp;

                var d = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(
                    File.ReadAllText(BridgePath, Encoding.UTF8));
                if (d == null) return;

                object v;
                if (d.TryGetValue("enabled", out v) && v is bool) _cfg.Enabled = (bool)v;
                if (d.TryGetValue("apiKey", out v) && v is string) _cfg.ApiKey = (string)v;
                if (d.TryGetValue("baseUrl", out v) && v is string && ((string)v).Length > 0) _cfg.BaseUrl = (string)v;
                if (d.TryGetValue("model", out v) && v is string && ((string)v).Length > 0) _cfg.Model = (string)v;
                if (d.TryGetValue("idleMs", out v) && v is int && (int)v > 0) _cfg.IdleMs = (int)v;

                if (d.TryGetValue("quit", out v) && v is bool && (bool)v)
                {
                    HandleQuit();
                    return;
                }

                WriteStatus("");
                Log("bridge applied: " + _cfg.Describe());
            }
            catch (Exception ex) { Log("bridge: " + ex.Message); }
        }

        private static void WriteBridgeEnabled(bool on)
        {
            try
            {
                var d = new Dictionary<string, object>();
                if (File.Exists(BridgePath))
                {
                    var cur = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(
                        File.ReadAllText(BridgePath, Encoding.UTF8));
                    if (cur != null) d = cur;
                }
                d["enabled"] = on;
                Directory.CreateDirectory(Dir);
                File.WriteAllText(BridgePath, new JavaScriptSerializer().Serialize(d), new UTF8Encoding(false));
                _bridgeStamp = File.GetLastWriteTimeUtc(BridgePath);
            }
            catch (Exception ex) { Log("bridge write: " + ex.Message); }
        }


        #endregion

        #region control surface (no tray: the task pane drives everything)

        private static void SetEnabled(bool on)
        {
            try
            {
                _cfg.Enabled = on;
                if (!on) DiscardGhost();
                PersistEnabled(on);
                WriteBridgeEnabled(on);
                WriteStatus("");
                Log("toggled enabled=" + on);
            }
            catch (Exception ex) { Log("toggle failed: " + ex.Message); }
        }

        private static void PersistEnabled(bool on)
        {
            try
            {
                string path = Path.Combine(Dir, "config.ini");
                var lines = File.Exists(path)
                    ? new List<string>(File.ReadAllLines(path, Encoding.UTF8))
                    : new List<string>();
                bool found = false;
                for (int i = 0; i < lines.Count; i++)
                {
                    if (lines[i].TrimStart().StartsWith("enabled", StringComparison.OrdinalIgnoreCase))
                    {
                        lines[i] = "enabled=" + (on ? "1" : "0");
                        found = true;
                    }
                }
                if (!found) lines.Insert(0, "enabled=" + (on ? "1" : "0"));
                File.WriteAllLines(path, lines.ToArray(), Encoding.UTF8);
            }
            catch (Exception ex) { Log("persist failed: " + ex.Message); }
        }

        // Surface the last failure to the task pane instead of a tray balloon.
        internal static void WriteStatus(string error)
        {
            try
            {
                var d = new Dictionary<string, object>();
                d["enabled"] = _cfg.Enabled;
                d["model"] = _cfg.Model;
                d["error"] = error ?? "";
                d["at"] = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                File.WriteAllText(StatusPath, new JavaScriptSerializer().Serialize(d), new UTF8Encoding(false));
            }
            catch { }
        }

        private static void HandleQuit()
        {
            try
            {
                var d = new Dictionary<string, object>();
                if (File.Exists(BridgePath))
                {
                    var cur = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(
                        File.ReadAllText(BridgePath, Encoding.UTF8));
                    if (cur != null) d = cur;
                }
                d.Remove("quit");
                File.WriteAllText(BridgePath, new JavaScriptSerializer().Serialize(d), new UTF8Encoding(false));
                _bridgeStamp = File.GetLastWriteTimeUtc(BridgePath);
                Log("quit requested from panel");
                Application.Exit();
            }
            catch { Application.Exit(); }
        }
        #endregion

        private static void Tick(object sender, EventArgs e)
        {
            try
            {
                PollBridge();

                if (Interlocked.Exchange(ref _pendingAccept, 0) == 1) { AcceptGhost(); return; }
                if (Interlocked.Exchange(ref _pendingEsc, 0) == 1) { DiscardAndSuppress(); return; }
                if (Interlocked.Exchange(ref _pendingDiscard, 0) == 1) { DiscardGhost(); }

                if (_hasResult)
                {
                    string t; int caret;
                    lock (_lock) { t = _resultText; caret = _resultCaret; _hasResult = false; }
                    _busy = false;
                    if (!string.IsNullOrEmpty(t)) ShowGhost(t, caret);
                    return;
                }

                Word.Application app = App();
                if (app == null) { _ghostActive = false; return; }

                Word.Document doc = app.ActiveDocument;
                if (doc == null) return;

                if (_ghostActive && _ghostDoc != doc.Name)
                {
                    _ghostActive = false;
                    _ghostText = "";
                }

                Word.Selection sel = app.Selection;
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

                if (_busy || !_cfg.Enabled) return;
                if (Hooks.TabPending() != 0 || Hooks.EscPending() != 0) return;
                if (!IsWordForeground()) return;
                if (sel.Type != Word.WdSelectionType.wdSelectionIP) return;
                if (DateTime.Now < _suppressUntil) return;
                if ((DateTime.Now - _lastChange).TotalMilliseconds < _cfg.IdleMs) return;

                int from = Math.Max(0, caretNow - _cfg.ContextChars);
                string ctx = doc.Range(from, caretNow).Text;
                if (string.IsNullOrEmpty(ctx)) return;

                string trimmed = ctx.TrimEnd();
                bool longEnough = (caretNow - from) >= _cfg.MinContextChars;
                bool endsSentence = trimmed.Length >= 2 &&
                    "。！？.!?；;…".IndexOf(trimmed[trimmed.Length - 1]) >= 0;
                if (!longEnough && !endsSentence) return;

                _busy = true;
                _lastChange = DateTime.Now;
                int caretAtRequest = caretNow;
                ThreadPool.QueueUserWorkItem(delegate
                {
                    string text = null;
                    try { text = Llm.Complete(_cfg, ctx); }
                    catch (Exception ex)
                    {
                        Log("llm: " + ex.Message);
                        WriteStatus("模型调用失败：" + ex.Message);
                    }
                    if (string.IsNullOrEmpty(text)) WriteStatus("模型没有返回内容（检查 key / 余额 / 网络）");
                    lock (_lock) { _resultText = text; _resultCaret = caretAtRequest; _hasResult = true; }
                });
            }
            catch (Exception ex)
            {
                if ((DateTime.Now - _lastTickLog).TotalSeconds > 10) { _lastTickLog = DateTime.Now; Log("tick: " + ex.Message); }
            }
        }

        // Bind to the Word instance the user is actually typing in. Word's frame
        // window refuses AccessibleObjectFromWindow, but its "_WwG" child works.
        private static Word.Application App()
        {
            try
            {
                IntPtr fg = Hooks.GetForegroundWindow();
                uint pid;
                Hooks.GetWindowThreadProcessId(fg, out pid);

                if (pid != 0 && IsWinword((int)pid))
                {
                    if (fg != _appHwnd || _app == null)
                    {
                        Word.Application bound = AppFromWindow(fg) ?? AppFromWindow(Hooks.FindChild(fg, "_WwG"));
                        if (bound != null)
                        {
                            _app = bound;
                            _appHwnd = fg;
                            _appPid = pid;
                            Log("bound to word pid=" + pid);
                        }
                    }
                    if (_app != null && HasDocument(_app)) return _app;
                }

                if (_app != null && HasDocument(_app)) return _app;

                try
                {
                    var fallback = (Word.Application)Marshal.GetActiveObject("Word.Application");
                    if (fallback != null && HasDocument(fallback))
                    {
                        _app = fallback;
                        _appHwnd = IntPtr.Zero;
                        return _app;
                    }
                }
                catch { }

                return _app;
            }
            catch { return _app; }
        }

        private static bool IsWinword(int pid)
        {
            try
            {
                return string.Equals(System.Diagnostics.Process.GetProcessById(pid).ProcessName,
                    "WINWORD", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private static bool HasDocument(Word.Application app)
        {
            try { return app.ActiveDocument != null; }
            catch { return false; }
        }
        private static Word.Application AppFromWindow(IntPtr hwnd)
        {
            try
            {
                Guid iid = new Guid("00020400-0000-0000-C000-000000000046");   // IID_IDispatch
                object native;
                int hr = Hooks.AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref iid, out native);
                if (hr != 0 || native == null) return null;

                Word.Window win = native as Word.Window;
                if (win != null) return win.Application;

                Word.Application app = native as Word.Application;
                return app;
            }
            catch { return null; }
        }

        #region ghost

        private static void ShowGhost(string text, int caret)
        {
            try
            {
                Word.Application app = App();
                if (app == null) return;
                Word.Document doc = app.ActiveDocument;
                if (doc == null) return;
                Word.Selection sel = app.Selection;
                if (sel.Start != caret) return;

                string clean = Clean(text);
                if (clean.Length == 0) return;

                Word.Range r = doc.Range(caret, caret);
                r.Text = clean;
                r.Font.Color = (Word.WdColor)0x8C8C8C;
                r.Font.Italic = 0;

                _ghostStart = caret;
                _ghostText = clean;
                _ghostDoc = doc.Name;
                _ghostEnd = r.End;
                _ghostActive = true;
                sel.SetRange(caret, caret);
                Log("ghost shown (" + clean.Length + ")");
            }
            catch (Exception ex) { Log("show: " + ex.Message); }
        }

        public static void AcceptGhost()
        {
            try
            {
                Word.Application app = App();
                if (app == null || !_ghostActive) { _ghostActive = false; return; }
                Word.Document doc = app.ActiveDocument;
                if (doc == null) { _ghostActive = false; return; }

                Word.Range g = SafeGhostRange(doc);
                if (g == null) { _ghostActive = false; return; }

                g.Font.Color = Word.WdColor.wdColorAutomatic;
                int end = g.End;
                _ghostActive = false;
                app.Selection.SetRange(end, end);
                _suppressUntil = DateTime.Now.AddSeconds(1.5);
                Log("accepted");
            }
            catch (Exception ex) { Log("accept: " + ex.Message); }
        }

        public static void DiscardGhost()
        {
            try
            {
                Word.Application app = App();
                if (app == null || !_ghostActive) { _ghostActive = false; return; }
                Word.Document doc = app.ActiveDocument;
                if (doc == null) { _ghostActive = false; return; }

                Word.Range g = SafeGhostRange(doc);
                if (g == null)
                {
                    _ghostActive = false;
                    Log("discard skipped: range no longer matches");
                    return;
                }

                int start = _ghostStart;
                g.Delete();
                _ghostActive = false;
                app.Selection.SetRange(start, start);
                Log("discarded");
            }
            catch (Exception ex) { Log("discard: " + ex.Message); }
        }

        /// <summary>
        /// Returns the range that holds the ghost text, or null when the document
        /// changed underneath us (so we never delete the user's own text).
        /// </summary>
        private static Word.Range SafeGhostRange(Word.Document doc)
        {
            int len = _ghostText.Length;
            if (len <= 0) return null;
            int end = _ghostStart + len;
            if (end > doc.Content.End) return null;

            string seg = doc.Range(_ghostStart, end).Text;
            if (seg.Length < len || !seg.StartsWith(_ghostText, StringComparison.Ordinal)) return null;

            _ghostEnd = end;
            return doc.Range(_ghostStart, end);
        }

        public static void DiscardAndSuppress()
        {
            _suppressCaret = _ghostStart;
            _suppressUntil = DateTime.Now.AddSeconds(3);
            DiscardGhost();
        }

        public static bool GhostActive { get { return _ghostActive; } }

        public static void QueueAccept() { Interlocked.Exchange(ref _pendingAccept, 1); }
        public static void QueueEsc() { Interlocked.Exchange(ref _pendingEsc, 1); }
        public static void QueueDiscard() { Interlocked.Exchange(ref _pendingDiscard, 1); }

        private static string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = s.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Replace("\u0007", "");
            while (s.IndexOf("  ", StringComparison.Ordinal) >= 0) s = s.Replace("  ", " ");
            s = s.Trim();
            if (s.Length > 2 && s.StartsWith("\"", StringComparison.Ordinal) && s.EndsWith("\"", StringComparison.Ordinal))
                s = s.Substring(1, s.Length - 2);
            return s;
        }

        #endregion

        private static bool IsWordForeground()
        {
            IntPtr h = Hooks.GetForegroundWindow();
            if (h == IntPtr.Zero) return false;
            uint pid;
            Hooks.GetWindowThreadProcessId(h, out pid);
            try
            {
                var p = System.Diagnostics.Process.GetProcessById((int)pid);
                return string.Equals(p.ProcessName, "WINWORD", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        internal static void Log(string msg)
        {
            try
            {
                Directory.CreateDirectory(Dir);
                File.AppendAllText(LogPath,
                    DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) + "  " + msg + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch { }
        }
    }

    internal static class Hooks
    {
        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out object ppvObject);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, EnumChildProc cb, IntPtr lparam);
        private delegate bool EnumChildProc(IntPtr h, IntPtr l);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr h, StringBuilder s, int n);

        public static IntPtr FindChild(IntPtr parent, string cls)
        {
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(parent, delegate(IntPtr h, IntPtr l)
            {
                var sb = new StringBuilder(128);
                GetClassName(h, sb, 128);
                if (sb.ToString() == cls) { found = h; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int VK_TAB = 0x09;
        private const int VK_ESCAPE = 0x1B;

        private static HookProc _proc;          // keep alive
        private static IntPtr _hook = IntPtr.Zero;
        private static long _tabAt;
        private static long _escAt;

        public static long TabPending() { return Interlocked.Read(ref _tabAt); }
        public static long EscPending() { return Interlocked.Read(ref _escAt); }

        public static void Install()
        {
            _proc = Callback;
            _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(null), 0);
            Program.Log(_hook == IntPtr.Zero ? "hook failed: " + Marshal.GetLastWin32Error() : "hook installed");
        }

        private static IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vk = Marshal.ReadInt32(lParam);
                if (Program.GhostActive)
                {
                    if (vk == VK_TAB)
                    {
                        Program.QueueAccept();
                        return (IntPtr)1;
                    }
                    if (vk == VK_ESCAPE)
                    {
                        Program.QueueEsc();
                        return (IntPtr)1;
                    }
                    Program.QueueDiscard();
                }
            }
            return CallNextHookEx(_hook, nCode, wParam, lParam);
        }
    }

    internal class Settings
    {
        public bool Enabled = true;
        public int IdleMs = 1100;
        public int ContextChars = 400;
        public int MinContextChars = 10;
        public string ApiKey = "";
        public string BaseUrl = "https://api.deepseek.com/v1";
        public string Model = "deepseek-chat";

        public static Settings Load()
        {
            var s = new Settings();
            try
            {
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AmbientWriter", "config.ini");
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
            if (string.IsNullOrEmpty(cfg.ApiKey)) return Mock();

            var payload = new Dictionary<string, object>();
            payload["model"] = cfg.Model;
            payload["temperature"] = 0.8;
            payload["max_tokens"] = 400;
            payload["stream"] = false;
            var messages = new List<object>();
            messages.Add(new Dictionary<string, object> { { "role", "system" }, { "content", SystemPrompt } });
            messages.Add(new Dictionary<string, object> { { "role", "user" }, { "content", context } });
            payload["messages"] = messages;

            byte[] bytes = Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(payload));
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
                string raw = sr.ReadToEnd();
                var obj = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(raw);
                if (obj == null || !obj.ContainsKey("choices"))
                {
                    Program.Log("llm raw: " + (raw.Length > 500 ? raw.Substring(0, 500) : raw));
                    return null;
                }
                object choices;
                if (obj != null && obj.TryGetValue("choices", out choices))
                {
                    object firstItem = null;
                    var al = choices as System.Collections.ArrayList;
                    if (al != null && al.Count > 0) firstItem = al[0];
                    else
                    {
                        var arr = choices as object[];
                        if (arr != null && arr.Length > 0) firstItem = arr[0];
                    }

                    var first = firstItem as Dictionary<string, object>;
                    object message;
                    if (first != null && first.TryGetValue("message", out message))
                    {
                        var m = message as Dictionary<string, object>;
                        object content;
                        if (m != null && m.TryGetValue("content", out content))
                        {
                            string s = content as string;
                            if (!string.IsNullOrEmpty(s)) return s;
                            Program.Log("llm empty content; raw=" + (raw.Length > 500 ? raw.Substring(0, 500) : raw));
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

        private static string Mock()
        {
            return "（模拟续写）这里会出现 AI 生成的下一句，按 Tab 接受，Esc 放弃。";
        }
    }
}
