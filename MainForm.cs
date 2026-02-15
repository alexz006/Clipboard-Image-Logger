using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.LinkLabel;

namespace ClipboardImageLogger
{
    public partial class MainForm : Form
    {
        // WinAPI: подписка на события изменения буфера
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetClipboardOwner();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetClipboardSequenceNumber();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetUpdatedClipboardFormats(uint[] lpuiFormats, uint cFormats, out uint pcFormatsOut);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClipboardFormatName(uint format, StringBuilder lpszFormatName, int cchMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_CLIPBOARDUPDATE = 0x031D;

        private const int EM_SETSEL = 0x00B1;
        private const int EM_SCROLLCARET = 0x00B7;

        private readonly TextBox _logBox;
        private readonly string _logPath;

        private bool _scrollToEndOnShown;

        // Menu / settings
        private MenuStrip _menu = null!;
        private ToolStripMenuItem _miSettings = null!;
        private ToolStripMenuItem _miRunAtStartup = null!;
        private ToolStripMenuItem _miMinimizeToTray = null!;
        private ToolStripMenuItem _miStartMinimized = null!;
        private ToolStripMenuItem _miNotifications = null!;
        private ToolStripMenuItem _miSaveScreenshots = null!;
        private ToolStripMenuItem _miHelp = null!;
        private ToolStripMenuItem _miAbout = null!;

        // Tray
        private NotifyIcon _trayIcon = null!;
        private ContextMenuStrip _trayMenu = null!;
        private bool _allowExit;

        // Notification window
        private NotificationForm _notifyForm;

        // Run at startup (registry mechanism)
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValueName = "ClipboardImageLogger";

        // INI settings (stored near exe)
        private const string IniFileName = "ClipboardImageLogger.ini";
        private const string IniRunAtStartup = "RunAtStartup";
        private const string IniMinimizeToTray = "MinimizeToTray";
        private const string IniStartMinimized = "StartMinimized";
        private const string IniNotifications = "Notifications";
        private const string IniSaveScreenshots = "SaveScreenshots";

        private const string IniWinLeft = "WinLeft";
        private const string IniWinTop = "WinTop";
        private const string IniWinWidth = "WinWidth";
        private const string IniWinHeight = "WinHeight";

        // INI storage
        private readonly string _iniPath;
        private readonly Dictionary<string, string> _ini = new(StringComparer.OrdinalIgnoreCase);

        public MainForm()
        {
            Text = "Clipboard Image Logger";
            Width = 700;
            Height = 400;

            // По умолчанию (первый запуск) — по центру
            StartPosition = FormStartPosition.CenterScreen;

            try
            {
                Icon = Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath) ?? Icon;
            }
            catch { }

            _iniPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, IniFileName);
            LoadIni();

            // Если уже сохраняли позицию — откроем там же (и с тем же размером)
            ApplySavedWindowBounds();

            _logBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                HideSelection = true
            };

            _logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClipboardImageLogger.log");

            InitializeMenu();
            InitializeTray();

            Controls.Add(_logBox);
            MainMenuStrip = _menu;
            Controls.Add(_menu);

            LoadOrInitLog();

            LoadSettingsIntoMenu();
            ApplyTrayVisibility();

            // Стартовать свернутым — только если включён трей и включена настройка Start minimized
            if (_miMinimizeToTray.Checked && _miStartMinimized.Checked)
            {
                _trayIcon.Visible = true;

                // чтобы не появлялось в панели задач
                ShowInTaskbar = false;

                WindowState = FormWindowState.Minimized;
                Hide();
            }

            Resize += MainForm_Resize;
            FormClosing += MainForm_FormClosing;

            Shown += (_, __) =>
            {
                if (_scrollToEndOnShown)
                {
                    _scrollToEndOnShown = false;

                    // 1) сразу
                    ScrollLogToEnd();

                    // 2) и ещё раз после первой отрисовки
                    BeginInvoke(new Action(() =>
                    {
                        ScrollLogToEnd();
                    }));
                }
            };
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            if (!AddClipboardFormatListener(this.Handle))
            {
                Log("ERROR: AddClipboardFormatListener failed. Win32=" + Marshal.GetLastWin32Error());
            }
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            RemoveClipboardFormatListener(this.Handle);
            base.OnHandleDestroyed(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE)
            {
                OnClipboardUpdated();
            }

            base.WndProc(ref m);
        }

        private void OnClipboardUpdated()
        {
            uint seq = 0;
            try { seq = GetClipboardSequenceNumber(); } catch { }

            IntPtr ownerHwnd = IntPtr.Zero;
            try { ownerHwnd = GetClipboardOwner(); } catch { }

            IntPtr fgHwnd = IntPtr.Zero;
            try { fgHwnd = GetForegroundWindow(); } catch { }

            //string formatsInfo = GetUpdatedFormatsInfo(); // {formatsInfo} |

            // Снимаем инфо по hwnd один раз (чтобы дальше не было расхождений)
            string ownerHwndInfo = (ownerHwnd == IntPtr.Zero) ? "" : GetHwndInfo(ownerHwnd);
            string fgHwndInfo = (fgHwnd == IntPtr.Zero) ? "" : GetHwndInfo(fgHwnd);

            // Готовим “шапку”, но НЕ пишем её сразу
            string ownerInfo = string.IsNullOrWhiteSpace(ownerHwndInfo) ? "unknown" : $"[{ownerHwndInfo}]";
            string header =
                (ownerHwnd == fgHwnd)
                    ? $"CLIPBOARD | seq={seq} | owner={ownerInfo}"
                    : $"CLIPBOARD | seq={seq} | owner={ownerInfo} | foreground=[{fgHwndInfo}]";

            // Ретраи: буфер часто занят сразу после обновления
            for (int attempt = 1; attempt <= 5; attempt++)
            {
                try
                {
                    if (Clipboard.ContainsImage())
                    {
                        using var img = Clipboard.GetImage();

                        if (img != null)
                        {
                            Log($"{header} | IMAGE | attempt={attempt} | size={img.Width}x{img.Height}");

                            string proc = "";

                            if (ownerHwnd != IntPtr.Zero)
                            {
                                proc = ExtractProcFromHwndInfo(ownerHwndInfo);
                                if (string.IsNullOrWhiteSpace(proc))
                                    proc = GetProcFromHwnd(ownerHwnd);
                            }
                            else
                            {
                                proc = ExtractProcFromHwndInfo(fgHwndInfo);
                                if (string.IsNullOrWhiteSpace(proc))
                                    proc = GetProcFromHwnd(fgHwnd);
                            }

                            if (string.IsNullOrWhiteSpace(proc)) proc = "unknown";

                            // сохранение (если включено)
                            if (GetSaveScreenshotsSetting())
                            {
                                SaveScreenshot(img, proc);
                            }

                            if (GetNotificationsSetting())
                            {
                                if (proc != "punto")
                                {
                                    ShowOrUpdateNotification(proc);
                                }

                            }

                            return;
                        }

                        // Если ContainsImage=true, но картинку не получили — ничего не пишем
                        return;
                    }

                    // Не картинка — ничего не пишем
                    return;
                }
                catch (Exception)
                {
                    Thread.Sleep(30);
                }
            }
        }

        private string GetProcFromHwnd(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return "";

            try
            {
                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                if (pid == 0) return "";

                using var p = Process.GetProcessById((int)pid);
                return p.ProcessName ?? "";
            }
            catch
            {
                return "";
            }
        }

        private string ExtractProcFromHwndInfo(string hwndInfo)
        {
            if (string.IsNullOrWhiteSpace(hwndInfo)) return "";

            int i = hwndInfo.IndexOf("proc=", StringComparison.Ordinal);
            if (i < 0) return "";

            i += "proc=".Length;

            int j = hwndInfo.IndexOf(' ', i);
            if (j < 0) j = hwndInfo.Length;

            if (j <= i) return "";
            return hwndInfo.Substring(i, j - i).Trim();
        }

        private void ShowOrUpdateNotification(string proc)
        {
            if (_notifyForm == null || _notifyForm.IsDisposed)
                _notifyForm = new NotificationForm(this);

            _notifyForm.UpdateProc(proc);
            _notifyForm.ShowNotification();
        }

        // вызывается из NotificationForm при клике
        internal void OpenFromNotification()
        {
            RestoreFromTray();
            try
            {
                _notifyForm?.HideNotification();
            }
            catch { }
        }

        private string GetHwndInfo(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                return "hwnd=0x0";

            var title = new StringBuilder(256);
            GetWindowText(hwnd, title, title.Capacity);

            var cls = new StringBuilder(256);
            GetClassName(hwnd, cls, cls.Capacity);

            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);

            string pname = "";
            string path = "";
            if (pid != 0)
            {
                try
                {
                    using var p = Process.GetProcessById((int)pid);
                    pname = p.ProcessName;
                    try { path = p.MainModule?.FileName ?? ""; } catch { }
                }
                catch { }
            }

            if (!string.IsNullOrWhiteSpace(path))
            {
                return
                    $"hwnd=0x{hwnd.ToInt64():X} " +
                    $"pid={pid} " +
                    $"proc={pname} " +
                    $"title=\"{title}\" " +
                    $"class=\"{cls}\" " +
                    $"path={path}";
            }

            if (pid != 0 || !string.IsNullOrWhiteSpace(pname))
            {
                return
                    $"hwnd=0x{hwnd.ToInt64():X} " +
                    $"pid={pid} " +
                    $"proc={pname} " +
                    $"title=\"{title}\" " +
                    $"class=\"{cls}\"";
            }

            return
                $"hwnd=0x{hwnd.ToInt64():X} " +
                $"title=\"{title}\" " +
                $"class=\"{cls}\"";
        }

        private string GetUpdatedFormatsInfo()
        {
            try
            {
                // Обычно хватает 64, если нет — можно увеличить
                uint[] fmts = new uint[64];
                uint outCount;
                if (!GetUpdatedClipboardFormats(fmts, (uint)fmts.Length, out outCount))
                    return "formats=GetUpdatedClipboardFormats_failed";

                if (outCount == 0)
                    return "formats=[]";

                var sb = new StringBuilder();
                sb.Append("formats=[");
                for (int i = 0; i < outCount; i++)
                {
                    if (i > 0) sb.Append(", ");
                    sb.Append(FormatName(fmts[i]));
                }
                sb.Append("]");
                return sb.ToString();
            }
            catch
            {
                return "formats=error";
            }
        }

        private string FormatName(uint fmt)
        {
            switch (fmt)
            {
                case 2: return "CF_BITMAP(2)";
                case 8: return "CF_DIB(8)";
                case 14: return "CF_ENHMETAFILE(14)";
                case 15: return "CF_HDROP(15)";
                case 17: return "CF_DIBV5(17)";
            }

            // Для зарегистрированных/кастомных форматов попробуем имя
            var name = new StringBuilder(256);
            int len = GetClipboardFormatName(fmt, name, name.Capacity);
            if (len > 0) return $"{name}({fmt})";

            return $"FMT({fmt})";
        }

        private void SaveScreenshot(System.Drawing.Image img, string proc)
        {
            try
            {
                if (img == null) return;

                if (string.IsNullOrWhiteSpace(proc))
                    proc = "unknown";

                // подчистим proc, чтобы было валидно как имя файла
                foreach (char c in Path.GetInvalidFileNameChars())
                    proc = proc.Replace(c, '_');

                string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"image_{ts}_{proc}.png";
                string fullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

                img.Save(fullPath, ImageFormat.Png);
            }
            catch
            { }
        }

        private void Log(string line)
        {
            string msg = $"{line}";

            if (line.IndexOf("Listening clipboard updates...", StringComparison.Ordinal) < 0)
                msg = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {msg}";

            _logBox.AppendText(msg + Environment.NewLine + Environment.NewLine);

            try
            {
                File.AppendAllText(_logPath, msg + Environment.NewLine + Environment.NewLine);
            }
            catch
            { }
        }

        private void ScrollLogToEnd()
        {
            try
            {
                if (_logBox == null) return;
                if (!_logBox.IsHandleCreated) return;

                // каретка в конец и прокрутка
                SendMessage(_logBox.Handle, EM_SETSEL, new IntPtr(-1), new IntPtr(-1));
                SendMessage(_logBox.Handle, EM_SCROLLCARET, IntPtr.Zero, IntPtr.Zero);

                // снять выделение (важно при старте в трее)
                _logBox.SelectionStart = _logBox.TextLength;
                _logBox.SelectionLength = 0;
            }
            catch { }
        }

        private void LoadOrInitLog()
        {
            // если лог есть — читаем и выводим в форму
            try
            {
                if (File.Exists(_logPath))
                {
                    _logBox.Text = File.ReadAllText(_logPath);

                    // Проскроллим после показа формы (в Shown)
                    _scrollToEndOnShown = true;

                    return;
                }
            }
            catch
            {
                // если чтение не удалось — попробуем инициализировать
            }

            // если лога нет (или не прочитали) — пишем стандартный стартовый лог
            try
            {
                _logBox.Clear();
            }
            catch { }

            try
            {
                File.WriteAllText(_logPath, string.Empty);
            }
            catch { }

            Log($"Log file: {_logPath}" + Environment.NewLine + "Listening clipboard updates...");
        }

        private void ClearLogAndWriteHeader()
        {
            try
            {
                _logBox.Clear();
            }
            catch { }

            try
            {
                File.WriteAllText(_logPath, string.Empty);
            }
            catch { }

            Log($"Log file: {_logPath}" + Environment.NewLine + "Listening clipboard updates...");
        }

        // ===== Menu / Settings / Tray =====

        private void InitializeMenu()
        {
            _menu = new MenuStrip();

            var miFile = new ToolStripMenuItem("File");

            var miClearLog = new ToolStripMenuItem("Clear log");
            miClearLog.Click += (_, __) => ClearLogAndWriteHeader();

            var miExit = new ToolStripMenuItem("Exit");
            miExit.Click += (_, __) =>
            {
                _allowExit = true;
                Close();
            };

            miFile.DropDownItems.Add(miClearLog);
            miFile.DropDownItems.Add(new ToolStripSeparator());
            miFile.DropDownItems.Add(miExit);

            _miSettings = new ToolStripMenuItem("Settings");
            _miRunAtStartup = new ToolStripMenuItem("Run at startup") { CheckOnClick = true };
            _miMinimizeToTray = new ToolStripMenuItem("Minimize to tray") { CheckOnClick = true };
            _miStartMinimized = new ToolStripMenuItem("Start minimized") { CheckOnClick = true };
            _miNotifications = new ToolStripMenuItem("Notifications") { CheckOnClick = true };
            _miSaveScreenshots = new ToolStripMenuItem("Save screenshots") { CheckOnClick = true };

            _miRunAtStartup.Click += (_, __) =>
            {
                SetRunAtStartupSetting(_miRunAtStartup.Checked);
                SetRunAtStartup(_miRunAtStartup.Checked);

                // галочку делаем по фактическому состоянию
                _miRunAtStartup.Checked = IsRunAtStartupEnabled();
                SetRunAtStartupSetting(_miRunAtStartup.Checked);
            };

            _miMinimizeToTray.Click += (_, __) =>
            {
                SetMinimizeToTraySetting(_miMinimizeToTray.Checked);

                // Start minimized активен только если включён Minimize to tray
                _miStartMinimized.Enabled = _miMinimizeToTray.Checked;

                ApplyTrayVisibility();
            };

            _miStartMinimized.Click += (_, __) =>
            {
                SetStartMinimizedSetting(_miStartMinimized.Checked);
            };

            _miNotifications.Click += (_, __) =>
            {
                SetNotificationsSetting(_miNotifications.Checked);
            };

            _miSaveScreenshots.Click += (_, __) =>
            {
                SetSaveScreenshotsSetting(_miSaveScreenshots.Checked);
            };

            _miSettings.DropDownItems.Add(_miRunAtStartup);
            _miSettings.DropDownItems.Add(_miMinimizeToTray);
            _miSettings.DropDownItems.Add(_miStartMinimized);
            _miSettings.DropDownItems.Add(_miNotifications);
            _miSettings.DropDownItems.Add(_miSaveScreenshots);

            _miHelp = new ToolStripMenuItem("Help");
            _miAbout = new ToolStripMenuItem("About");
            _miAbout.Click += (_, __) =>
            {
                var ver = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown";

                MessageBox.Show(
                    $"Clipboard Image Logger v{ver}\nLogs clipboard image events (screenshots) with process/window context.",
                    "About",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            };

            _miHelp.DropDownItems.Add(_miAbout);

            _menu.Items.Add(miFile);
            _menu.Items.Add(_miSettings);
            _menu.Items.Add(_miHelp);
        }

        private void InitializeTray()
        {
            _trayMenu = new ContextMenuStrip();
            var miRestore = new ToolStripMenuItem("Restore");
            var miExit = new ToolStripMenuItem("Exit");

            miRestore.Click += (_, __) => RestoreFromTray();
            miExit.Click += (_, __) =>
            {
                _allowExit = true;
                Close();
            };

            _trayMenu.Items.Add(miRestore);
            _trayMenu.Items.Add(miExit);

            Icon exeIcon = SystemIcons.Application;
            try
            {
                exeIcon = Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath) ?? SystemIcons.Application;
            }
            catch { }

            _trayIcon = new NotifyIcon
            {
                Text = "Clipboard Image Logger",
                Visible = false,
                ContextMenuStrip = _trayMenu,
                Icon = exeIcon
            };

            _trayIcon.DoubleClick += (_, __) => RestoreFromTray();
        }

        private void LoadSettingsIntoMenu()
        {
            _miRunAtStartup.Checked = GetRunAtStartupSetting();
            _miMinimizeToTray.Checked = GetMinimizeToTraySetting();

            _miStartMinimized.Checked = GetStartMinimizedSetting();
            _miStartMinimized.Enabled = _miMinimizeToTray.Checked;

            _miNotifications.Checked = GetNotificationsSetting();
            _miSaveScreenshots.Checked = GetSaveScreenshotsSetting();

            // синхронизируем реестр с ini
            SetRunAtStartup(_miRunAtStartup.Checked);
        }

        private void ApplyTrayVisibility()
        {
            bool trayEnabled = _miMinimizeToTray.Checked;

            // Иконка в трее видна ТОЛЬКО когда окно реально скрыто
            _trayIcon.Visible = false;

            // Если выключили "Minimize to tray", но окно скрыто — вернём его
            if (!trayEnabled && !Visible)
            {
                Show();
                WindowState = FormWindowState.Normal;
                Activate();
            }
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (_miMinimizeToTray.Checked && WindowState == FormWindowState.Minimized)
            {
                _trayIcon.Visible = true;

                // чтобы не оставалось в панели задач
                ShowInTaskbar = false;

                Hide();
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_miMinimizeToTray.Checked && !_allowExit)
            {
                // Вместо закрытия — уходим в трей
                e.Cancel = true;
                WindowState = FormWindowState.Minimized;
                _trayIcon.Visible = true;

                // чтобы не оставалось в панели задач
                ShowInTaskbar = false;

                Hide();
            }
            else
            {
                // сохраняем позицию/размер
                SaveWindowBounds();

                // гасим иконку
                if (_trayIcon != null)
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                }
            }
        }

        internal void RestoreFromTray()
        {
            // вернуть в панель задач
            ShowInTaskbar = true;

            Show();
            WindowState = FormWindowState.Normal;
            Activate();
            _trayIcon.Visible = false;

            BeginInvoke(new Action(() =>
            {
                ScrollLogToEnd();

                // повтор после первого layout
                BeginInvoke(new Action(() =>
                {
                    ScrollLogToEnd();
                }));
            }));
        }

        private bool GetRunAtStartupSetting()
        {
            return IniGetBool(IniRunAtStartup, false);
        }

        private void SetRunAtStartupSetting(bool enable)
        {
            IniSetBool(IniRunAtStartup, enable);
            SaveIni();
        }

        private bool IsRunAtStartupEnabled()
        {
            try
            {
                using var rk = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
                var val = rk?.GetValue(RunValueName) as string;
                if (string.IsNullOrWhiteSpace(val)) return false;

                string exe = System.Windows.Forms.Application.ExecutablePath;

                // убираем кавычки и пробелы по краям
                val = val.Trim().Trim('"');

                return val.IndexOf(exe, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private void SetRunAtStartup(bool enable)
        {
            try
            {
                using var rk = Registry.CurrentUser.OpenSubKey(RunKeyPath, true);
                if (rk == null) return;

                if (enable)
                {
                    string exe = System.Windows.Forms.Application.ExecutablePath;
                    rk.SetValue(RunValueName, $"\"{exe}\"");
                }
                else
                {
                    rk.DeleteValue(RunValueName, false);
                }
            }
            catch
            { }
        }

        private void ApplySavedWindowBounds()
        {
            try
            {
                int left = IniGetInt(IniWinLeft, int.MinValue);
                int top = IniGetInt(IniWinTop, int.MinValue);
                int width = IniGetInt(IniWinWidth, -1);
                int height = IniGetInt(IniWinHeight, -1);

                if (left == int.MinValue || top == int.MinValue) return;
                if (width <= 0 || height <= 0) return;

                StartPosition = FormStartPosition.Manual;
                DesktopBounds = new System.Drawing.Rectangle(left, top, width, height);
            }
            catch
            { }
        }

        private void SaveWindowBounds()
        {
            try
            {
                var r = (WindowState == FormWindowState.Normal) ? Bounds : RestoreBounds;

                IniSetInt(IniWinLeft, r.Left);
                IniSetInt(IniWinTop, r.Top);
                IniSetInt(IniWinWidth, r.Width);
                IniSetInt(IniWinHeight, r.Height);

                SaveIni();
            }
            catch
            { }
        }

        private bool GetStartMinimizedSetting()
        {
            return IniGetBool(IniStartMinimized, false);
        }

        private void SetStartMinimizedSetting(bool enable)
        {
            IniSetBool(IniStartMinimized, enable);
            SaveIni();
        }

        private bool GetNotificationsSetting()
        {
            return IniGetBool(IniNotifications, false);
        }

        private void SetNotificationsSetting(bool enable)
        {
            IniSetBool(IniNotifications, enable);
            SaveIni();
        }

        private bool GetSaveScreenshotsSetting()
        {
            return IniGetBool(IniSaveScreenshots, false);
        }

        private void SetSaveScreenshotsSetting(bool enable)
        {
            IniSetBool(IniSaveScreenshots, enable);
            SaveIni();
        }

        private bool GetMinimizeToTraySetting()
        {
            return IniGetBool(IniMinimizeToTray, false);
        }

        private void SetMinimizeToTraySetting(bool enable)
        {
            IniSetBool(IniMinimizeToTray, enable);
            SaveIni();
        }

        private void LoadIni()
        {
            _ini.Clear();

            try
            {
                if (!File.Exists(_iniPath))
                    return;

                foreach (var raw in File.ReadAllLines(_iniPath))
                {
                    var line = raw.Trim();
                    if (line.Length == 0) continue;
                    if (line.StartsWith("#") || line.StartsWith(";")) continue;

                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;

                    var key = line.Substring(0, eq).Trim();
                    var val = line.Substring(eq + 1).Trim();

                    if (key.Length == 0) continue;
                    _ini[key] = val;
                }
            }
            catch
            { }
        }

        private void SaveIni()
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("; ClipboardImageLogger settings");
                foreach (var kv in _ini)
                    sb.AppendLine($"{kv.Key}={kv.Value}");

                File.WriteAllText(_iniPath, sb.ToString(), Encoding.UTF8);
            }
            catch
            { }
        }

        private bool IniGetBool(string key, bool def)
        {
            if (_ini.TryGetValue(key, out var s))
            {
                if (bool.TryParse(s, out var b)) return b;
                if (int.TryParse(s, out var i)) return i != 0;
            }
            return def;
        }

        private int IniGetInt(string key, int def)
        {
            if (_ini.TryGetValue(key, out var s))
            {
                if (int.TryParse(s, out var i)) return i;
            }
            return def;
        }

        private void IniSetBool(string key, bool val)
        {
            _ini[key] = val ? "1" : "0";
        }

        private void IniSetInt(string key, int val)
        {
            _ini[key] = val.ToString();
        }

    }

    internal sealed class NotificationForm : Form
    {
        private readonly MainForm _main;
        private readonly Panel _content;
        private readonly Label _lbl;
        private readonly Button _btnClose;

        private readonly System.Windows.Forms.Timer _mousePollTimer;
        private readonly System.Windows.Forms.Timer _autoHideTimer;

        private Point _initialMouse;
        private Point _lastMouse;
        private bool _mouseMovedAfterShow;

        private bool _autoHideArmed;

        public NotificationForm(MainForm main)
        {
            _main = main;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;

            Width = 320;
            Height = 80;

            // рамка
            BackColor = SystemColors.ControlDark;
            Padding = new Padding(1);

            _content = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Window
            };
            Controls.Add(_content);

            _lbl = new Label
            {
                AutoSize = false,
                Left = 12,
                Top = 12,
                Width = 320 - 12 - 12 - 28,
                Height = 80 - 24,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font(SystemFonts.MessageBoxFont!.FontFamily, 11f, FontStyle.Regular)
            };

            _btnClose = new Button
            {
                Text = "×",
                Width = 24,
                Height = 24,
                Left = Width - 24 - 8,
                Top = 8,
                FlatStyle = FlatStyle.Flat,
                TabStop = false
            };
            _btnClose.FlatAppearance.BorderSize = 0;

            _content.Controls.Add(_lbl);
            _content.Controls.Add(_btnClose);

            _btnClose.Click += (_, __) => HideNotification();

            // клик по окну/тексту открывает главное и скрывает уведомление
            Click += (_, __) => _main.OpenFromNotification();
            _content.Click += (_, __) => _main.OpenFromNotification();
            _lbl.Click += (_, __) => _main.OpenFromNotification();

            _mousePollTimer = new System.Windows.Forms.Timer { Interval = 100 };
            _mousePollTimer.Tick += (_, __) =>
            {
                var p = Cursor.Position;

                // фиксируем первое движение мыши после показа/обновления уведомления
                if (!_mouseMovedAfterShow && p != _initialMouse)
                {
                    _mouseMovedAfterShow = true;
                }

                if (_mouseMovedAfterShow && Visible)
                {
                    bool over = Bounds.Contains(p);

                    if (over)
                    {
                        // курсор над уведомлением: таймер не считает и сбрасывается
                        if (_autoHideTimer!.Enabled)
                            _autoHideTimer!.Stop();

                        _autoHideArmed = false;
                    }
                    else
                    {
                        // курсор вне уведомления: запускаем таймер только один раз (не перезапускаем каждый тик)
                        if (!_autoHideArmed)
                        {
                            _autoHideArmed = true;
                            _autoHideTimer!.Stop();
                            _autoHideTimer!.Interval = 5000;
                            _autoHideTimer!.Start();
                        }
                    }
                }

                _lastMouse = p;
            };

            _autoHideTimer = new System.Windows.Forms.Timer { Interval = 5000 };
            _autoHideTimer.Tick += (_, __) =>
            {
                _autoHideTimer.Stop();
                HideNotification();
            };
        }

        // не забирать фокус
        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_EX_TOOLWINDOW = 0x00000080;
                const int WS_EX_NOACTIVATE = 0x08000000;

                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
                return cp;
            }
        }

        public void UpdateProc(string proc)
        {
            _lbl.Text = $"Clipboard Image:" + Environment.NewLine + $"{proc}";

            // если уведомление уже показывается — считаем это "обновлением появления"
            // и логика автоскрытия начинается заново: ждём движения мыши
            ResetAutoHide();
        }

        public void ShowNotification()
        {
            PositionBottomRight();

            if (!Visible)
                Show();

            // снова убедимся что оно сверху (без активации)
            TopMost = true;

            ResetAutoHide();
        }

        public void HideNotification()
        {
            _autoHideTimer.Stop();
            _mousePollTimer.Stop();

            if (Visible)
                Hide();
        }

        private void ResetAutoHide()
        {
            _autoHideTimer.Stop();
            _mousePollTimer.Stop();

            _mouseMovedAfterShow = false;
            _autoHideArmed = false;
            _initialMouse = Cursor.Position;
            _lastMouse = _initialMouse;

            // ждать движения мыши и управлять автоскрытием/hover
            _mousePollTimer.Start();
        }

        private void PositionBottomRight()
        {
            var wa = Screen.PrimaryScreen!.WorkingArea;
            int margin = 12;

            Left = wa.Right - Width - margin;
            Top = wa.Bottom - Height - margin;
        }
    }
}
