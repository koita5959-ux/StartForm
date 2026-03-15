using System.Diagnostics;
using System.Management;
using System.Text;
using StartForm.Helpers;
using StartForm.Models;

namespace StartForm.Services
{
    public static class ScreenCapturer
    {
        internal static readonly string[] ExcludedProcesses =
        {
            // "explorer" は除外しない → フォルダウィンドウを取込対象にする
            "StartForm",
            "SystemSettings",
            "TextInputHost",
            "ApplicationFrameHost",
            "ShellExperienceHost",
            "SearchHost",
            "StartMenuExperienceHost",
            "LockApp",
            "CompPkgSrv",
            "RuntimeBroker",
            "svchost",
            "dwm",
            "csrss",
            "WinStore.App"
        };

        private static readonly HashSet<string> ChromeLikeBrowsers = new(StringComparer.OrdinalIgnoreCase)
        {
            "chrome", "msedge"
        };

        private static readonly HashSet<string> AllBrowsers = new(StringComparer.OrdinalIgnoreCase)
        {
            "chrome", "msedge", "firefox", "opera", "brave", "vivaldi"
        };

        private static bool ShouldIncludeInvisibleWindow(string processName)
        {
            return processName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                   processName.Equals("Teams", StringComparison.OrdinalIgnoreCase) ||
                   processName.Equals("soffice", StringComparison.OrdinalIgnoreCase);
        }

        private static string? GetProcessCommandLine(Process process)
        {
            try
            {
                using var searcher = new ManagementObjectSearcher(
                    $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {process.Id}");

                foreach (ManagementObject obj in searcher.Get())
                {
                    return obj["CommandLine"]?.ToString();
                }
            }
            catch
            {
            }

            return null;
        }

        private static List<string> SplitCommandLineArguments(string commandLine)
        {
            var args = new List<string>();
            if (string.IsNullOrWhiteSpace(commandLine))
                return args;

            var current = new StringBuilder();
            bool inQuotes = false;

            foreach (char c in commandLine)
            {
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (char.IsWhiteSpace(c) && !inQuotes)
                {
                    if (current.Length > 0)
                    {
                        args.Add(current.ToString());
                        current.Clear();
                    }

                    continue;
                }

                current.Append(c);
            }

            if (current.Length > 0)
                args.Add(current.ToString());

            return args;
        }

        private static bool IsLibreOfficeDocumentPath(string arg)
        {
            if (string.IsNullOrWhiteSpace(arg))
                return false;

            if (!Path.IsPathRooted(arg))
                return false;

            string extension = Path.GetExtension(arg);
            if (string.IsNullOrWhiteSpace(extension))
                return false;

            return extension.Equals(".odt", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".ods", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".odp", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".odg", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".odf", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".doc", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".docx", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".ppt", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase);
        }

        private static string? GetLibreOfficeDocumentPath(Process process)
        {
            var commandLine = GetProcessCommandLine(process);
            if (string.IsNullOrWhiteSpace(commandLine))
                return null;

            return GetLibreOfficeDocumentPathFromCommandLine(commandLine);
        }

        private static string? GetLibreOfficeDocumentPathFromCommandLine(string commandLine)
        {
            if (string.IsNullOrWhiteSpace(commandLine))
                return null;

            var args = SplitCommandLineArguments(commandLine);
            foreach (var arg in args.Skip(1))
            {
                if (arg.StartsWith("-", StringComparison.Ordinal))
                    continue;

                if (IsLibreOfficeDocumentPath(arg))
                    return arg;
            }

            return null;
        }

        private static bool IsTeamsProcess(string processName)
        {
            return processName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                   processName.Equals("Teams", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsLibreOfficeProcess(string processName)
        {
            return processName.Equals("soffice", StringComparison.OrdinalIgnoreCase) ||
                   processName.Equals("soffice.bin", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTeamsAuxiliaryWindowTitle(string title)
        {
            return title.Equals("Default IME", StringComparison.OrdinalIgnoreCase) ||
                   title.Equals("DDE Server Window", StringComparison.OrdinalIgnoreCase) ||
                   title.Equals("MSCTFIME UI", StringComparison.OrdinalIgnoreCase) ||
                   title.Equals("Rtc Video PnP Listener", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsMeaningfulTeamsWindow(IntPtr hWnd, string title, string className)
        {
            if (WindowHelper.IsCloaked(hWnd))
                return false;

            var stableBounds = WindowHelper.GetStableWindowBounds(hWnd);
            int width = Math.Abs(stableBounds.Width);
            int height = Math.Abs(stableBounds.Height);

            if (width < 300 || height < 200)
                return false;

            if (IsTeamsAuxiliaryWindowTitle(title))
                return false;

            if (className.Equals("TeamsMessageWindow", StringComparison.OrdinalIgnoreCase))
                return false;

            if (string.IsNullOrWhiteSpace(title))
            {
                return className.Equals("TeamsWebView", StringComparison.OrdinalIgnoreCase) ||
                       className.Contains("Chrome_WidgetWin", StringComparison.OrdinalIgnoreCase);
            }

            return className.Equals("TeamsWebView", StringComparison.OrdinalIgnoreCase) &&
                   (title.Contains("Teams", StringComparison.OrdinalIgnoreCase) ||
                   title.Contains("チャット", StringComparison.OrdinalIgnoreCase) ||
                   title.Contains("会議", StringComparison.OrdinalIgnoreCase) ||
                   title.Contains("通話", StringComparison.OrdinalIgnoreCase) ||
                   title.Contains("Microsoft", StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsUwpApp(string processPath, string processName)
        {
            if (string.IsNullOrEmpty(processPath))
                return false;

            // Teams (ms-teams.exe / Teams.exe) は例外として許可する
            if (processName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                processName.Equals("Teams", StringComparison.OrdinalIgnoreCase))
                return false;

            return processPath.Contains(@"\WindowsApps\", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// エクスプローラーウィンドウから開いているフォルダのパスを取得する
        /// Shell.Application COMを使用
        /// ProfileExecutor からも参照するため internal
        /// </summary>
        internal static string? GetExplorerFolderPath(IntPtr hWnd)
        {
            try
            {
                var shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType == null) return null;

                var shell = Activator.CreateInstance(shellType);
                if (shell == null) return null;

                var windows = shellType.InvokeMember(
                    "Windows",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, shell, null);

                if (windows == null) return null;

                var countObj = windows.GetType().InvokeMember(
                    "Count",
                    System.Reflection.BindingFlags.GetProperty,
                    null, windows, null);

                int count = Convert.ToInt32(countObj);

                for (int i = 0; i < count; i++)
                {
                    var item = windows.GetType().InvokeMember(
                        "Item",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null, windows, new object[] { i });

                    if (item == null) continue;

                    var hwndObj = item.GetType().InvokeMember(
                        "HWND",
                        System.Reflection.BindingFlags.GetProperty,
                        null, item, null);

                    if (hwndObj == null) continue;

                    IntPtr itemHwnd = new IntPtr(Convert.ToInt64(hwndObj));
                    if (itemHwnd != hWnd) continue;

                    var locationUrl = item.GetType().InvokeMember(
                        "LocationURL",
                        System.Reflection.BindingFlags.GetProperty,
                        null, item, null)?.ToString();

                    if (string.IsNullOrEmpty(locationUrl)) continue;

                    if (locationUrl.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                    {
                        return Uri.UnescapeDataString(locationUrl.Substring(8).Replace('/', '\\'));
                    }

                    return locationUrl;
                }
            }
            catch
            {
                // COM操作失敗時は無視
            }

            return null;
        }

        /// <summary>
        /// タイトルバーからファイルパスを取得する（notepad, EmEditor, Excel, Word, PowerPoint に対応）
        /// </summary>
        private static string? GetTitleFilePath(string processName, string windowTitle)
        {
            var suffixes = processName.ToLower() switch
            {
                "notepad"  => new[] { " - メモ帳", " - Notepad" },
                "excel"    => new[] { " - Excel" },
                "winword"  => new[] { " - Word" },
                "powerpnt" => new[] { " - PowerPoint" },
                "emeditor" => new[] { " - EmEditor" },
                "acrobat"  => new[] { " - Adobe Acrobat Pro", " - Adobe Acrobat", " - Adobe Acrobat Reader", " - Adobe Reader" },
                "acrord32" => new[] { " - Adobe Acrobat Reader", " - Adobe Reader", " - Adobe Acrobat" },
                _ => null
            };

            if (suffixes == null) return null;

            foreach (var suffix in suffixes)
            {
                if (!windowTitle.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) continue;

                var name = windowTitle.Substring(0, windowTitle.Length - suffix.Length).Trim();
                if (!string.IsNullOrEmpty(name) && name != "無題" && name != "Untitled")
                    return name;

                return null;
            }

            return null;
        }

        public static List<ProfileEntry> CaptureCurrentWindows()
        {
            var entries = new List<ProfileEntry>();
            var screens = Screen.AllScreens.OrderBy(s => s.Bounds.X).ToArray();
            ExecutionLogger.Info($"CaptureCurrentWindows start screens={screens.Length}");

            // WMIコマンドラインを1回だけ一括取得（全ウィンドウ処理で使い回す）
            var commandLineCache = CommandLineParser.GetAllCommandLines();
            ExecutionLogger.Info($"WMI CommandLine一括取得完了 count={commandLineCache.Count}");

            // Chrome/Edgeのプロファイルマップを1回だけロード（全ウィンドウ処理で使い回す）
            var chromeProfileMap = ChromeProfileDetector.LoadProfileMap();

            NativeMethods.EnumWindows((hWnd, lParam) =>
            {
                NativeMethods.GetWindowThreadProcessId(hWnd, out uint processId);
                Process process;
                try
                {
                    process = Process.GetProcessById((int)processId);
                }
                catch
                {
                    return true;
                }

                if (!NativeMethods.IsWindowVisible(hWnd) &&
                    !ShouldIncludeInvisibleWindow(process.ProcessName))
                {
                    return true;
                }

                int titleLength = NativeMethods.GetWindowTextLength(hWnd);
                var titleBuilder = new StringBuilder(titleLength + 1);
                if (titleLength > 0)
                {
                    NativeMethods.GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
                }

                var classNameBuilder = new StringBuilder(256);
                NativeMethods.GetClassName(hWnd, classNameBuilder, classNameBuilder.Capacity);
                string className = classNameBuilder.ToString();

                // Teams (ms-teams.exe / Teams.exe) の特別対応: 最小化時にタイトルが空でも
                // 本体ウィンドウらしいサイズ・クラスなら抽出対象にする
                if (IsTeamsProcess(process.ProcessName))
                {
                    if (titleLength == 0)
                    {
                        if (IsMeaningfulTeamsWindow(hWnd, string.Empty, className))
                        {
                            titleBuilder.Append("Microsoft Teams");
                            titleLength = titleBuilder.Length;
                        }
                    }
                }

                // それでもタイトルがない場合は抽出対象外
                if (titleLength == 0)
                    return true;

                if (IsTeamsProcess(process.ProcessName) &&
                    !IsMeaningfulTeamsWindow(hWnd, titleBuilder.ToString(), className))
                {
                    return true;
                }

                if (IsLibreOfficeProcess(process.ProcessName) &&
                    !className.Equals("SALFRAME", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (ExcludedProcesses.Contains(process.ProcessName, StringComparer.OrdinalIgnoreCase))
                    return true;

                string processPath = string.Empty;
                try
                {
                    processPath = process.MainModule?.FileName ?? string.Empty;
                }
                catch
                {
                    // アクセス拒否時は空文字
                }

                // explorerはプロセスパスが取れなくてもフォルダウィンドウとして扱う
                bool isExplorer = process.ProcessName.Equals("explorer", StringComparison.OrdinalIgnoreCase);

                if (!isExplorer)
                {
                    if (IsUwpApp(processPath, process.ProcessName)) return true;
                    if (string.IsNullOrEmpty(processPath)) return true;
                }
                else
                {
                    processPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                        "explorer.exe");
                }

                // 表示状態・座標を取得
                var placement = new NativeMethods.WINDOWPLACEMENT();
                placement.length = System.Runtime.InteropServices.Marshal.SizeOf<NativeMethods.WINDOWPLACEMENT>();
                NativeMethods.GetWindowPlacement(hWnd, ref placement);

                string windowState = placement.showCmd switch
                {
                    NativeMethods.SW_SHOWMINIMIZED => "Minimized",
                    NativeMethods.SW_MAXIMIZE => "Maximized",
                    _ => "Normal"
                };

                var stableBounds = WindowHelper.GetStableWindowBounds(hWnd);
                int absX = stableBounds.X;
                int absY = stableBounds.Y;
                int width = stableBounds.Width;
                int height = stableBounds.Height;

                // Screen.FromHandle を使って正確な表示先モニターを判定
                // （Maximized時の -8 など、Bounds外の座標でも正しく判定してくれる）
                var screen = System.Windows.Forms.Screen.FromHandle(hWnd);
                int monitorNumber = Array.IndexOf(screens, screen) + 1;
                if (monitorNumber <= 0) monitorNumber = 1; // 見つからない場合の保険

                // モニター相対座標に変換
                int relX = absX - screens[monitorNumber - 1].Bounds.X;
                int relY = absY - screens[monitorNumber - 1].Bounds.Y;

                // ファイル/URL取得とChromeプロファイル識別
                string? filePath = null;
                string[]? filePaths = null;
                string? chromeProfileDir  = null;
                string? chromeProfileName = null;

                if (isExplorer)
                {
                    filePath = GetExplorerFolderPath(hWnd);
                }
                else if (AllBrowsers.Contains(process.ProcessName))
                {
                    // BrowserUrlCapturer（ReDesktop移植版）でURL取得
                    var urls = BrowserUrlCapturer.GetUrls(hWnd, process.ProcessName);
                    if (urls.Count > 0)
                    {
                        filePath = urls[0];
                        if (urls.Count > 1)
                        {
                            filePaths = urls.ToArray();
                        }
                    }

                    // Chrome/Edge はプロファイルも識別する
                    if (ChromeLikeBrowsers.Contains(process.ProcessName))
                    {
                        // 優先①：UIAutomation（アバターボタン）でウィンドウ単位に特定
                        chromeProfileDir = ChromeProfileDetector.GetProfileDirectoryFromHwnd(hWnd, chromeProfileMap);

                        // プロファイルディレクトリが取れたら表示名も引く
                        if (!string.IsNullOrEmpty(chromeProfileDir))
                        {
                            chromeProfileName = chromeProfileMap
                                .FirstOrDefault(kvp => kvp.Value == chromeProfileDir).Key;
                        }
                    }
                }
                else if (process.ProcessName.Equals("soffice", StringComparison.OrdinalIgnoreCase) ||
                         process.ProcessName.Equals("soffice.bin", StringComparison.OrdinalIgnoreCase))
                {
                    // LibreOffice: キャッシュ済みコマンドラインから文書パスを抽出
                    if (commandLineCache.TryGetValue((uint)process.Id, out var loCmd))
                    {
                        filePath = GetLibreOfficeDocumentPathFromCommandLine(loCmd);
                    }
                    if (string.IsNullOrEmpty(filePath))
                    {
                        filePath = GetLibreOfficeDocumentPath(process);
                    }
                }
                else
                {
                    string windowTitle = titleBuilder.ToString();

                    // 優先1: WMIコマンドラインからファイルパスを抽出（フルパス取得可能）
                    if (commandLineCache.TryGetValue((uint)process.Id, out var cmdLine))
                    {
                        filePath = CommandLineParser.ExtractFilePath(cmdLine, processPath);
                    }

                    // 優先2: タイトルバーからファイル名を取得（フォールバック）
                    if (string.IsNullOrEmpty(filePath))
                    {
                        filePath = GetTitleFilePath(process.ProcessName, windowTitle);
                    }
                }

                // explorerはフォルダパスが取れた場合のみ追加
                if (isExplorer && string.IsNullOrWhiteSpace(filePath))
                    return true;

                entries.Add(new ProfileEntry
                {
                    IsActive               = true,
                    ProcessPath            = processPath,
                    ProcessName            = process.ProcessName,
                    FilePath               = filePath,
                    FilePaths              = filePaths,
                    ChromeProfileDirectory = chromeProfileDir,
                    ChromeProfileName      = chromeProfileName,
                    Monitor                = monitorNumber,
                    Order                  = null,
                    PositionMode           = "none",
                    CapturedX              = relX,
                    CapturedY              = relY,
                    CapturedWidth          = width,
                    CapturedHeight         = height,
                    CapturedMonitorDeviceName = screens[monitorNumber - 1].DeviceName,
                    CapturedMonitorLeft    = screens[monitorNumber - 1].Bounds.Left,
                    CapturedMonitorTop     = screens[monitorNumber - 1].Bounds.Top,
                    CapturedMonitorWidth   = screens[monitorNumber - 1].Bounds.Width,
                    CapturedMonitorHeight  = screens[monitorNumber - 1].Bounds.Height,
                    CapturedWindowTitle    = titleBuilder.ToString(),
                    CapturedWindowClassName = className,
                    ZOrder                 = entries.Count + 1,
                    WindowState            = windowState
                });

                ExecutionLogger.Info(
                    $"Captured window process={process.ProcessName} title=\"{titleBuilder}\" state={windowState} " +
                    $"bounds={stableBounds} monitor={monitorNumber} class=\"{className}\"");

                return true;
            }, IntPtr.Zero);

            // 全収集後にIsVisible（前面・背面）を判定
            for (int i = 0; i < entries.Count; i++)
            {
                var target = entries[i];
                if (target.WindowState == "Minimized")
                {
                    target.IsVisible = false;
                    continue;
                }

                var targetRect = new Rectangle(
                    target.CapturedX ?? 0, target.CapturedY ?? 0,
                    target.CapturedWidth ?? 0, target.CapturedHeight ?? 0);

                var coverers = entries.Where(e => e.ZOrder < target.ZOrder
                    && e.WindowState != "Minimized").ToList();

                target.IsVisible = !IsCoveredCompletely(targetRect, coverers);
            }

            ExecutionLogger.Info($"CaptureCurrentWindows completed count={entries.Count}");
            return entries;
        }

        private static bool IsCoveredCompletely(Rectangle targetRect, List<ProfileEntry> coverers)
        {
            if (coverers.Count == 0) return false;
            if (targetRect.Width <= 0 || targetRect.Height <= 0) return true;

            var uncovered = new List<Rectangle> { targetRect };

            foreach (var coverer in coverers)
            {
                var covererRect = new Rectangle(
                    coverer.CapturedX ?? 0, coverer.CapturedY ?? 0,
                    coverer.CapturedWidth ?? 0, coverer.CapturedHeight ?? 0);

                if (covererRect.Width <= 0 || covererRect.Height <= 0) continue;

                var newUncovered = new List<Rectangle>();
                foreach (var rect in uncovered)
                    newUncovered.AddRange(SubtractRect(rect, covererRect));

                uncovered = newUncovered;
                if (uncovered.Count == 0) return true;
            }

            return uncovered.Count == 0;
        }

        private static List<Rectangle> SubtractRect(Rectangle baseRect, Rectangle subtractRect)
        {
            if (!baseRect.IntersectsWith(subtractRect))
                return new List<Rectangle> { baseRect };

            var result = new List<Rectangle>();

            int iLeft   = Math.Max(baseRect.Left,   subtractRect.Left);
            int iTop    = Math.Max(baseRect.Top,    subtractRect.Top);
            int iRight  = Math.Min(baseRect.Right,  subtractRect.Right);
            int iBottom = Math.Min(baseRect.Bottom, subtractRect.Bottom);

            if (baseRect.Top < iTop)
                result.Add(new Rectangle(baseRect.Left, baseRect.Top, baseRect.Width, iTop - baseRect.Top));

            if (iBottom < baseRect.Bottom)
                result.Add(new Rectangle(baseRect.Left, iBottom, baseRect.Width, baseRect.Bottom - iBottom));

            if (baseRect.Left < iLeft)
                result.Add(new Rectangle(baseRect.Left, iTop, iLeft - baseRect.Left, iBottom - iTop));

            if (iRight < baseRect.Right)
                result.Add(new Rectangle(iRight, iTop, baseRect.Right - iRight, iBottom - iTop));

            return result;
        }
    }
}
