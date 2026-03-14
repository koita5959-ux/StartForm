using System.Diagnostics;
using System.Management;
using StartForm.Helpers;
using StartForm.Models;
using System.Runtime.InteropServices;

namespace StartForm.Services
{
    public class ProfileExecutor
    {
        private readonly LastPositionService _lastPositionService;
        private static readonly string[] UnsupportedStableApps = { };

        // 既起動プロセスの既存ウィンドウを再利用せず、差分で新しく現れたウィンドウを捕捉する。
        // LibreOffice は既存 soffice プロセスへ起動要求が渡ることがあるため、この扱いが必要。
        private static readonly HashSet<string> NewWindowPreferredProcesses = new(StringComparer.OrdinalIgnoreCase)
        {
            "chrome", "msedge", "soffice"
        };

        private static readonly HashSet<string> InvisibleWindowAllowedProcesses = new(StringComparer.OrdinalIgnoreCase)
        {
            "ms-teams", "Teams", "soffice"
        };

        private static readonly HashSet<string> DelayedVerificationProcesses = new(StringComparer.OrdinalIgnoreCase)
        {
            "soffice", "ms-teams", "Teams", "chrome", "msedge"
        };

        public ProfileExecutor(LastPositionService lastPositionService)
        {
            _lastPositionService = lastPositionService;
        }

        private static void ThrowIfCancellationRequested(CancellationToken cancellationToken, string context)
        {
            if (!cancellationToken.IsCancellationRequested)
                return;

            ExecutionLogger.Warn($"Execution cancelled context={context}");
            cancellationToken.ThrowIfCancellationRequested();
        }

        private static string DescribeEntry(ProfileEntry entry)
        {
            return $"process={entry.ProcessName} mode={entry.PositionMode} monitor={entry.Monitor} " +
                   $"title=\"{entry.CapturedWindowTitle}\" file=\"{entry.FilePath}\"";
        }

        private static bool ShouldPreferNewWindow(ProfileEntry entry)
        {
            if (UnsupportedStableApps.Contains(entry.ProcessName, StringComparer.OrdinalIgnoreCase))
                return true;

            // Chrome/Edge/LibreOffice/Explorerは常に新規ウィンドウを捕捉する
            if (NewWindowPreferredProcesses.Contains(entry.ProcessName, StringComparer.OrdinalIgnoreCase) ||
                entry.ProcessName.Equals("explorer", StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private static HashSet<IntPtr> GetWindowHandlesByProcessName(string processName)
        {
            var handles = new HashSet<IntPtr>();
            var processNames = GetEquivalentProcessNames(processName);

            NativeMethods.EnumWindows((hWnd, lParam) =>
            {
                if (!NativeMethods.IsWindow(hWnd))
                    return true;

                if (!NativeMethods.IsWindowVisible(hWnd) &&
                    !processNames.Any(InvisibleWindowAllowedProcesses.Contains))
                    return true;

                NativeMethods.GetWindowThreadProcessId(hWnd, out uint processId);
                try
                {
                    var process = Process.GetProcessById((int)processId);
                    if (processNames.Contains(process.ProcessName) &&
                        IsEligibleWindow(processName, hWnd))
                    {
                        handles.Add(hWnd);
                    }
                }
                catch
                {
                }

                return true;
            }, IntPtr.Zero);

            return handles;
        }

        private static HashSet<string> GetEquivalentProcessNames(string processName)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                processName
            };

            if (processName.Equals("soffice", StringComparison.OrdinalIgnoreCase) ||
                processName.Equals("soffice.bin", StringComparison.OrdinalIgnoreCase))
            {
                names.Add("soffice");
                names.Add("soffice.bin");
            }

            return names;
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

        private static bool IsMeaningfulLibreOfficeWindow(string title, string className)
        {
            if (!className.Equals("SALFRAME", StringComparison.OrdinalIgnoreCase))
                return false;

            if (title.Equals("DDE Server Window", StringComparison.OrdinalIgnoreCase))
                return false;

            return true;
        }

        private static bool IsTeamsEntryExcluded(ProfileEntry entry)
        {
            if (!IsTeamsProcess(entry.ProcessName))
                return false;

            if (IsTeamsAuxiliaryWindowTitle(entry.CapturedWindowTitle ?? string.Empty))
                return true;

            if (string.Equals(entry.CapturedWindowClassName, "TeamsMessageWindow", StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private static ProfileEntry? SelectPrimaryTeamsEntry(IEnumerable<ProfileEntry> entries)
        {
            return entries
                .Where(entry => !IsTeamsEntryExcluded(entry))
                .OrderByDescending(entry => entry.IsVisible)
                .ThenByDescending(entry => string.Equals(entry.CapturedWindowClassName, "TeamsWebView", StringComparison.OrdinalIgnoreCase))
                .ThenByDescending(entry => (entry.CapturedWidth ?? 0) * (entry.CapturedHeight ?? 0))
                .ThenBy(entry => entry.ZOrder ?? int.MaxValue)
                .FirstOrDefault();
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

        private static string? ExtractProfileDirectoryFromCommandLine(string? commandLine)
        {
            if (string.IsNullOrWhiteSpace(commandLine))
                return null;

            var match = System.Text.RegularExpressions.Regex.Match(
                commandLine,
                @"--profile-directory=(?:""([^""]+)""|([^\s]+))",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (!match.Success)
                return null;

            return !string.IsNullOrWhiteSpace(match.Groups[1].Value)
                ? match.Groups[1].Value
                : match.Groups[2].Value;
        }

        private static string? GetProcessProfileDirectory(Process process)
        {
            var visitedPids = new HashSet<int>();
            Process? current = process;

            while (current != null && visitedPids.Add(current.Id))
            {
                if (current.ProcessName.Equals("chrome", StringComparison.OrdinalIgnoreCase) ||
                    current.ProcessName.Equals("msedge", StringComparison.OrdinalIgnoreCase))
                {
                    var profileDirectory = ExtractProfileDirectoryFromCommandLine(GetProcessCommandLine(current));
                    if (!string.IsNullOrWhiteSpace(profileDirectory))
                        return profileDirectory;
                }

                try
                {
                    using var searcher = new ManagementObjectSearcher(
                        $"SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {current.Id}");

                    Process? parent = null;
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        if (obj["ParentProcessId"] is uint parentPid && parentPid > 0)
                        {
                            try
                            {
                                parent = Process.GetProcessById((int)parentPid);
                            }
                            catch
                            {
                                parent = null;
                            }
                        }
                    }

                    current = parent;
                }
                catch
                {
                    current = null;
                }
            }

            return null;
        }

        private static double CalculateWindowMatchScore(ProfileEntry entry, IntPtr handle, Screen[] orderedScreens)
        {
            var capturedBounds = GetCapturedBounds(entry, orderedScreens);
            if (!capturedBounds.HasValue)
                return double.MaxValue / 4;

            NativeMethods.GetWindowRect(handle, out NativeMethods.RECT rect);
            int currentX = rect.Left;
            int currentY = rect.Top;
            int currentWidth = rect.Right - rect.Left;
            int currentHeight = rect.Bottom - rect.Top;

            // 最小化されている場合は、復元時の座標を使用してスコア計算を行う
            var placement = new NativeMethods.WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf<NativeMethods.WINDOWPLACEMENT>();
            if (NativeMethods.GetWindowPlacement(handle, ref placement) && placement.showCmd == 2) // SW_SHOWMINIMIZED
            {
                currentX = placement.rcNormalPosition.Left;
                currentY = placement.rcNormalPosition.Top;
                currentWidth = placement.rcNormalPosition.Right - placement.rcNormalPosition.Left;
                currentHeight = placement.rcNormalPosition.Bottom - placement.rcNormalPosition.Top;
            }

            double distanceScore = Math.Sqrt(
                Math.Pow(capturedBounds.Value.X - currentX, 2) +
                Math.Pow(capturedBounds.Value.Y - currentY, 2) +
                Math.Pow(capturedBounds.Value.Width - currentWidth, 2) +
                Math.Pow(capturedBounds.Value.Height - currentHeight, 2));

            var targetScreen = ResolveScreen(entry, orderedScreens);
            bool sameScreen = false;
            if (targetScreen != null)
            {
                var currentScreen = Screen.FromHandle(handle);
                sameScreen = string.Equals(currentScreen.DeviceName, targetScreen.DeviceName, StringComparison.OrdinalIgnoreCase);
            }

            int titleLength = NativeMethods.GetWindowTextLength(handle);
            var titleBuilder = new System.Text.StringBuilder(titleLength + 1);
            if (titleLength > 0)
                NativeMethods.GetWindowText(handle, titleBuilder, titleBuilder.Capacity);
            var currentTitle = titleBuilder.ToString();

            bool sameTitle = string.Equals(currentTitle, entry.CapturedWindowTitle, StringComparison.OrdinalIgnoreCase);
            // Teams の場合、仮想タイトルの可能性を考慮して部分一致も許容
            if (!sameTitle && (entry.ProcessName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase) || entry.ProcessName.Equals("Teams", StringComparison.OrdinalIgnoreCase)))
            {
                if ((currentTitle != null && currentTitle.Contains("Microsoft Teams")) || (entry.CapturedWindowTitle != null && entry.CapturedWindowTitle.Contains("Microsoft Teams")))
                    sameTitle = true;
            }

            var classNameBuilder = new System.Text.StringBuilder(256);
            NativeMethods.GetClassName(handle, classNameBuilder, classNameBuilder.Capacity);
            bool sameClassName = string.Equals(classNameBuilder.ToString(), entry.CapturedWindowClassName, StringComparison.OrdinalIgnoreCase);

            bool sameProfileDirectory = true;

            if (!string.IsNullOrWhiteSpace(entry.ChromeProfileDirectory))
            {
                NativeMethods.GetWindowThreadProcessId(handle, out uint processId);
                try
                {
                    var process = Process.GetProcessById((int)processId);
                    var currentProfileDirectory = GetProcessProfileDirectory(process);
                    sameProfileDirectory = string.Equals(
                        currentProfileDirectory,
                        entry.ChromeProfileDirectory,
                        StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    sameProfileDirectory = false;
                }
            }

            // スコアリングの重み付けを調整
            return distanceScore
                + (sameScreen ? 0 : 1_000_000_000d) // 画面が異なる場合のペナルティを大きく
                + (sameProfileDirectory ? 0 : 100_000_000d) // プロファイルが異なる場合のペナルティ
                + (sameTitle ? 0 : 10_000_000d) // タイトルが異なる場合のペナルティ
                + (sameClassName ? 0 : 1_000_000d); // クラス名が異なる場合のペナルティ
        }

        private static Dictionary<ProfileEntry, IntPtr> AssignBestWindowHandles(
            List<ProfileEntry> entries,
            List<IntPtr> handles,
            Screen[] orderedScreens)
        {
            var assignments = new Dictionary<ProfileEntry, IntPtr>();
            if (entries.Count == 0 || handles.Count == 0)
                return assignments;

            if (handles.Count < entries.Count)
            {
                return AssignWindowHandlesGreedy(entries, handles, orderedScreens, preferVisibleEntries: true);
            }

            // 総当たり探索は候補数が少ない場合だけに限定する。
            // Teams のように隠しウィンドウが多いプロセスでは組み合わせ爆発して固まりやすい。
            long combinationCount = 1;
            for (int i = 0; i < entries.Count; i++)
            {
                combinationCount *= Math.Max(1, handles.Count - i);
                if (combinationCount > 200_000)
                    break;
            }

            if (entries.Count > 6 || handles.Count > 8 || combinationCount > 200_000)
            {
                return AssignWindowHandlesGreedy(entries, handles, orderedScreens);
            }

            double bestTotalScore = double.MaxValue;
            var currentAssignments = new Dictionary<ProfileEntry, IntPtr>();
            var usedHandles = new HashSet<IntPtr>();

            void Search(int index, double totalScore)
            {
                if (index >= entries.Count)
                {
                    if (totalScore < bestTotalScore)
                    {
                        bestTotalScore = totalScore;
                        assignments = new Dictionary<ProfileEntry, IntPtr>(currentAssignments);
                    }
                    return;
                }

                if (totalScore >= bestTotalScore)
                    return;

                var entry = entries[index];
                foreach (var handle in handles)
                {
                    if (usedHandles.Contains(handle))
                        continue;

                    double score = CalculateWindowMatchScore(entry, handle, orderedScreens);
                    usedHandles.Add(handle);
                    currentAssignments[entry] = handle;
                    Search(index + 1, totalScore + score);
                    currentAssignments.Remove(entry);
                    usedHandles.Remove(handle);
                }
            }

            Search(0, 0);
            return assignments;
        }

        private static Dictionary<ProfileEntry, IntPtr> AssignWindowHandlesGreedy(
            List<ProfileEntry> entries,
            List<IntPtr> handles,
            Screen[] orderedScreens,
            bool preferVisibleEntries = false)
        {
            var assignments = new Dictionary<ProfileEntry, IntPtr>();
            var remainingHandles = new HashSet<IntPtr>(handles);
            var orderedEntries = entries
                .Select(entry => new
                {
                    Entry = entry,
                    BestScore = remainingHandles
                        .Select(handle => CalculateWindowMatchScore(entry, handle, orderedScreens))
                        .DefaultIfEmpty(double.MaxValue)
                        .Min()
                })
                .OrderByDescending(item => preferVisibleEntries && item.Entry.IsVisible)
                .ThenBy(item => item.BestScore)
                .Select(item => item.Entry)
                .ToList();

            foreach (var entry in orderedEntries)
            {
                IntPtr bestHandle = IntPtr.Zero;
                double bestScore = double.MaxValue;

                foreach (var handle in remainingHandles)
                {
                    double score = CalculateWindowMatchScore(entry, handle, orderedScreens);
                    if (score < bestScore)
                    {
                        bestScore = score;
                        bestHandle = handle;
                    }
                }

                if (bestHandle != IntPtr.Zero)
                {
                    assignments[entry] = bestHandle;
                    remainingHandles.Remove(bestHandle);
                }
            }

            ExecutionLogger.Info(
                $"AssignWindowHandlesGreedy used entries={entries.Count} handles={handles.Count} assigned={assignments.Count}");

            return assignments;
        }

        private static IntPtr FindNewWindowHandle(string processName, HashSet<IntPtr> existingHandles, HashSet<IntPtr>? excludedHandles = null)
        {
            var currentHandles = GetWindowHandlesByProcessName(processName);
            foreach (var handle in currentHandles)
            {
                if (!existingHandles.Contains(handle) &&
                    (excludedHandles == null || !excludedHandles.Contains(handle)))
                    return handle;
            }

            return IntPtr.Zero;
        }

        private static bool IsMinimizeMode(ProfileEntry entry)
        {
            return entry.PositionMode.Equals("minimize", StringComparison.OrdinalIgnoreCase) ||
                   entry.PositionMode.Equals("minimized", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeWindowState(string? windowState)
        {
            return windowState switch
            {
                "Maximized" => "Maximized",
                "Minimized" => "Minimized",
                _ => "Normal"
            };
        }

        private static string GetDesiredWindowState(ProfileEntry entry)
        {
            if (IsMinimizeMode(entry))
                return "Minimized";

            return NormalizeWindowState(entry.WindowState);
        }

        private static bool ShouldVerifyAfterPlacement(ProfileEntry entry)
        {
            return DelayedVerificationProcesses.Contains(entry.ProcessName);
        }

        private static int GetVerificationAttempts(ProfileEntry entry)
        {
            if (entry.ProcessName.Equals("soffice", StringComparison.OrdinalIgnoreCase))
                return 40;

            if (entry.ProcessName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                entry.ProcessName.Equals("Teams", StringComparison.OrdinalIgnoreCase))
                return 16;

            return 12;
        }

        private static int GetBoundsTolerance(ProfileEntry entry)
        {
            if (entry.ProcessName.Equals("soffice", StringComparison.OrdinalIgnoreCase))
                return 2;

            if (IsTeamsProcess(entry.ProcessName))
                return 6;

            return 12;
        }

        private static int GetRequiredStablePasses(ProfileEntry entry)
        {
            if (entry.ProcessName.Equals("soffice", StringComparison.OrdinalIgnoreCase))
                return 8;

            return 2;
        }

        private static IntPtr ResolveHandleForEntry(ProfileEntry entry, IntPtr currentHandle, Screen[] orderedScreens)
        {
            if (currentHandle != IntPtr.Zero && NativeMethods.IsWindow(currentHandle))
                return currentHandle;

            var candidates = GetWindowHandlesByProcessName(entry.ProcessName).ToList();
            if (candidates.Count == 0)
                return IntPtr.Zero;

            if (candidates.Count == 1)
                return candidates[0];

            return candidates
                .OrderBy(hWnd => CalculateWindowMatchScore(entry, hWnd, orderedScreens))
                .FirstOrDefault();
        }

        private static bool IsEligibleWindow(string processName, IntPtr hWnd)
        {
            int titleLength = NativeMethods.GetWindowTextLength(hWnd);
            var titleBuilder = new System.Text.StringBuilder(titleLength + 1);
            if (titleLength > 0)
                NativeMethods.GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
            var title = titleBuilder.ToString();

            var classNameBuilder = new System.Text.StringBuilder(256);
            NativeMethods.GetClassName(hWnd, classNameBuilder, classNameBuilder.Capacity);
            var className = classNameBuilder.ToString();

            if (processName.Equals("explorer", StringComparison.OrdinalIgnoreCase))
            {
                return className.Equals("CabinetWClass", StringComparison.OrdinalIgnoreCase) ||
                       className.Equals("ExploreWClass", StringComparison.OrdinalIgnoreCase);
            }

            // Teams (ms-teams.exe / Teams.exe) は最小化時にタイトルが空になることがあるため特別扱い
            if (IsTeamsProcess(processName))
            {
                return IsMeaningfulTeamsWindow(hWnd, title, className);
            }

            if (IsLibreOfficeProcess(processName))
            {
                return IsMeaningfulLibreOfficeWindow(title, className);
            }

            return titleLength > 0;
        }

        private static Screen[] GetOrderedScreens()
        {
            return Screen.AllScreens
                .OrderBy(s => s.Bounds.X)
                .ThenBy(s => s.Bounds.Y)
                .ToArray();
        }

        private static Screen? ResolveScreen(ProfileEntry entry, Screen[] orderedScreens)
        {
            if (!string.IsNullOrWhiteSpace(entry.CapturedMonitorDeviceName))
            {
                var matchedByDevice = orderedScreens.FirstOrDefault(
                    s => string.Equals(s.DeviceName, entry.CapturedMonitorDeviceName, StringComparison.OrdinalIgnoreCase));
                if (matchedByDevice != null)
                    return matchedByDevice;
            }

            if (entry.Monitor.HasValue)
            {
                int monitorIndex = entry.Monitor.Value - 1;
                if (monitorIndex >= 0 && monitorIndex < orderedScreens.Length)
                    return orderedScreens[monitorIndex];
            }

            return null;
        }

        private static Rectangle? GetCapturedBounds(ProfileEntry entry, Screen[] orderedScreens)
        {
            if (!entry.CapturedX.HasValue || !entry.CapturedY.HasValue ||
                !entry.CapturedWidth.HasValue || !entry.CapturedHeight.HasValue)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(entry.CapturedMonitorDeviceName))
            {
                var deviceScreen = orderedScreens.FirstOrDefault(
                    s => string.Equals(s.DeviceName, entry.CapturedMonitorDeviceName, StringComparison.OrdinalIgnoreCase));
                if (deviceScreen != null)
                {
                    return new Rectangle(
                        deviceScreen.Bounds.Left + entry.CapturedX.Value,
                        deviceScreen.Bounds.Top + entry.CapturedY.Value,
                        entry.CapturedWidth.Value,
                        entry.CapturedHeight.Value);
                }

                return new Rectangle(
                    entry.CapturedX.Value,
                    entry.CapturedY.Value,
                    entry.CapturedWidth.Value,
                    entry.CapturedHeight.Value);
            }

            var screen = ResolveScreen(entry, orderedScreens);
            if (screen != null)
            {
                return new Rectangle(
                    screen.Bounds.Left + entry.CapturedX.Value,
                    screen.Bounds.Top + entry.CapturedY.Value,
                    entry.CapturedWidth.Value,
                    entry.CapturedHeight.Value);
            }

            return new Rectangle(
                entry.CapturedX.Value,
                entry.CapturedY.Value,
                entry.CapturedWidth.Value,
                entry.CapturedHeight.Value);
        }

        private static Rectangle GetPreMaximizeBounds(ProfileEntry entry, Rectangle fallbackBounds, Screen[] orderedScreens)
        {
            var screen = ResolveScreen(entry, orderedScreens);
            if (screen == null)
                return fallbackBounds;

            var workingArea = screen.WorkingArea;
            int width = Math.Min(Math.Max(fallbackBounds.Width, workingArea.Width / 2), workingArea.Width);
            int height = Math.Min(Math.Max(fallbackBounds.Height, workingArea.Height / 2), workingArea.Height);

            int x = workingArea.X + Math.Max(0, (workingArea.Width - width) / 2);
            int y = workingArea.Y + Math.Max(0, (workingArea.Height - height) / 2);

            return new Rectangle(x, y, width, height);
        }

        private static bool RestoreAndMoveWindow(IntPtr hWnd, Rectangle bounds, string windowState)
        {
            return WindowHelper.ApplyWindowPlacement(hWnd, bounds, windowState);
        }

        private static bool ShouldUseNewWindowOnly(ProfileEntry entry)
        {
            return UnsupportedStableApps.Contains(entry.ProcessName, StringComparer.OrdinalIgnoreCase);
        }

        // ShouldSkipEntry は「配置もスキップ」する場合のみ true にする
        // Chrome/Edgeは新規ウィンドウで配置するため false のまま
        private static bool ShouldSkipEntry(ProfileEntry entry)
        {
            if (IsTeamsEntryExcluded(entry))
                return true;

            return UnsupportedStableApps.Contains(entry.ProcessName, StringComparer.OrdinalIgnoreCase);
        }

        private static async Task VerifyAndReapplyAsync(
            List<ReapplyAction> actions,
            Dictionary<ProfileEntry, (Process? process, IntPtr hWnd)> entryHandles,
            Screen[] orderedScreens,
            CancellationToken cancellationToken)
        {
            if (actions.Count == 0)
            {
                ExecutionLogger.Info("VerifyAndReapplyAsync skipped: no actions");
                return;
            }

            ExecutionLogger.Info($"VerifyAndReapplyAsync start actions={actions.Count}");

            int retries = actions.Max(action => GetVerificationAttempts(action.Entry));
            const int intervalMs = 250;
            int stablePasses = 0;
            int requiredStablePasses = actions.Max(action => GetRequiredStablePasses(action.Entry));

            for (int attempt = 0; attempt < retries; attempt++)
            {
                ThrowIfCancellationRequested(cancellationToken, $"verify-attempt-{attempt + 1}");
                bool anyReapplied = false;

                foreach (var action in actions)
                {
                    ThrowIfCancellationRequested(cancellationToken, $"verify-entry-{action.Entry.ProcessName}");
                    if (attempt >= GetVerificationAttempts(action.Entry))
                        continue;

                    if (!entryHandles.TryGetValue(action.Entry, out var current))
                        continue;

                    var resolvedHandle = ResolveHandleForEntry(action.Entry, current.hWnd, orderedScreens);
                    if (resolvedHandle == IntPtr.Zero)
                        continue;

                    if (resolvedHandle != current.hWnd)
                        entryHandles[action.Entry] = (current.process, resolvedHandle);

                    string currentState = NormalizeWindowState(WindowHelper.GetWindowState(resolvedHandle));
                    bool boundsMatch = action.WindowState == "Minimized" ||
                        WindowHelper.IsPlacementClose(resolvedHandle, action.Bounds, GetBoundsTolerance(action.Entry));
                    bool stateMatch = currentState == action.WindowState;

                    if (!boundsMatch || !stateMatch)
                    {
                        anyReapplied = true;
                        var currentBounds = WindowHelper.GetStableWindowBounds(resolvedHandle);
                        ExecutionLogger.Warn(
                            $"Reapply needed attempt={attempt + 1} {DescribeEntry(action.Entry)} hwnd=0x{resolvedHandle.ToInt64():X} " +
                            $"boundsMatch={boundsMatch} stateMatch={stateMatch} currentState={currentState} target={action.Bounds} current={currentBounds}");
                        RestoreAndMoveWindow(resolvedHandle, action.Bounds, action.WindowState);
                    }
                }

                if (anyReapplied)
                {
                    stablePasses = 0;
                }
                else
                {
                    stablePasses++;
                    if (stablePasses >= requiredStablePasses)
                    {
                        ExecutionLogger.Info($"VerifyAndReapplyAsync early exit attempt={attempt + 1}");
                        break;
                    }
                }

                await Task.Delay(intervalMs, cancellationToken);
            }

            ExecutionLogger.Info("VerifyAndReapplyAsync completed");
        }

        private static void RestoreRelativeZOrder(
            IEnumerable<ProfileEntry> entries,
            Dictionary<ProfileEntry, (Process? process, IntPtr hWnd)> entryHandles,
            Screen[] orderedScreens)
        {
            var orderedEntries = entries
                .Where(e => !ShouldSkipEntry(e) && e.ZOrder.HasValue && NormalizeWindowState(e.WindowState) != "Minimized")
                .OrderByDescending(e => e.ZOrder!.Value)
                .ToList();

            if (orderedEntries.Count <= 1)
            {
                ExecutionLogger.Info("RestoreRelativeZOrder skipped: insufficient ordered entries");
                return;
            }

            ExecutionLogger.Info($"RestoreRelativeZOrder start count={orderedEntries.Count}");

            foreach (var entry in orderedEntries)
            {
                if (!entryHandles.TryGetValue(entry, out var current))
                    continue;

                var hWnd = ResolveHandleForEntry(entry, current.hWnd, orderedScreens);
                if (hWnd == IntPtr.Zero)
                    continue;

                entryHandles[entry] = (current.process, hWnd);
                bool result = NativeMethods.SetWindowPos(
                    hWnd,
                    NativeMethods.HWND_TOP,
                    0, 0, 0, 0,
                    NativeMethods.SWP_NOMOVE |
                    NativeMethods.SWP_NOSIZE |
                    NativeMethods.SWP_NOACTIVATE |
                    NativeMethods.SWP_ASYNCWINDOWPOS);

                ExecutionLogger.Info(
                    $"RestoreRelativeZOrder entry={DescribeEntry(entry)} hwnd=0x{hWnd.ToInt64():X} zOrder={entry.ZOrder} result={result}");
            }

            ExecutionLogger.Info("RestoreRelativeZOrder completed");
        }

        public async Task ExecuteAsync(Profile profile, CancellationToken cancellationToken = default)
        {
            var activeEntries = profile.Entries.Where(e => e.IsActive).ToList();
            var entryHandles = new Dictionary<ProfileEntry, (Process? process, IntPtr hWnd)>();
            var launchBaselines = new Dictionary<ProfileEntry, HashSet<IntPtr>>();
            var orderedScreens = GetOrderedScreens();
            var lastPositions = _lastPositionService.Load(profile.ProfileName);
            var skippedEntries = new List<ProfileEntry>();
            var runtimeSkippedEntries = new HashSet<ProfileEntry>();
            var processHandleAssignments = new Dictionary<string, Dictionary<ProfileEntry, IntPtr>>(StringComparer.OrdinalIgnoreCase);
            var claimedNewHandles = new HashSet<IntPtr>();
            var reapplyList = new List<ReapplyAction>(); // Moved declaration here

            ExecutionLogger.Info(
                $"ExecuteAsync start profile=\"{profile.ProfileName}\" activeEntries={activeEntries.Count} log={ExecutionLogger.LogFilePath}");

            ThrowIfCancellationRequested(cancellationToken, "execute-start");

            var primaryTeamsEntry = SelectPrimaryTeamsEntry(activeEntries);
            if (primaryTeamsEntry != null)
            {
                foreach (var teamsEntry in activeEntries.Where(e => IsTeamsProcess(e.ProcessName) && !ReferenceEquals(e, primaryTeamsEntry)))
                {
                    skippedEntries.Add(teamsEntry);
                    runtimeSkippedEntries.Add(teamsEntry);
                    entryHandles[teamsEntry] = (null, IntPtr.Zero);
                    ExecutionLogger.Warn($"Entry skipped: secondary Teams entry {DescribeEntry(teamsEntry)}");
                }
            }

            foreach (var processGroup in activeEntries
                .Where(e => !runtimeSkippedEntries.Contains(e) && !ShouldSkipEntry(e) && !ShouldPreferNewWindow(e))
                .GroupBy(e => e.ProcessName, StringComparer.OrdinalIgnoreCase))
            {
                ThrowIfCancellationRequested(cancellationToken, $"assign-existing-{processGroup.Key}");
                var existingHandles = GetWindowHandlesByProcessName(processGroup.Key).ToList();
                if (existingHandles.Count == 0)
                    continue;

                ExecutionLogger.Info(
                    $"Existing handle assignment process={processGroup.Key} handles={existingHandles.Count} entries={processGroup.Count()}");

                processHandleAssignments[processGroup.Key] = AssignBestWindowHandles(
                    processGroup.ToList(),
                    existingHandles,
                    orderedScreens);
            }

            // ステップ1：既起動判定
            foreach (var entry in activeEntries)
            {
                ThrowIfCancellationRequested(cancellationToken, $"step1-{entry.ProcessName}");
                if (ShouldSkipEntry(entry))
                {
                    skippedEntries.Add(entry);
                    runtimeSkippedEntries.Add(entry);
                    entryHandles[entry] = (null, IntPtr.Zero);
                    ExecutionLogger.Warn($"Entry skipped {DescribeEntry(entry)}");
                    continue;
                }

                if (ShouldPreferNewWindow(entry))
                {
                    entryHandles[entry] = (null, IntPtr.Zero);
                    ExecutionLogger.Info($"Entry will launch as new window {DescribeEntry(entry)}");
                    continue;
                }

                var existing = Process.GetProcessesByName(entry.ProcessName);
                if (existing.Length > 0)
                {
                    IntPtr handle = IntPtr.Zero;
                    bool hasProcessAssignments = processHandleAssignments.TryGetValue(entry.ProcessName, out var assignedHandles);
                    if (hasProcessAssignments &&
                        assignedHandles!.TryGetValue(entry, out var assignedHandle))
                    {
                        handle = assignedHandle;
                    }

                    if (hasProcessAssignments && handle == IntPtr.Zero)
                    {
                        skippedEntries.Add(entry);
                        runtimeSkippedEntries.Add(entry);
                        entryHandles[entry] = (null, IntPtr.Zero);
                        ExecutionLogger.Warn(
                            $"Entry skipped: no unique existing handle available {DescribeEntry(entry)}");
                        continue;
                    }

                    Process target;
                    if (handle != IntPtr.Zero)
                    {
                        NativeMethods.GetWindowThreadProcessId(handle, out uint processId);
                        target = existing.FirstOrDefault(p => p.Id == (int)processId)
                            ?? existing.FirstOrDefault(p => p.MainWindowHandle == handle)
                            ?? existing.FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero)
                            ?? existing[0];
                    }
                    else
                    {
                        target = existing.FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero)
                            ?? existing[0];
                    }

                    var resolvedHandle = handle != IntPtr.Zero
                        ? handle
                        : ResolveHandleForEntry(entry, target.MainWindowHandle, orderedScreens);

                    entryHandles[entry] = (target, resolvedHandle);
                    ExecutionLogger.Info(
                        $"Existing process matched {DescribeEntry(entry)} pid={target.Id} hwnd=0x{resolvedHandle.ToInt64():X}");
                }
                else
                {
                    entryHandles[entry] = (null, IntPtr.Zero);
                    ExecutionLogger.Info($"No existing process found {DescribeEntry(entry)}");
                }
            }

            // ステップ2+3：起動処理とウィンドウ待機（Chrome/Edgeは直列、それ以外は従来通り）
            //
            // Chrome/Edge の直列処理：
            //   1エントリずつ「起動 → 新規ウィンドウが現れるまで待機 → claimed に追加 → 次へ」
            //   これにより同じプロファイルで複数ウィンドウがあっても確実に割り当てられる。
            //   （ReDesktopのWindowRestorerと同じ方式）
            //
            // それ以外のアプリ（非Chrome）：
            //   従来の「起動してから差分で検出」方式

            foreach (var entry in activeEntries)
            {
                ThrowIfCancellationRequested(cancellationToken, $"launch-{entry.ProcessName}");
                if (ShouldSkipEntry(entry) || runtimeSkippedEntries.Contains(entry))
                    continue;

                // 既にhWndが取得済みならスキップ（ステップ1で既起動を割り当て済み）
                if (entryHandles[entry].process != null || entryHandles[entry].hWnd != IntPtr.Zero)
                    continue;

                // Chrome/Edge/LibreOffice、および explorer は直列処理（起動して新規ウィンドウ待機）する
                bool isSequentialApp = NewWindowPreferredProcesses.Contains(entry.ProcessName, StringComparer.OrdinalIgnoreCase) ||
                                       entry.ProcessName.Equals("explorer", StringComparison.OrdinalIgnoreCase);

                // 起動前のウィンドウ一覧を記録
                var baselineBeforeLaunch = GetWindowHandlesByProcessName(entry.ProcessName);

                // 起動引数の組み立て
                var startInfo = new ProcessStartInfo
                {
                    FileName = entry.ProcessPath,
                    // Chrome/Edge/Explorer は UseShellExecute=false で直接起動する。
                    UseShellExecute = !isSequentialApp
                };

                bool isChromeLike =
                    entry.ProcessName.Equals("chrome", StringComparison.OrdinalIgnoreCase) ||
                    entry.ProcessName.Equals("msedge", StringComparison.OrdinalIgnoreCase);

                if (isChromeLike)
                {
                    // ArgumentList を使うことでスペースを含むプロファイル名も安全に渡せる
                    startInfo.ArgumentList.Add("--new-window");
                    if (!string.IsNullOrEmpty(entry.ChromeProfileDirectory))
                        startInfo.ArgumentList.Add($"--profile-directory={entry.ChromeProfileDirectory}");

                    // 複数URL（FilePaths）があれば全て開く
                    if (entry.FilePaths != null && entry.FilePaths.Length > 0)
                    {
                        foreach (var url in entry.FilePaths)
                        {
                            if (!string.IsNullOrEmpty(url))
                                startInfo.ArgumentList.Add(url);
                        }
                    }
                    else if (!string.IsNullOrEmpty(entry.FilePath))
                    {
                        startInfo.ArgumentList.Add(entry.FilePath);
                    }
                }
                else
                {
                    // 非Chrome系（Explorer含む）の引数設定
                    // ProcessStartInfo.ArgumentList を使用すれば、空白を含むパス等も安全に渡される
                    if (entry.FilePaths != null && entry.FilePaths.Length > 0)
                    {
                        foreach (var p in entry.FilePaths)
                            if (!string.IsNullOrEmpty(p)) startInfo.ArgumentList.Add(p);
                    }
                    else if (!string.IsNullOrEmpty(entry.FilePath))
                    {
                        startInfo.ArgumentList.Add(entry.FilePath);
                    }
                }

                // .binファイルはWindowsで直接起動できない場合があるため、.exeへフォールバックする
                // 例: soffice.bin → soffice.exe
                var fileName = startInfo.FileName;
                if (fileName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
                {
                    var exePath = Path.ChangeExtension(fileName, ".exe");
                    if (File.Exists(exePath))
                        startInfo.FileName = exePath;
                }

                Process? process;
                try
                {
                    process = Process.Start(startInfo);
                    ExecutionLogger.Info(
                        $"Process launch requested {DescribeEntry(entry)} file=\"{startInfo.FileName}\" args=\"{startInfo.Arguments}\" useShell={startInfo.UseShellExecute}");
                }
                catch
                {
                    // 直接起動失敗時はUseShellExecute=trueで再試行
                    startInfo.UseShellExecute = true;
                    // 引数の渡し直し
                    if (startInfo.ArgumentList.Count > 0)
                    {
                        // UseShellExecute=true環境ではArgumentsを使う必要がある
                        var args = new string[startInfo.ArgumentList.Count];
                        for (int i = 0; i < args.Length; i++)
                        {
                            var a = startInfo.ArgumentList[i];
                            args[i] = a.Contains(" ") ? $"\"{a}\"" : a;
                        }
                        startInfo.Arguments = string.Join(" ", args);
                        startInfo.ArgumentList.Clear();
                    }
                    process = Process.Start(startInfo);
                    ExecutionLogger.Warn(
                        $"Process launch retried with shell execute {DescribeEntry(entry)} file=\"{startInfo.FileName}\" args=\"{startInfo.Arguments}\"");
                }

                if (process == null) continue;

                // Chrome/Edge/Explorer は直列で新規ウィンドウが現れるまで待機
                if (isSequentialApp)
                {
                    try { process.WaitForInputIdle(3000); } catch { }
                    await Task.Delay(1500, cancellationToken);

                    IntPtr newHandle = IntPtr.Zero;
                    // エクスプローラー等、表示に時間がかかる場合を考慮し、20回（最大10秒）待機
                    for (int retry = 0; retry < 20 && newHandle == IntPtr.Zero; retry++)
                    {
                        ThrowIfCancellationRequested(cancellationToken, $"wait-new-window-{entry.ProcessName}-{retry + 1}");
                        var current = GetWindowHandlesByProcessName(entry.ProcessName);
                        // 起動前に存在したウィンドウ・他エントリで確保済みのウィンドウを除外
                        foreach (var h in current)
                        {
                            if (!baselineBeforeLaunch.Contains(h) && !claimedNewHandles.Contains(h))
                            {
                                newHandle = h;
                                break;
                            }
                        }
                        if (newHandle == IntPtr.Zero)
                            await Task.Delay(500, cancellationToken);
                    }

                    claimedNewHandles.Add(newHandle); // Zero でも追加してよい（重複防止）
                    entryHandles[entry] = (process, newHandle);
                    ExecutionLogger.Info(
                        $"New window resolved {DescribeEntry(entry)} pid={process.Id} hwnd=0x{newHandle.ToInt64():X}");
                }
                else
                {
                    // 非直列アプリ：プロセスを記録してまとめて後処理
                    launchBaselines[entry] = baselineBeforeLaunch;
                    entryHandles[entry] = (process, IntPtr.Zero);
                    ExecutionLogger.Info(
                        $"Process launched awaiting handle {DescribeEntry(entry)} pid={process.Id}");
                }
            }

            // 非直列アプリのウィンドウ待機（差分方式）
            foreach (var entry in activeEntries)
            {
                ThrowIfCancellationRequested(cancellationToken, $"resolve-deferred-{entry.ProcessName}");
                if (ShouldSkipEntry(entry) || runtimeSkippedEntries.Contains(entry)) continue;
                var (process, hWnd) = entryHandles[entry];
                if (process == null || hWnd != IntPtr.Zero) continue;
                if (NewWindowPreferredProcesses.Contains(entry.ProcessName, StringComparer.OrdinalIgnoreCase) ||
                    entry.ProcessName.Equals("explorer", StringComparison.OrdinalIgnoreCase)) continue;

                try { process.WaitForInputIdle(5000); } catch { }
                await Task.Delay(500, cancellationToken);

                IntPtr resolvedHandle = IntPtr.Zero;
                for (int retry = 0; retry < 10 && resolvedHandle == IntPtr.Zero; retry++)
                {
                    ThrowIfCancellationRequested(cancellationToken, $"deferred-loop-{entry.ProcessName}-{retry + 1}");
                    process.Refresh();
                    resolvedHandle = ResolveHandleForEntry(entry, process.MainWindowHandle, orderedScreens);
                    if (resolvedHandle == IntPtr.Zero &&
                        launchBaselines.TryGetValue(entry, out var bl))
                    {
                        resolvedHandle = FindNewWindowHandle(entry.ProcessName, bl, claimedNewHandles);
                    }

                    if (resolvedHandle == IntPtr.Zero)
                        await Task.Delay(500, cancellationToken);
                }
                if (resolvedHandle == IntPtr.Zero && launchBaselines.TryGetValue(entry, out var fb))
                    resolvedHandle = FindNewWindowHandle(entry.ProcessName, fb, claimedNewHandles);

                if (resolvedHandle != IntPtr.Zero)
                    claimedNewHandles.Add(resolvedHandle);
                entryHandles[entry] = (process, resolvedHandle);
                ExecutionLogger.Info(
                    $"Deferred handle resolution {DescribeEntry(entry)} pid={process.Id} hwnd=0x{resolvedHandle.ToInt64():X}");
            }


            // ステップ4：通常配置/最小化配置
            var backEntries = activeEntries
                .Where(e => e.PositionMode == "none" || IsMinimizeMode(e))
                .ToList();

            foreach (var entry in backEntries)
            {
                ThrowIfCancellationRequested(cancellationToken, $"back-placement-{entry.ProcessName}");
                if (ShouldSkipEntry(entry) || runtimeSkippedEntries.Contains(entry)) continue;
                var (_, hWnd) = entryHandles[entry];
                if (hWnd == IntPtr.Zero) continue;

                var captured = GetCapturedBounds(entry, orderedScreens);
                Rectangle? targetBounds = captured;
                if (!targetBounds.HasValue)
                {
                    var lastPos = lastPositions.FirstOrDefault(
                        p => p.ProcessName == entry.ProcessName || p.ProcessPath == entry.ProcessPath);

                    if (lastPos != null)
                        targetBounds = new Rectangle(lastPos.PosX, lastPos.PosY, lastPos.Width, lastPos.Height);
                }

                if (targetBounds.HasValue)
                {
                    var desiredState = GetDesiredWindowState(entry);
                    ExecutionLogger.Info(
                        $"Applying back/minimize placement {DescribeEntry(entry)} hwnd=0x{hWnd.ToInt64():X} bounds={targetBounds.Value} state={desiredState}");
                    RestoreAndMoveWindow(hWnd, targetBounds.Value, desiredState);

                    if (ShouldVerifyAfterPlacement(entry))
                        reapplyList.Add(new ReapplyAction(entry, targetBounds.Value, desiredState));
                }
                else
                {
                    ExecutionLogger.Warn($"No target bounds available {DescribeEntry(entry)}");
                }
            }

            // ステップ5：前面配置（均等割り or 比率指定）
            var frontEntries = activeEntries
                .Where(e => e.PositionMode == "ratio" && e.Monitor.HasValue)
                .ToList();

            var monitorGroups = frontEntries.GroupBy(e => e.Monitor!.Value);

            foreach (var group in monitorGroups)
            {
                ThrowIfCancellationRequested(cancellationToken, $"front-group-{group.Key}");
                int monitorIndex = group.Key - 1;
                if (monitorIndex < 0 || monitorIndex >= orderedScreens.Length) continue;

                var workingArea = orderedScreens[monitorIndex].WorkingArea;
                var sorted = group.OrderBy(e => e.Order ?? 0).ToList();
                int count = sorted.Count;
                if (count == 0) continue;

                // 均等割り判定：
                // Profile.EvenSplit=true、またはRatioが全行nullの場合は均等割り
                bool useEvenSplit = profile.EvenSplit
                    || sorted.All(e => !e.Ratio.HasValue);

                if (useEvenSplit)
                {
                    // 均等割り
                    for (int i = 0; i < count; i++)
                    {
                        ThrowIfCancellationRequested(cancellationToken, $"front-even-{group.Key}-{i + 1}");
                        var entry = sorted[i];
                        if (ShouldSkipEntry(entry) || runtimeSkippedEntries.Contains(entry))
                            continue;
                        var (_, hWnd) = entryHandles[entry];
                        if (hWnd == IntPtr.Zero) continue;

                        int x = workingArea.X + (workingArea.Width * i / count);
                        int width = workingArea.Width * (i + 1) / count - workingArea.Width * i / count;

                        var targetBounds = new Rectangle(x, workingArea.Y, width, workingArea.Height);
                        ExecutionLogger.Info(
                            $"Applying ratio-even placement {DescribeEntry(entry)} hwnd=0x{hWnd.ToInt64():X} bounds={targetBounds}");
                        RestoreAndMoveWindow(hWnd, targetBounds, "Normal");

                        if (ShouldVerifyAfterPlacement(entry))
                            reapplyList.Add(new ReapplyAction(entry, targetBounds, "Normal"));
                    }
                }
                else
                {
                    // 比率指定
                    // Ratioがnullの行は残りを均等に割り当て
                    int totalRatio = sorted.Sum(e => e.Ratio ?? 0);
                    int nullCount = sorted.Count(e => !e.Ratio.HasValue);
                    int remainingRatio = Math.Max(0, 10 - totalRatio);
                    int nullRatio = nullCount > 0 ? remainingRatio / nullCount : 0;

                    int currentX = workingArea.X;
                    int totalWidth = workingArea.Width;

                    for (int i = 0; i < count; i++)
                    {
                        ThrowIfCancellationRequested(cancellationToken, $"front-ratio-{group.Key}-{i + 1}");
                        var entry = sorted[i];
                        if (ShouldSkipEntry(entry) || runtimeSkippedEntries.Contains(entry))
                            continue;
                        var (_, hWnd) = entryHandles[entry];
                        if (hWnd == IntPtr.Zero) continue;

                        int ratio = entry.Ratio ?? nullRatio;
                        int effectiveTotal = sorted.Sum(e => e.Ratio ?? nullRatio);
                        if (effectiveTotal == 0) effectiveTotal = count;

                        int width = totalWidth * ratio / effectiveTotal;

                        var targetBounds = new Rectangle(currentX, workingArea.Y, width, workingArea.Height);
                        ExecutionLogger.Info(
                            $"Applying ratio placement {DescribeEntry(entry)} hwnd=0x{hWnd.ToInt64():X} bounds={targetBounds} ratio={ratio}/{effectiveTotal}");
                        RestoreAndMoveWindow(hWnd, targetBounds, "Normal");

                        if (ShouldVerifyAfterPlacement(entry))
                            reapplyList.Add(new ReapplyAction(entry, targetBounds, "Normal"));

                        currentX += width;
                    }
                }
            }

            // ステップ6：自己回復アプリ対策の監視再適用
            await VerifyAndReapplyAsync(reapplyList, entryHandles, orderedScreens, cancellationToken);

            // ステップ7：対象ウィンドウ間の相対Z順を復元
            ThrowIfCancellationRequested(cancellationToken, "restore-z-order");
            RestoreRelativeZOrder(activeEntries.Where(e => !runtimeSkippedEntries.Contains(e)), entryHandles, orderedScreens);

            // ステップ8：位置記録
            var positions = new List<LastPosition>();
            foreach (var entry in activeEntries)
            {
                ThrowIfCancellationRequested(cancellationToken, $"save-position-{entry.ProcessName}");
                if (ShouldSkipEntry(entry) || runtimeSkippedEntries.Contains(entry))
                    continue;

                var (_, hWnd) = entryHandles[entry];
                if (hWnd == IntPtr.Zero) continue;

                var stableBounds = WindowHelper.GetStableWindowBounds(hWnd);
                positions.Add(new LastPosition
                {
                    ProcessName = entry.ProcessName,
                    ProcessPath = entry.ProcessPath,
                    PosX = stableBounds.Left,
                    PosY = stableBounds.Top,
                    Width = stableBounds.Width,
                    Height = stableBounds.Height,
                    WindowState = NormalizeWindowState(WindowHelper.GetWindowState(hWnd))
                });
            }

            _lastPositionService.Save(profile.ProfileName, positions);
            ExecutionLogger.Info($"ExecuteAsync completed profile=\"{profile.ProfileName}\" savedPositions={positions.Count}");

            if (skippedEntries.Count > 0)
            {
                var appNames = string.Join(", ", skippedEntries
                    .Select(e => e.ProcessName)
                    .Distinct(StringComparer.OrdinalIgnoreCase));
                ExecutionLogger.Warn(
                    $"Skipped apps summary profile=\"{profile.ProfileName}\" apps={appNames}");
            }
        }

        private class ReapplyAction
        {
            public ProfileEntry Entry { get; }
            public Rectangle Bounds { get; }
            public string WindowState { get; }

            public ReapplyAction(ProfileEntry entry, Rectangle bounds, string windowState)
            {
                Entry = entry;
                Bounds = bounds;
                WindowState = windowState;
            }
        }
    }
}
