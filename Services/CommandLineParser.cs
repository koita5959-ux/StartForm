using System.Management;
using StartForm.Helpers;

namespace StartForm.Services
{
    /// <summary>
    /// プロセスのコマンドライン引数からファイルパスを抽出する。
    /// WMI（Win32_Process）経由でコマンドラインを一括取得し、
    /// 引数部分からファイルパスを特定する。
    ///
    /// ReDesktop の CommandLineParser を StartForm 向けに移植。
    ///
    /// 取得できるケース：
    ///   - ファイルをダブルクリックして開いた場合
    ///   - 「ファイルを開く」でアプリに渡された場合
    ///   - コマンドラインでファイルパスを指定して起動した場合
    ///
    /// 取得できないケース：
    ///   - アプリ内の「ファイルを開く」ダイアログで後から開いた場合
    ///   - アプリが独自のファイル管理をしている場合（最近使ったファイル等）
    /// </summary>
    public static class CommandLineParser
    {
        /// <summary>
        /// 全プロセスのコマンドラインを一括取得する。
        /// CaptureCurrentWindows() の冒頭で1回だけ呼び、結果をキャッシュとして使う。
        /// WMIクエリはプロセスごとに呼ぶと重いため、一括取得が必須。
        /// </summary>
        public static Dictionary<uint, string> GetAllCommandLines()
        {
            var result = new Dictionary<uint, string>();
            try
            {
                using var searcher = new ManagementObjectSearcher(
                    "SELECT ProcessId, CommandLine FROM Win32_Process WHERE CommandLine IS NOT NULL");
                using var collection = searcher.Get();

                foreach (ManagementObject obj in collection)
                {
                    try
                    {
                        var pid = Convert.ToUInt32(obj["ProcessId"]);
                        var cmdLine = obj["CommandLine"]?.ToString();
                        if (!string.IsNullOrEmpty(cmdLine))
                        {
                            result[pid] = cmdLine;
                        }
                    }
                    catch { continue; }
                }
            }
            catch (Exception ex)
            {
                ExecutionLogger.Error("WMI CommandLine一括取得エラー: " + ex.Message);
            }
            return result;
        }

        /// <summary>
        /// コマンドラインからファイルパスを抽出する。
        ///
        /// 戦略:
        ///   1. コマンドラインから実行ファイル部分を除去して引数部分を取得
        ///   2. 引数からファイルパスらしきものを探す
        ///   3. File.Exists で実在確認 → 実在すればフルパスとして採用
        /// </summary>
        public static string? ExtractFilePath(string commandLine, string processPath)
        {
            if (string.IsNullOrWhiteSpace(commandLine)) return null;
            if (string.IsNullOrWhiteSpace(processPath)) return null;

            try
            {
                string arguments = RemoveExecutablePart(commandLine, processPath);
                if (string.IsNullOrWhiteSpace(arguments)) return null;

                var candidates = ExtractPathCandidates(arguments);

                foreach (var candidate in candidates)
                {
                    string trimmed = candidate.Trim().Trim('"').Trim();
                    if (string.IsNullOrEmpty(trimmed)) continue;

                    if (File.Exists(trimmed))
                    {
                        return Path.GetFullPath(trimmed);
                    }

                    // フラグやオプションはスキップ
                    if (trimmed.StartsWith("-") ||
                        (trimmed.StartsWith("/") && !IsLikelyDrivePath(trimmed)))
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                ExecutionLogger.Error("CommandLine解析エラー: " + ex.Message);
            }

            return null;
        }

        /// <summary>
        /// コマンドラインから実行ファイル部分を除去し、引数部分を返す。
        /// </summary>
        private static string RemoveExecutablePart(string commandLine, string processPath)
        {
            // パターン1: "C:\Program Files\...\app.exe" args...
            if (commandLine.StartsWith("\""))
            {
                int closeQuote = commandLine.IndexOf('"', 1);
                if (closeQuote > 0 && closeQuote + 1 < commandLine.Length)
                {
                    return commandLine.Substring(closeQuote + 1).Trim();
                }
                return "";
            }

            // パターン2: processPathが含まれている場合
            if (!string.IsNullOrEmpty(processPath))
            {
                int idx = commandLine.IndexOf(processPath, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    int afterExe = idx + processPath.Length;
                    if (afterExe < commandLine.Length)
                    {
                        return commandLine.Substring(afterExe).Trim();
                    }
                    return "";
                }

                // ファイル名だけでマッチ
                string exeName = Path.GetFileName(processPath);
                idx = commandLine.IndexOf(exeName, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    int afterExe = idx + exeName.Length;
                    if (afterExe < commandLine.Length)
                    {
                        return commandLine.Substring(afterExe).Trim();
                    }
                    return "";
                }
            }

            // パターン3: 最初のスペースまでが実行ファイル名
            int space = commandLine.IndexOf(' ');
            if (space > 0)
            {
                return commandLine.Substring(space + 1).Trim();
            }

            return "";
        }

        /// <summary>
        /// 引数文字列からファイルパス候補を抽出する。
        /// </summary>
        private static List<string> ExtractPathCandidates(string arguments)
        {
            var candidates = new List<string>();

            // クォートで囲まれた部分を優先的に抽出
            int i = 0;
            while (i < arguments.Length)
            {
                if (arguments[i] == '"')
                {
                    int close = arguments.IndexOf('"', i + 1);
                    if (close > i)
                    {
                        string quoted = arguments.Substring(i + 1, close - i - 1);
                        candidates.Add(quoted);
                        i = close + 1;
                        continue;
                    }
                }
                i++;
            }

            // クォートなしの引数も候補に加える
            string remaining = arguments;
            foreach (var quoted in candidates)
            {
                remaining = remaining.Replace("\"" + quoted + "\"", " ");
            }

            // ドライブレター（C:\ 等）で始まるパスを検出
            for (int j = 0; j < remaining.Length - 2; j++)
            {
                if (char.IsLetter(remaining[j]) && remaining[j + 1] == ':' && remaining[j + 2] == '\\')
                {
                    string pathPart = remaining.Substring(j);
                    int flagStart = FindNextFlag(pathPart);
                    if (flagStart > 0)
                    {
                        pathPart = pathPart.Substring(0, flagStart).Trim();
                    }
                    candidates.Add(pathPart);
                    break;
                }
            }

            // スペース区切りの個別トークンも追加
            var tokens = remaining.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            foreach (var token in tokens)
            {
                if (!candidates.Contains(token))
                {
                    candidates.Add(token);
                }
            }

            return candidates;
        }

        /// <summary>
        /// パス文字列中で次のコマンドラインフラグの位置を返す。
        /// </summary>
        private static int FindNextFlag(string s)
        {
            for (int i = 1; i < s.Length - 1; i++)
            {
                if (s[i] == ' ' && i + 1 < s.Length && (s[i + 1] == '-' || s[i + 1] == '/'))
                {
                    if (s[i + 1] == '/' && i + 2 < s.Length && char.IsLetter(s[i + 2]))
                    {
                        continue;
                    }
                    return i;
                }
            }
            return -1;
        }

        private static bool IsLikelyDrivePath(string s)
        {
            return s.Length >= 3 && char.IsLetter(s[0]) && s[1] == ':' && s[2] == '\\';
        }
    }
}
