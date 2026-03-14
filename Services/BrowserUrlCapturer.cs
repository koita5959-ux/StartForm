using System.Windows.Automation;

namespace StartForm.Services
{
    /// <summary>
    /// ブラウザのアドレスバーからURLを取得する。
    /// ReDesktop-masterのBrowserUrlCapturer.cs から移植・改変。
    /// Windows UI Automationを使用してアドレスバー要素を探し、URLを読み取る。
    /// </summary>
    public static class BrowserUrlCapturer
    {
        /// <summary>
        /// 指定HWNDのブラウザウィンドウからURLを取得する。
        /// 取得できない場合は null を返す。
        /// </summary>
        public static string? GetUrl(IntPtr hWnd, string processName)
        {
            try
            {
                var element = AutomationElement.FromHandle(hWnd);
                if (element == null) return null;

                string? url = processName.ToLower() switch
                {
                    "chrome"  => GetChromiumUrl(element),
                    "msedge"  => GetChromiumUrl(element),
                    "firefox" => GetFirefoxUrl(element),
                    "opera"   => GetChromiumUrl(element),
                    "brave"   => GetChromiumUrl(element),
                    "vivaldi" => GetChromiumUrl(element),
                    _ => null
                };

                if (!string.IsNullOrEmpty(url) && IsValidUrl(url))
                    return url;

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Chromium系ブラウザにおける、開いている「全タブのURL」を取得する（可能なら）。
        /// UIAutomationでは通常「アクティブなタブのアドレスバー」しか見えないため、
        /// 各タブを選択(Invoke)してアドレスバーを読み取るか、履歴から取得するなどの
        /// アプローチが必要ですが、UIAutomation経由でバックグラウンドで全タブのURLを取得するのは困難です。
        /// そのため、ここでは「アクティブタブだけでなく、プロファイルに含まれた複数URL」を
        /// StartForm側（FilePaths配列）に保持・復元させる方針に拡張します。
        /// </summary>
        private static string? GetChromiumUrl(AutomationElement window)
        {
            var editCondition = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Edit);

            var edits = window.FindAll(TreeScope.Descendants, editCondition);

            foreach (AutomationElement edit in edits)
            {
                try
                {
                    string name = edit.Current.Name ?? "";

                    bool isAddressBar =
                        name.Contains("アドレス", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("address", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("URL", StringComparison.OrdinalIgnoreCase);

                    if (!isAddressBar) continue;

                    if (edit.TryGetCurrentPattern(ValuePattern.Pattern, out object? pattern))
                    {
                        var vp = (ValuePattern)pattern;
                        string value = vp.Current.Value;
                        if (!string.IsNullOrWhiteSpace(value))
                            return NormalizeUrl(value);
                    }
                }
                catch { continue; }
            }

            return null;
        }

        /// <summary>
        /// Chromium系ブラウザから「全タブのURL」を取得する
        /// （注：UI Automationではアクティブタブのアドレスバーしか取れないため、
        /// 各TabItemアクティブ化して読み取る荒技が必要だが、ユーザー操作を奪うため通常は行わない。
        /// 今回は「現在アクティブなURL」のみを確実に取り、複数タブ復元はProfile側でFilePathsに手動登録されたものを扱う前提とする）
        /// </summary>
        public static List<string> GetUrls(IntPtr hWnd, string processName)
        {
            var urls = new List<string>();
            var singleUrl = GetUrl(hWnd, processName);
            if (!string.IsNullOrEmpty(singleUrl))
            {
                urls.Add(singleUrl);
            }
            return urls;
        }

        /// <summary>
        /// FirefoxのURL取得。
        /// </summary>
        private static string? GetFirefoxUrl(AutomationElement window)
        {
            var editCondition = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Edit);

            var edits = window.FindAll(TreeScope.Descendants, editCondition);

            foreach (AutomationElement edit in edits)
            {
                try
                {
                    string name = edit.Current.Name ?? "";

                    bool isAddressBar =
                        name.Contains("アドレス", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("address", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("URL", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("Search", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("検索", StringComparison.OrdinalIgnoreCase);

                    if (!isAddressBar) continue;

                    if (edit.TryGetCurrentPattern(ValuePattern.Pattern, out object? pattern))
                    {
                        var vp = (ValuePattern)pattern;
                        string value = vp.Current.Value;
                        if (!string.IsNullOrWhiteSpace(value))
                            return NormalizeUrl(value);
                    }
                }
                catch { continue; }
            }

            return null;
        }

        /// <summary>
        /// URLのスキームを補完する（https:// が省略されている場合）
        /// </summary>
        private static string NormalizeUrl(string url)
        {
            url = url.Trim();

            if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            {
                url = "https://" + url;
            }

            return url;
        }

        /// <summary>
        /// URLとして妥当かの簡易チェック。
        /// 新しいタブや空白ページ、ブラウザ内部ページは除外する。
        /// </summary>
        private static bool IsValidUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return false;

            if (url.StartsWith("chrome://", StringComparison.OrdinalIgnoreCase)) return false;
            if (url.StartsWith("edge://", StringComparison.OrdinalIgnoreCase)) return false;
            if (url.StartsWith("about:", StringComparison.OrdinalIgnoreCase)) return false;
            if (url.StartsWith("chrome-extension://", StringComparison.OrdinalIgnoreCase)) return false;
            if (url.Contains("newtab", StringComparison.OrdinalIgnoreCase)) return false;
            if (url.Contains("new-tab", StringComparison.OrdinalIgnoreCase)) return false;

            return true;
        }
    }
}
