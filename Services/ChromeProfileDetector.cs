using System.IO;
using System.Text.Json;
using System.Windows.Automation;
using StartForm.Models;

namespace StartForm.Services
{
    /// <summary>
    /// Chromeのプロファイル情報を保持するレコード
    /// </summary>
    public record ChromeProfileInfo(string Directory, string DisplayName, string GaiaName);

    /// <summary>
    /// Chromeの開いているウィンドウからプロファイルを識別するクラス。
    /// ReDesktop-masterのChromeProfileDetector.cs から移植・改変。
    ///
    /// 識別方法：
    ///  優先① UIAutomation で「AvatarToolbarButton」（顔アイコン）の Name を読む
    ///         → ウィンドウごとに確実にプロファイルを特定できる唯一の手法
    ///  優先② フォールバック：コマンドライン引数 --profile-directory= を解析
    /// </summary>
    public static class ChromeProfileDetector
    {
        /// <summary>
        /// ChromeのLocal Stateファイルから gaia_name/name → profileDirectory のマップを返す。
        /// 例: { "仕事用" -> "Profile 1", "個人" -> "Default" }
        /// </summary>
        public static Dictionary<string, string> LoadProfileMap()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var localStatePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Google", "Chrome", "User Data", "Local State");

                if (!File.Exists(localStatePath)) return result;

                var json = File.ReadAllText(localStatePath);
                using var doc = JsonDocument.Parse(json);

                var infoCache = doc.RootElement
                    .GetProperty("profile")
                    .GetProperty("info_cache");

                foreach (var profile in infoCache.EnumerateObject())
                {
                    var dirName = profile.Name;
                    var gaiaName = "";
                    var displayName = "";

                    if (profile.Value.TryGetProperty("gaia_name", out var g))
                        gaiaName = g.GetString() ?? "";
                    if (profile.Value.TryGetProperty("name", out var n))
                        displayName = n.GetString() ?? "";

                    // gaia_name（Googleアカウント名）を優先、なければ表示名
                    var key = !string.IsNullOrEmpty(gaiaName) ? gaiaName : displayName;
                    if (!string.IsNullOrEmpty(key) && !result.ContainsKey(key))
                        result[key] = dirName;
                }
            }
            catch { }

            return result;
        }

        /// <summary>
        /// 指定HWNDのChromeウィンドウからプロファイルディレクトリ名を取得する。
        /// UIAutomation で「AvatarToolbarButton」のボタン名（アカウント表示名）を読み取り、
        /// Local Stateのプロファイルマップと照合する。
        /// 取得できない場合は null を返す。
        /// </summary>
        public static string? GetProfileDirectoryFromHwnd(IntPtr hwnd, Dictionary<string, string> profileMap)
        {
            try
            {
                var element = AutomationElement.FromHandle(hwnd);

                // ReDesktopオリジナルと同じ: Descendants で AvatarToolbarButton を直接検索
                var avatarCondition = new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.ClassNameProperty, "AvatarToolbarButton"));
                var avatarButtons = element.FindAll(TreeScope.Descendants, avatarCondition);

                foreach (AutomationElement btn in avatarButtons)
                {
                    var name = btn.Current.Name;
                    if (string.IsNullOrEmpty(name)) continue;

                    // 完全一致
                    if (profileMap.TryGetValue(name, out var dir))
                        return dir;

                    // 部分一致
                    foreach (var kvp in profileMap)
                    {
                        if (kvp.Key.Contains(name, StringComparison.OrdinalIgnoreCase) ||
                            name.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase))
                            return kvp.Value;
                    }
                }
            }
            catch { }

            return null;
        }
    }
}
