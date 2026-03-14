namespace StartForm.Models
{
    public class ProfileEntry
    {
        // 起動するかどうか
        public bool IsActive { get; set; } = true;

        // アプリ識別
        public string ProcessPath { get; set; } = string.Empty;
        public string ProcessName { get; set; } = string.Empty;

        // ファイル指定
        public string? FilePath { get; set; }
        public string[]? FilePaths { get; set; }

        // ブラウザ拡張情報
        /// <summary>Chromeのプロファイルフォルダ名（例: "Default", "Profile 1"）</summary>
        public string? ChromeProfileDirectory { get; set; }
        /// <summary>Chromeのプロファイル表示名（例: "仕事用", "個人"）。ユーザーへの表示用。</summary>
        public string? ChromeProfileName { get; set; }

        // 配置指定
        public int? Monitor { get; set; }                          // 配置先モニター番号（1始まり、nullは配置指定なし）
        public int? Order { get; set; }                            // 前面配置時の左からの順番（1始まり）
        public string PositionMode { get; set; } = "none";          // "ratio" / "none" / "minimize"
        public int? Ratio { get; set; }                            // 比率指定配置時の比率（合計10以下）。nullの場合は均等割りに従う

        // 抽出時の内部データ（そのまま再現用）
        public int? CapturedX { get; set; }
        public int? CapturedY { get; set; }
        public int? CapturedWidth { get; set; }
        public int? CapturedHeight { get; set; }
        public string? CapturedMonitorDeviceName { get; set; }
        public int? CapturedMonitorLeft { get; set; }
        public int? CapturedMonitorTop { get; set; }
        public int? CapturedMonitorWidth { get; set; }
        public int? CapturedMonitorHeight { get; set; }
        public string? CapturedWindowTitle { get; set; }
        public string? CapturedWindowClassName { get; set; }
        public int? ZOrder { get; set; }
        public bool IsVisible { get; set; } = true;
        public string WindowState { get; set; } = "Normal";       // Normal / Maximized / Minimized
    }
}
