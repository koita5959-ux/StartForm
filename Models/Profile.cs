namespace StartForm.Models
{
    public class Profile
    {
        public string ProfileName { get; set; } = string.Empty;    // プロファイル名（利用者が呼ぶ名前）
        public DateTime CreatedAt { get; set; } = DateTime.Now;     // 作成日時
        public DateTime UpdatedAt { get; set; } = DateTime.Now;     // 最終更新日時
        public List<ProfileEntry> Entries { get; set; } = new();    // アプリ一覧

        // 前面アプリの分割方式
        // true = 均等割り（比率列を無視して自動等分割）
        // false = 比率指定（各エントリのRatioに従って分割）
        public bool EvenSplit { get; set; } = true;
    }
}
