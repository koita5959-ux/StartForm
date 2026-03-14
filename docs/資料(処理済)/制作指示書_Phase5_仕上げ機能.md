# 制作指示書：第5フェーズ - 仕上げ機能

## この指示書について

C#開発（司令塔）からClaudeCode（実行者）への制作指示です。
第4フェーズ（プロファイル実行機能）が完了したことを前提に、仕上げ機能を構築します。

**まず本指示書を読み、把握レポートを出力してください。コード生成はまだ行わないでください。**

---

## 第5フェーズの範囲

### 作るもの
- Mutex制御による多重起動防止（Program.cs）
- Services/ScreenCapturer.cs（「現在を拾う」機能）
- EditForm.cs に「現在を拾う」ボタンの処理を接続
- アセンブリ情報の設定（csproj）

### 作らないもの（次フェーズ）
- dotnet publish
- Inno Setupインストーラー
- カスタムアイコン

---

## 1. Mutex制御（Program.cs の変更）

多重起動を防止する。StartFormが既に起動している場合、2つ目の起動を抑止する。

```csharp
namespace StartForm
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            const string mutexName = "StartForm_SingleInstance_Mutex";
            using var mutex = new Mutex(true, mutexName, out bool createdNew);

            if (!createdNew)
            {
                MessageBox.Show("StartFormは既に起動しています。", "StartForm",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ApplicationConfiguration.Initialize();
            Application.Run(new Forms.MainForm());
        }
    }
}
```

---

## 2. ScreenCapturer.cs（Services/）

現在開いているウィンドウの情報を取得し、List<ProfileEntry> として返す。
ReDesktop Resumeの ScreenCapturer.CaptureAll() の技術を転用する。

```
namespace StartForm.Services

機能：
  public static List<ProfileEntry> CaptureCurrentWindows()

処理：
  1. EnumWindows を使って表示中のウィンドウ一覧を取得
     - 必要なP/Invoke追加（NativeMethods.csに追加）：
       - EnumWindows(EnumWindowsProc, IntPtr)
       - IsWindowVisible(IntPtr)
       - GetWindowText(IntPtr, StringBuilder, int)
       - GetWindowTextLength(IntPtr)
       - GetWindowThreadProcessId(IntPtr, out uint)

  2. 各ウィンドウについて以下をフィルタ：
     - IsWindowVisible == true
     - ウィンドウタイトルが空でない
     - 除外プロセスリストに含まれない

  3. 除外プロセスリスト（internal static なフィールド）：
     - "explorer"（デスクトップシェル。エクスプローラーウィンドウはhWndベースで別途判定）
     - "StartForm"（自分自身）
     - "SystemSettings"（Windows設定、UWPイマーシブ）
     - "TextInputHost"
     - "ApplicationFrameHost"（UWPフレーム）
     - "ShellExperienceHost"
     - "SearchHost"
     - "StartMenuExperienceHost"

  4. 各ウィンドウからProfileEntryを構築：
     - ProcessPath：Process.MainModule.FileName で取得（アクセス拒否時は空文字）
     - ProcessName：Process.ProcessName
     - FilePath：null（自動取得は複雑なため、第1版では空欄。利用者が手動で補完）
     - Monitor：Screen.FromHandle(hWnd) で所属モニターを判定し、Screen.AllScreens のインデックス+1
     - Ratio：null（利用者が設定）
     - Order：null（利用者が設定）
     - PositionMode："none"（デフォルト）

  5. List<ProfileEntry> を返す
```

### NativeMethods.cs への追加

```csharp
// EnumWindows用
public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

[DllImport("user32.dll")]
public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

[DllImport("user32.dll")]
public static extern bool IsWindowVisible(IntPtr hWnd);

[DllImport("user32.dll", CharSet = CharSet.Unicode)]
public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

[DllImport("user32.dll")]
public static extern int GetWindowTextLength(IntPtr hWnd);

[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
```

---

## 3. EditForm.cs の変更

### 「現在を拾う」ボタンの追加

ボタン行の左側グループに追加する。「行を追加」「行を削除」の前に配置。

| ボタン | 変数名 | テキスト | 幅 |
|--------|--------|---------|-----|
| 現在を拾う | btnCapture | 現在を拾う | 120 |
| 行を追加 | btnAddRow | 行を追加 | 100 |
| 行を削除 | btnDeleteRow | 行を削除 | 100 |

### btnCapture.Click の処理

```
処理：
  1. 確認ダイアログ：
     "現在のデスクトップ状態を取り込みます。\n既存の設定は置き換えられます。よろしいですか？"
     ボタン：はい / いいえ
  2. 「はい」の場合：
     a. ScreenCapturer.CaptureCurrentWindows() を呼び出す
     b. DataGridViewをクリア
     c. 取得したList<ProfileEntry>を DataGridView に流し込む
     d. 自動取得できなかったフィールド（FilePath等）は空欄のまま
  3. 「いいえ」の場合：何もしない
```

---

## 4. アセンブリ情報（StartForm.csproj への追加）

```xml
<PropertyGroup>
    <Version>1.0.0</Version>
    <AssemblyVersion>1.0.0.0</AssemblyVersion>
    <FileVersion>1.0.0.0</FileVersion>
    <Product>StartForm</Product>
    <Company>便利アプリシリーズ</Company>
    <Description>自分が決めた作業環境を、いつでもワンアクションで整えるデスクトップアプリ</Description>
    <Copyright>Copyright © 2026</Copyright>
</PropertyGroup>
```

既存のPropertyGroupに追記する形で。

---

## 変更・追加するファイル

| ファイル | 内容 |
|---------|------|
| Program.cs | Mutex制御に変更 |
| Services/ScreenCapturer.cs | 新規作成 |
| Helpers/NativeMethods.cs | EnumWindows関連API追加 |
| Forms/EditForm.cs | 「現在を拾う」ボタン追加・処理実装 |
| StartForm.csproj | アセンブリ情報追加 |

---

## 完了条件

1. `dotnet build` がエラーなく通ること
2. StartFormを2つ起動しようとすると「既に起動しています」と表示されること
3. プロファイル編集画面で「現在を拾う」を押すと、現在開いているウィンドウが一覧に取り込まれること
4. 取り込まれたエントリにプロセスパスとモニター番号が設定されていること
5. StartForm自身が一覧に含まれないこと

---

## 完了後の報告

作業完了したら、以下をレポートとして出力してください：

- 変更・追加したファイルの一覧と変更内容の概要
- dotnet buildの結果
- 指示書との差異がある場合はその理由
