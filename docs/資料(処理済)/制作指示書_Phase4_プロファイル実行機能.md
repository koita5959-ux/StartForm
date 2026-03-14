# 制作指示書：第4フェーズ - プロファイル実行機能

## この指示書について

C#開発（司令塔）からClaudeCode（実行者）への制作指示です。
第3フェーズ（プロファイル編集画面）が完了したことを前提に、プロファイル実行機能を構築します。
ここがStartFormの核心機能です。

**まず本指示書を読み、把握レポートを出力してください。コード生成はまだ行わないでください。**

---

## 第4フェーズの範囲

### 作るもの
- Services/ProfileExecutor.cs（プロファイル実行エンジン）
- Helpers/NativeMethods.cs（Win32 API宣言の集約）
- Helpers/WindowHelper.cs（ウィンドウ操作のユーティリティ）
- MainForm.cs の「実行」ボタンからProfileExecutorを呼び出す連携
- 前回位置記録の保存処理

### 作らないもの（次フェーズ以降）
- 「現在を拾う」機能（ScreenCapturer）
- インストーラー
- アイコン

---

## フォルダ構成（追加分）

```
StartForm/
├── Helpers/                    ← 新規フォルダ
│   ├── NativeMethods.cs       ← Win32 API宣言
│   └── WindowHelper.cs        ← ウィンドウ操作ユーティリティ
├── Services/
│   ├── ProfileExecutor.cs     ← 新規：プロファイル実行エンジン
│   ├── ProfileService.cs      ← 既存（変更なし）
│   └── LastPositionService.cs ← 既存（変更なし）
└── ...
```

---

## NativeMethods.cs（Helpers/）

Win32 APIのP/Invoke宣言を集約するクラス。

```csharp
using System.Runtime.InteropServices;

namespace StartForm.Helpers
{
    internal static class NativeMethods
    {
        // ウィンドウの位置・サイズを設定
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        // ウィンドウの位置・サイズ・Z順序を設定
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        // ウィンドウの表示状態を変更
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // DWM拡張フレーム境界の取得（シャドウ補正用）
        [DllImport("dwmapi.dll")]
        public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

        // ウィンドウの配置情報を取得
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        // ウィンドウの矩形を取得
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        // 定数
        public static readonly IntPtr HWND_TOP = IntPtr.Zero;
        public const uint SWP_NOSIZE = 0x0001;
        public const uint SWP_NOMOVE = 0x0002;
        public const uint SWP_SHOWWINDOW = 0x0040;
        public const int SW_RESTORE = 9;
        public const int SW_SHOW = 5;
        public const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;

        // 構造体
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }
    }
}
```

---

## WindowHelper.cs（Helpers/）

ウィンドウ操作のユーティリティメソッド。

```
namespace StartForm.Helpers

機能：

1. GetExtendedFrameBounds(IntPtr hWnd) → NativeMethods.RECT
   - DwmGetWindowAttribute で DWMWA_EXTENDED_FRAME_BOUNDS を取得
   - シャドウを除いた実際の描画領域を返す

2. GetShadowMargins(IntPtr hWnd) → (int left, int right, int top, int bottom)
   - GetWindowRect と GetExtendedFrameBounds の差分からシャドウ幅を算出
   - 返り値の例：(7, 7, 0, 7) ← 実機テストで確認済みの値

3. MoveWindowWithShadowCompensation(IntPtr hWnd, int x, int y, int width, int height)
   - シャドウ分を補正してMoveWindowを呼ぶ
   - 実際に渡す座標：x - shadowLeft, y - shadowTop, width + shadowLeft + shadowRight, height + shadowTop + shadowBottom
   - これにより、見た目上の位置が指定した座標に一致する

4. BringToFront(IntPtr hWnd)
   - ShowWindow(hWnd, SW_RESTORE) で最小化解除
   - SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW) で前面に

5. GetWindowState(IntPtr hWnd) → string
   - GetWindowPlacement で状態を取得
   - showCmd: 1=Normal, 2=Minimized, 3=Maximized を文字列で返す
```

---

## ProfileExecutor.cs（Services/）

プロファイル実行エンジン。制作仕様書v1.1の実行フロー（7ステップ）を実装する。

```
namespace StartForm.Services

コンストラクタ：
  public ProfileExecutor(LastPositionService lastPositionService)

メインメソッド：
  public async Task ExecuteAsync(Profile profile)

実行フロー：

  ステップ1：既起動判定
    - profile.Entries の各エントリについて、Process.GetProcessesByName(processName) で既起動か確認
    - 既起動のプロセスは起動処理をスキップし、ウィンドウハンドル（MainWindowHandle）を記録

  ステップ2：起動処理
    - 未起動のエントリについて Process.Start で起動
    - ProcessStartInfo を使用：
      - FileName = entry.ProcessPath
      - Arguments = entry.FilePath（nullでなければ）
      - filePaths（ブラウザ用複数URL）がある場合：
        Arguments = string.Join(" ", entry.FilePaths)
    - 起動したProcessオブジェクトを記録

  ステップ3：ウィンドウ待機
    - 起動したプロセスについて WaitForInputIdle(5000) を呼ぶ（最大5秒）
    - その後 await Task.Delay(1000) で追加待機（ウィンドウ生成の安定化）
    - MainWindowHandle が IntPtr.Zero の場合、最大10回 × 500ms でリトライ
      Process.Refresh() を呼んでから MainWindowHandle を再チェック

  ステップ4：位置配置（ratio モード）
    - positionMode == "ratio" のエントリについて：
      a. Screen.AllScreens から対象モニター（entry.Monitor - 1 でインデックス変換）を取得
      b. モニターの WorkingArea を取得（タスクバーを除いた領域）
      c. 同じモニターの ratio エントリをorder順にソート
      d. 各エントリの位置を算出：
         - 全エントリの ratio 合計を算出（ratioTotal）
         - 左からの累積位置を追跡
         - x = workingArea.X + (workingArea.Width * 左側の比率累計 / ratioTotal)
         - width = workingArea.Width * entry.Ratio / ratioTotal
         - y = workingArea.Y
         - height = workingArea.Height
      e. WindowHelper.MoveWindowWithShadowCompensation で配置

  ステップ5：位置なし配置
    - positionMode == "none" のエントリについて：
      a. LastPositionService.Load(profile.ProfileName) で前回位置を取得
      b. 該当するプロセスの前回位置があれば、MoveWindow で配置
      c. なければ何もしない（アプリのデフォルト位置に委ねる）

  ステップ6：Z順序整列
    - 全エントリを逆順（最後のエントリから）で WindowHelper.BringToFront を呼ぶ
    - これにより、Entries の先頭が最前面に来る

  ステップ7：位置記録
    - 全エントリのウィンドウ位置を取得（GetWindowRect）
    - List<LastPosition> を構築
    - LastPositionService.Save(profile.ProfileName, positions) で保存
```

### ウィンドウハンドルの管理

実行フロー全体で、各エントリのプロセスとウィンドウハンドルを紐付ける必要がある。
内部で Dictionary<ProfileEntry, (Process process, IntPtr hWnd)> のような構造を持つ。

---

## MainForm.cs の変更

### 実行ボタンの変更

```
現在：MessageBox "未実装"
変更後：
  1. 選択中のプロファイル名を取得
  2. ProfileService.Load(profileName) でプロファイルを読み込む
  3. ProfileExecutor executor = new ProfileExecutor(new LastPositionService())
  4. await executor.ExecuteAsync(profile)
  5. 実行中はボタンを無効化し、完了後に復帰
  6. エラーが発生した場合は MessageBox でエラー内容を表示

注意：btnExecute_Click を async にする必要がある
  private async void BtnExecute_Click(object sender, EventArgs e)
```

---

## 変更・追加するファイル

| ファイル | 内容 |
|---------|------|
| Helpers/NativeMethods.cs | 新規作成 |
| Helpers/WindowHelper.cs | 新規作成 |
| Services/ProfileExecutor.cs | 新規作成 |
| Forms/MainForm.cs | 実行ボタンの処理を変更 |

---

## 完了条件

1. `dotnet build` がエラーなく通ること
2. プロファイルを作成し「実行」ボタンを押すと、指定したアプリが起動すること
3. positionMode="ratio" のエントリが、指定した画面の指定した比率の位置に配置されること
4. シャドウ補正により、ウィンドウ間に隙間が生じないこと
5. positionMode="none" のエントリは、起動のみ行われること
6. 実行後、前回位置が記録されること

---

## 完了後の報告

作業完了したら、以下をレポートとして出力してください：

- 変更・追加したファイルの一覧と変更内容の概要
- dotnet buildの結果
- 指示書との差異がある場合はその理由
