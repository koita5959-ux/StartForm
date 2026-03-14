# 制作指示書：第1フェーズ - プロジェクト骨格とデータ層

## この指示書について

本指示書は、C#開発（司令塔）からClaudeCode（実行者）への制作指示です。
StartFormプロジェクトの第1フェーズとして、プロジェクトの骨格とデータ層のみを構築します。

**まずCLAUDE.mdを読んでから、本指示書に取り組んでください。**

---

## 作業の進め方

1. 本指示書を読み、理解した内容をレポートとして出力してください
2. 不明点があれば質問してください
3. 認識が合ったことを確認してから、コード生成に入ります

**「把握しました」で進まないでください。何を理解したかを具体的に書いてください。**

---

## 第1フェーズの範囲

### 作るもの
- .NET 8.0 Windows Formsプロジェクトの作成
- フォルダ構成の確立
- プロファイルのデータモデル（クラス定義）
- 前回位置記録のデータモデル（クラス定義）
- JSON読み書きサービス（保存・読み込み・削除）
- Program.cs（エントリーポイント、最小限のメインウィンドウ起動）
- メインウィンドウの空フォーム（タイトルとサイズだけ）

### 作らないもの（次フェーズ以降）
- プロファイル編集画面のUI
- 「現在を拾う」機能
- アプリ起動・配置ロジック
- シャドウ補正
- Win32 API呼び出し
- インストーラー
- アイコン

---

## フォルダ構成

```
StartForm/
├── StartForm.csproj
├── Program.cs
├── Models/
│   ├── ProfileEntry.cs        ← プロファイル内の1行（アプリ1つ分）
│   ├── Profile.cs             ← プロファイル全体
│   └── LastPosition.cs        ← 前回位置記録の1行
├── Services/
│   ├── ProfileService.cs      ← プロファイルの読み書き・削除
│   └── LastPositionService.cs ← 前回位置記録の読み書き
├── Forms/
│   └── MainForm.cs            ← メインウィンドウ（空フォーム）
└── CLAUDE.md                  ← （既存、触らない）
```

---

## プロジェクトファイル（StartForm.csproj）

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net8.0-windows</TargetFramework>
    <UseWindowsForms>true</UseWindowsForms>
    <RootNamespace>StartForm</RootNamespace>
    <AssemblyName>StartForm</AssemblyName>
    <PublishSingleFile>true</PublishSingleFile>
    <SelfContained>true</SelfContained>
    <RuntimeIdentifier>win-x64</RuntimeIdentifier>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="System.Management" Version="8.0.0" />
  </ItemGroup>

</Project>
```

注意：System.Managementは第1フェーズでは使わないが、後のフェーズで必要になるため最初から入れておく。

---

## データモデル

### ProfileEntry.cs（Models/）

プロファイル内の1行。アプリ1つ分の設定。

```csharp
namespace StartForm.Models
{
    public class ProfileEntry
    {
        // アプリ識別
        public string ProcessPath { get; set; } = string.Empty;    // 実行ファイルのフルパス
        public string ProcessName { get; set; } = string.Empty;    // プロセス名

        // ファイル指定
        public string? FilePath { get; set; }                      // 開くファイルのパスまたはURL（nullならアプリ起動のみ）
        public string[]? FilePaths { get; set; }                   // ブラウザ用：複数URLの配列

        // 位置指定
        public int? Monitor { get; set; }                          // 配置先モニター番号（1始まり、nullは位置指定なし）
        public int? Ratio { get; set; }                            // 比率値（合計10の中の割り当て）
        public int? Order { get; set; }                            // 同一モニター内の左からの順番（1始まり）
        public string PositionMode { get; set; } = "none";         // "ratio" または "none"
    }
}
```

### Profile.cs（Models/）

プロファイル全体。

```csharp
namespace StartForm.Models
{
    public class Profile
    {
        public string ProfileName { get; set; } = string.Empty;    // プロファイル名（利用者が命名）
        public DateTime CreatedAt { get; set; } = DateTime.Now;     // 作成日時
        public DateTime UpdatedAt { get; set; } = DateTime.Now;     // 最終更新日時
        public List<ProfileEntry> Entries { get; set; } = new();    // アプリ一覧
    }
}
```

### LastPosition.cs（Models/）

前回位置記録の1行。

```csharp
namespace StartForm.Models
{
    public class LastPosition
    {
        public string ProcessName { get; set; } = string.Empty;
        public string ProcessPath { get; set; } = string.Empty;
        public int PosX { get; set; }
        public int PosY { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string WindowState { get; set; } = "Normal";        // Normal / Maximized / Minimized
    }
}
```

---

## サービス層

### ProfileService.cs（Services/）

プロファイルのJSON読み書きを担当する。

```
保存先ディレクトリ：
  Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
  + "/StartForm/profiles/"

ファイル名：
  {profileName}.json
  ※ファイル名に使えない文字はアンダースコアに置換する

機能：
  - Save(Profile profile)         → JSONに変換して保存。UpdatedAtを現在時刻に更新。
  - Load(string profileName)      → JSONから読み込んでProfileオブジェクトを返す。
  - LoadAll()                     → profiles/ディレクトリ内の全JSONを読み込み、List<Profile>を返す。
  - Delete(string profileName)    → 対象のJSONファイルを削除する。
  - GetProfileDirectory()         → 保存先ディレクトリのパスを返す。ディレクトリが存在しなければ作成する。

JSON設定：
  - System.Text.Json.JsonSerializerOptionsを使用
  - WriteIndented = true（人が読めるように）
  - PropertyNamingPolicy = JsonNamingPolicy.CamelCase（JSON側はcamelCase）
```

### LastPositionService.cs（Services/）

前回位置記録のJSON読み書きを担当する。

```
保存先ディレクトリ：
  Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
  + "/StartForm/last_positions/"

ファイル名：
  {profileName}_positions.json

機能：
  - Save(string profileName, List<LastPosition> positions)  → 保存
  - Load(string profileName)                                → 読み込み。ファイルがなければ空リストを返す。

JSON設定：
  - ProfileServiceと同じ設定
```

---

## Program.cs

```csharp
namespace StartForm
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new Forms.MainForm());
        }
    }
}
```

---

## MainForm.cs（Forms/）

最小限の空フォーム。

```
- タイトル：「StartForm」
- サイズ：幅600 × 高さ400
- StartPosition：CenterScreen
- 中身は空（第2フェーズでUIを構築する）
```

---

## 完了条件

以下が全て満たされたら第1フェーズ完了：

1. `dotnet build` がエラーなく通ること
2. 上記のフォルダ構成通りにファイルが配置されていること
3. 各クラスのプロパティが制作仕様書v1.1のデータ設計と一致していること
4. ProfileServiceのSave/Load/LoadAll/Deleteが実装されていること
5. LastPositionServiceのSave/Loadが実装されていること
6. アプリを起動すると「StartForm」タイトルの空ウィンドウが表示されること

---

## 完了後の報告

作業完了したら、以下をレポートとして出力してください：

- 作成したファイルの一覧
- 各ファイルの内容の概要（何を実装したか）
- dotnet buildの結果
- 指示書との差異がある場合はその理由
