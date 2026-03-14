# 制作指示書：第6フェーズ - リリースビルドとインストーラー

## この指示書について

C#開発（司令塔）からClaudeCode（実行者）への制作指示です。
第5フェーズ（仕上げ機能）が完了したことを前提に、配布可能な形に仕上げます。

**まず本指示書を読み、把握レポートを出力してください。コード生成はまだ行わないでください。**

---

## 第6フェーズの範囲

### 作るもの
- dotnet publish によるリリースビルド
- Inno Setup スクリプト（setup.iss）
- ReadMe.txt

### 作らないもの
- カスタムアイコン（将来の課題。デフォルトアイコンで進める）
- テスト実績記録（実機テスト後に作成する別フェーズ）

---

## 1. リリースビルド

### publishコマンド

```
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish
```

このコマンドにより、publish/ フォルダに単一実行ファイル StartForm.exe が生成される。
self-contained により .NET ランタイムが同梱され、利用者のPCに .NET をインストールする必要がない。

### 確認事項
- publish/StartForm.exe が生成されること
- ファイルサイズの確認（self-contained の場合、概ね60-80MB程度）

---

## 2. Inno Setup スクリプト（setup.iss）

StartFormフォルダ直下に setup.iss を作成する。

```iss
; StartForm Inno Setup Script
; 便利アプリシリーズ 第二弾

#define MyAppName "StartForm"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "便利アプリシリーズ"
#define MyAppExeName "StartForm.exe"
#define MyAppDescription "自分が決めた作業環境を、いつでもワンアクションで整える"

[Setup]
AppId={{B8F3D4E2-A1C7-4F5E-9D6B-3E8A2C1F0D5E}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=StartForm_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableProgramGroupPage=yes

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成"; GroupDescription: "追加タスク:"; Flags: unchecked

[Files]
Source: "publish\StartForm.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{group}\{#MyAppName} をアンインストール"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "StartFormを起動する"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\StartForm"
```

### setup.iss のポイント
- **PrivilegesRequired=lowest**：管理者権限不要。利用者のユーザーフォルダにインストール
- **{autopf}**：Program Files（権限に応じて自動選択）
- **OutputDir=installer_output**：生成されたSetup.exeの出力先
- **Japanese.isl**：日本語のインストーラー画面
- **[UninstallDelete]**：アンインストール時に %AppData%/StartForm を完全削除（プロファイルデータも削除）
- **desktopicon**：デスクトップショートカットはオプション（デフォルトではチェックなし）

---

## 3. ReadMe.txt

StartFormフォルダ直下に ReadMe.txt を作成する。配布ZIPに同梱するためのファイル。

```
========================================
  StartForm - ご利用にあたって
  便利アプリシリーズ 第二弾
  バージョン 1.0.0
========================================

■ StartFormとは
  自分が決めた作業環境を、いつでもワンアクションで整えるアプリです。
  プロファイル（作業環境の設計図）を作成し、実行すると、
  指定したアプリが起動し、指定した画面位置に配置されます。

■ 基本的な使い方
  1. StartFormを起動する
  2. 「新規作成」でプロファイルを作成する
     - 「現在を拾う」で今の状態を取り込み、修正して保存できます
     - 手動でアプリやファイルを追加することもできます
  3. プロファイルを選んで「実行」するだけ

■ プロファイルの作り方
  - 「現在を拾う」を使うと、今開いているアプリを一括で取り込めます
  - 取り込んだ後、ファイルパスやURLを手動で追加できます
  - 画面位置は比率で指定します（合計10。左から順に配置）
    例：画面1 = 4:6 → 左40%と右60%に2つのアプリを並べる

■ 画面位置の決め方
  - 「モード」を「ratio」にすると、比率指定で配置されます
  - 「モード」を「none」にすると、起動だけしてアプリに位置を任せます
  - 比率の合計は画面ごとに10になるようにしてください

■ ブラウザのタブ
  - ファイル/URL欄にURLを入力すると、そのURLを開いた状態で起動します
  - 複数タブを開きたい場合は、同じブラウザで複数行を作成してください

■ 対応状況
  - メモ帳、Word（ダブルクリック起動）、Chrome、エクスプローラー等に対応
  - Windows設定（UWPアプリ）は対象外です
  - 詳細はテスト実績記録をご確認ください

■ アンインストール
  - スタートメニュー → StartForm → 「StartFormをアンインストール」
  - または、設定 → アプリ → StartForm → アンインストール
  - アンインストール後、設定データ（プロファイル等）も削除されます

■ 注意事項
  - 一部のアプリはファイルパスを自動取得できません。
    その場合は手動で入力してください。
  - ウィンドウの配置はシャドウ（影）を考慮して補正しています。
  - StartFormは多重起動できません（1つだけ起動できます）。

■ 動作環境
  - Windows 10 / Windows 11
  - .NETランタイムの追加インストールは不要です

========================================
  Copyright © 2026 便利アプリシリーズ
========================================
```

---

## 変更・追加するファイル

| ファイル | 内容 |
|---------|------|
| setup.iss | 新規作成（Inno Setupスクリプト） |
| ReadMe.txt | 新規作成 |

※ dotnet publish はコマンド実行のみ。ソースコードの変更なし。

---

## 完了条件

1. `dotnet publish` が成功し、publish/StartForm.exe が生成されること
2. setup.iss がStartFormフォルダ直下に存在すること
3. ReadMe.txt がStartFormフォルダ直下に存在すること
4. setup.iss の内容が上記の仕様通りであること

### 注意
Inno Setup Compiler によるインストーラー生成（Setup.exeのコンパイル）は、
西村さんのPC上で手動で実行する。ClaudeCodeは setup.iss の作成まで。

---

## 完了後の報告

作業完了したら、以下をレポートとして出力してください：

- 作成したファイルの一覧
- dotnet publish の結果（成功/失敗、出力ファイル）
- 指示書との差異がある場合はその理由
