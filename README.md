# StartForm

`StartForm` は、作業開始時のアプリ起動とウィンドウ配置をまとめて復元する Windows Forms アプリです。

## フォルダ構成

- `Forms/`
  - UI 本体
- `Helpers/`
  - Win32 API 呼び出し、ログ、配置補助
- `Models/`
  - JSON 保存モデル
- `Services/`
  - 抽出、復元、保存処理
- `docs/`
  - 仕様書、共有メモ、作業ドキュメント
- `publish/`
  - `dotnet publish` の出力
- `installer_output/`
  - Inno Setup で生成したインストーラー
- `artifacts/`
  - 完成配布物の zip など

## 主要ファイル

- `StartForm.csproj`
  - プロジェクト定義
- `Program.cs`
  - 起動処理
- `setup.iss`
  - インストーラー定義
- `ReadMe.txt`
  - 利用者向け説明
- `ご利用にあたって.txt`
  - 配布同梱用テキスト

## 現時点の完成配布物

- `installer_output/StartForm_Setup_1.0.0.exe`
- `artifacts/StartForm_Setup_1.0.0.zip`
