# 制作指示書：第3フェーズ - プロファイル編集画面

## この指示書について

C#開発（司令塔）からClaudeCode（実行者）への制作指示です。
第2フェーズ（メインウィンドウUI）が完了したことを前提に、プロファイル編集画面を構築します。

**まず本指示書を読み、把握レポートを出力してください。コード生成はまだ行わないでください。**

---

## 第3フェーズの範囲

### 作るもの
- Forms/EditForm.cs（プロファイル編集画面）
- DataGridViewによるプロファイルエントリの一覧編集
- プロファイル名の入力
- 行の追加・削除
- 保存機能（比率合計チェック付き）
- キャンセル機能
- MainForm.cs の「新規作成」「編集」ボタンからEditFormを呼び出す連携

### 作らないもの（次フェーズ以降）
- 「現在を拾う」機能
- ファイル参照ダイアログ（テキスト入力で手動指定）
- プロファイル実行機能
- Win32 API呼び出し

---

## EditForm.cs の設計

### ウィンドウ基本設定
- タイトル：「プロファイル編集 - StartForm」（新規作成時）/ 「プロファイル編集 - {プロファイル名} - StartForm」（編集時）
- サイズ：幅900 × 高さ550
- StartPosition：CenterParent
- MinimumSize：幅750 × 高さ400
- FormBorderStyle：Sizable

### コンストラクタ

2つのモードをサポートする：

```csharp
// 新規作成モード
public EditForm()

// 編集モード（既存プロファイルを渡す）
public EditForm(Profile existingProfile)
```

編集モードでは、渡されたProfileの内容をフォームに反映する。

### コントロール配置

```
┌───────────────────────────────────────────────────────────────┐
│  プロファイル編集 - StartForm                         [_][□][×] │
├───────────────────────────────────────────────────────────────┤
│                                                               │
│  プロファイル名： [________________________]                    │
│                                                               │
│  ┌───────────────────────────────────────────────────────┐   │
│  │ アプリ(パス) │ ファイル/URL │ 画面 │ 比率 │ 順番 │ モード │   │
│  │──────────────┼─────────────┼──────┼──────┼──────┼────────│   │
│  │ notepad.exe  │ C:\test.txt │  1   │  4   │  1   │ ratio  │   │
│  │ chrome.exe   │ https://... │  1   │  6   │  2   │ ratio  │   │
│  │ explorer.exe │ C:\Work     │      │      │      │ none   │   │
│  │              │             │      │      │      │        │   │
│  └───────────────────────────────────────────────────────┘   │
│                                                               │
│  [ 行を追加 ]  [ 行を削除 ]              [ 保存 ]  [ キャンセル ] │
│                                                               │
└───────────────────────────────────────────────────────────────┘
```

### 各コントロールの仕様

#### プロファイル名
- ラベル：「プロファイル名：」（Label）
- テキストボックス：txtProfileName
- 位置：上部、左マージン20px、上マージン15px
- 幅：300px
- Font：メイリオ, 10pt

#### データグリッド（DataGridView）
- 変数名：dgvEntries
- Anchor：Top, Left, Right, Bottom（リサイズ追従）
- 位置：プロファイル名の下、左右マージン20px
- AllowUserToAddRows：false（行追加はボタンで行う）
- AllowUserToDeleteRows：false（行削除はボタンで行う）
- SelectionMode：FullRowSelect
- MultiSelect：false
- Font：メイリオ, 9pt
- AutoSizeColumnsMode：Fill（列幅を自動調整）
- RowHeadersVisible：false

#### 列定義

| 列名 | ヘッダーテキスト | 型 | 幅の比率 | 説明 |
|------|----------------|-----|---------|------|
| colProcessPath | アプリ（パス） | TextBox | 25% | 実行ファイルのフルパス |
| colFilePath | ファイル/URL | TextBox | 25% | 開くファイルまたはURL |
| colMonitor | 画面 | TextBox | 8% | モニター番号（1始まり、空欄可） |
| colRatio | 比率 | TextBox | 8% | 比率値（空欄可） |
| colOrder | 順番 | TextBox | 8% | 左からの順番（空欄可） |
| colPositionMode | モード | ComboBox | 12% | "ratio" または "none" |

- colPositionModeのComboBox選択肢：「ratio」「none」
- colPositionModeのデフォルト値：「none」

#### ボタン行

左側グループ：
| ボタン | 変数名 | テキスト | 幅 |
|--------|--------|---------|-----|
| 行を追加 | btnAddRow | 行を追加 | 100 |
| 行を削除 | btnDeleteRow | 行を削除 | 100 |

右側グループ（Anchor: Bottom, Right）：
| ボタン | 変数名 | テキスト | 幅 |
|--------|--------|---------|-----|
| 保存 | btnSave | 保存 | 100 |
| キャンセル | btnCancel | キャンセル | 100 |

- ボタンの高さ：35px
- Font：メイリオ, 9pt

---

## 処理の実装

### 1. フォーム初期化

```
新規作成モード：
  - プロファイル名は空
  - DataGridViewは空（行なし）

編集モード：
  - プロファイル名にexistingProfile.ProfileNameを設定
  - DataGridViewにexistingProfile.Entriesの内容を行として表示
  - 元のプロファイル名を内部変数に保持（名前変更時にファイル名を更新するため）
```

### 2. 行を追加（btnAddRow.Click）

```
処理：
  - DataGridViewに空の行を1行追加
  - colPositionModeのデフォルト値は「none」
```

### 3. 行を削除（btnDeleteRow.Click）

```
処理：
  - 選択されている行を削除
  - 行が選択されていなければ何もしない
```

### 4. 保存（btnSave.Click）

```
処理：
  1. バリデーション
     a. プロファイル名が空でないことを確認
        → 空なら MessageBox "プロファイル名を入力してください"
     b. 行が1つ以上あることを確認
        → なければ MessageBox "アプリを1つ以上追加してください"
     c. 比率の合計チェック（後述）

  2. DataGridViewの内容からProfileオブジェクトを構築
     - 各行のセル値をProfileEntryに変換
     - monitor, ratio, order は空欄ならnull
     - monitor, ratio, order に数値以外が入っていたらエラー表示

  3. ProfileService.Save(profile) で保存

  4. 編集モードで名前が変更された場合：
     - 旧名のファイルを ProfileService.Delete(旧名) で削除

  5. DialogResult = DialogResult.OK で画面を閉じる
```

### 5. 比率の合計チェック

```
ルール：
  - positionMode が "ratio" の行について、monitor番号ごとにグループ化
  - 各グループの ratio の合計が 10 であることを確認
  - 合計が10でないグループがあれば：
    MessageBox "画面{monitor番号}の比率の合計が{合計値}です。合計は10にしてください。"
    → 保存を中止

例：
  画面1: 比率4 + 比率6 = 10 → OK
  画面2: 比率3 + 比率5 = 8  → NG "画面2の比率の合計が8です。合計は10にしてください。"
```

### 6. キャンセル（btnCancel.Click）

```
処理：
  - DialogResult = DialogResult.Cancel で画面を閉じる
```

---

## MainForm.cs の変更

### 新規作成ボタンの変更

```
現在：MessageBox "未実装"
変更後：
  1. EditForm editForm = new EditForm() で新規作成モードを開く
  2. editForm.ShowDialog(this) でモーダル表示
  3. DialogResult.OK が返ってきたら RefreshProfileList() で一覧を更新
```

### 編集ボタンの変更

```
現在：MessageBox "未実装"
変更後：
  1. 選択中のプロファイル名を取得
  2. ProfileService.Load(profileName) でプロファイルを読み込む
  3. EditForm editForm = new EditForm(profile) で編集モードを開く
  4. editForm.ShowDialog(this) でモーダル表示
  5. DialogResult.OK が返ってきたら RefreshProfileList() で一覧を更新
```

---

## 変更・追加するファイル

| ファイル | 内容 |
|---------|------|
| Forms/EditForm.cs | 新規作成 |
| Forms/MainForm.cs | 新規作成・編集ボタンの処理を変更 |

---

## 完了条件

1. `dotnet build` がエラーなく通ること
2. メインウィンドウの「新規作成」ボタンでプロファイル編集画面が開くこと
3. プロファイル名を入力し、行を追加して保存できること
4. 保存したプロファイルがメインウィンドウの一覧に表示されること
5. 一覧からプロファイルを選んで「編集」で編集画面が開き、内容が反映されていること
6. 比率の合計が10でない場合にエラーメッセージが出ること
7. プロファイル名が空の場合にエラーメッセージが出ること

---

## 完了後の報告

作業完了したら、以下をレポートとして出力してください：

- 変更・追加したファイルの一覧と変更内容の概要
- dotnet buildの結果
- 指示書との差異がある場合はその理由
