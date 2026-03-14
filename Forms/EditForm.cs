using System.Diagnostics;
using StartForm.Models;
using StartForm.Services;

namespace StartForm.Forms
{
    public class EditForm : Form
    {
        private readonly ProfileService _profileService = new();
        private readonly string? _originalProfileName;

        private TextBox txtProfileName = null!;
        private DataGridView dgvEntries = null!;
        private Button btnCapture = null!;
        private Button btnAddRow = null!;
        private Button btnDeleteRow = null!;
        private Button btnSave = null!;
        private Button btnCancel = null!;

        // 均等割りスイッチ
        private CheckBox chkEvenSplit = null!;

        // 新規作成モード
        public EditForm()
        {
            Text = "プロファイル編集 - StartForm [v3]";
            InitializeForm();
        }

        // 編集モード
        public EditForm(Profile existingProfile)
        {
            _originalProfileName = existingProfile.ProfileName;
            Text = $"プロファイル編集 - {existingProfile.ProfileName} - StartForm [v3]";
            InitializeForm();
            LoadProfile(existingProfile);
        }

        private void InitializeForm()
        {
            Width = 1200;
            Height = 550;
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(1000, 400);
            FormBorderStyle = FormBorderStyle.Sizable;

            InitializeControls();

            // フォームへのドロップを有効化
            this.AllowDrop = true;
            this.DragEnter += Form_DragEnter;
            this.DragDrop += Form_DragDrop;
        }

        private void InitializeControls()
        {
            int buttonHeight = 35;
            int buttonWidth = 100;
            int buttonSpacing = 10;
            int bottomMargin = 20;
            int sideMargin = 20;
            int buttonTop = ClientSize.Height - buttonHeight - bottomMargin;

            var buttonFont = new Font("メイリオ", 9f);

            // プロファイル名ラベル
            var lblProfileName = new Label
            {
                Text = "プロファイル名：",
                Location = new Point(sideMargin, 18),
                AutoSize = true,
                Font = new Font("メイリオ", 10f)
            };
            Controls.Add(lblProfileName);

            // プロファイル名テキストボックス
            txtProfileName = new TextBox
            {
                Location = new Point(sideMargin + lblProfileName.PreferredWidth + 5, 15),
                Width = 300,
                Font = new Font("メイリオ", 10f)
            };
            Controls.Add(txtProfileName);

            // 均等割りチェックボックス（右上）
            chkEvenSplit = new CheckBox
            {
                Text = "比率指定を均等割り",
                Checked = true,
                AutoSize = true,
                Font = new Font("メイリオ", 9f),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            // 位置はOnLoadで調整（ClientSizeが確定してから）
            chkEvenSplit.CheckedChanged += ChkEvenSplit_CheckedChanged;
            Controls.Add(chkEvenSplit);

            int gridTop = txtProfileName.Bottom + 15;

            // DataGridView
            dgvEntries = new DataGridView
            {
                Location = new Point(sideMargin, gridTop),
                Size = new Size(ClientSize.Width - sideMargin * 2, buttonTop - gridTop - 10),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                Font = new Font("メイリオ", 9f),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                AllowDrop = true
            };

            // 列定義
            var colIsActive = new DataGridViewCheckBoxColumn
            {
                Name = "colIsActive",
                HeaderText = "起動",
                FillWeight = 5,
                TrueValue = true,
                FalseValue = false
            };

            var colProcessPath = new DataGridViewTextBoxColumn
            {
                Name = "colProcessPath",
                HeaderText = "アプリ（パス）",
                FillWeight = 28
            };

            var colFilePath = new DataGridViewTextBoxColumn
            {
                Name = "colFilePath",
                HeaderText = "ファイル/URL",
                FillWeight = 24
            };

            var colMonitor = new DataGridViewTextBoxColumn
            {
                Name = "colMonitor",
                HeaderText = "画面",
                FillWeight = 6
            };

            var colPositionMode = new DataGridViewComboBoxColumn
            {
                Name = "colPositionMode",
                HeaderText = "配置",
                FillWeight = 9,
                Items = { "比率指定", "最小化", "位置指定なし" },
                DefaultCellStyle = new DataGridViewCellStyle { NullValue = "位置指定なし" }
            };

            var colOrder = new DataGridViewTextBoxColumn
            {
                Name = "colOrder",
                HeaderText = "順番",
                FillWeight = 6
            };

            var colRatio = new DataGridViewTextBoxColumn
            {
                Name = "colRatio",
                HeaderText = "比率",
                FillWeight = 6,
                ReadOnly = true  // 均等割りON時はReadOnly
            };

            var colZOrder = new DataGridViewTextBoxColumn
            {
                Name = "colZOrder",
                HeaderText = "Z順",
                FillWeight = 5
            };

            var colIsVisible = new DataGridViewTextBoxColumn
            {
                Name = "colIsVisible",
                HeaderText = "見え方",
                FillWeight = 6,
                ReadOnly = true
            };

            var colWindowState = new DataGridViewTextBoxColumn
            {
                Name = "colWindowState",
                HeaderText = "状態",
                FillWeight = 7,
                ReadOnly = true
            };

            var colCapturedCoords = new DataGridViewTextBoxColumn
            {
                Name = "colCapturedCoords",
                HeaderText = "座標W",
                FillWeight = 12,
                ReadOnly = true
            };

            dgvEntries.Columns.AddRange(colIsActive, colProcessPath, colFilePath, colMonitor, colPositionMode, colOrder, colRatio, colZOrder, colIsVisible, colWindowState, colCapturedCoords);

            dgvEntries.CellValueChanged += DgvEntries_CellValueChanged;
            dgvEntries.CurrentCellDirtyStateChanged += DgvEntries_CurrentCellDirtyStateChanged;

            // グリッドへのドロップ
            dgvEntries.DragEnter += Form_DragEnter;
            dgvEntries.DragDrop += DgvEntries_DragDrop;

            Controls.Add(dgvEntries);

            // ボタン群（下部）
            int captureWidth = 120;
            btnCapture = new Button
            {
                Text = "状況抽出",
                Location = new Point(sideMargin, buttonTop),
                Size = new Size(captureWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Font = buttonFont
            };
            btnCapture.Click += BtnCapture_Click;
            Controls.Add(btnCapture);

            btnAddRow = new Button
            {
                Text = "行を追加",
                Location = new Point(sideMargin + captureWidth + buttonSpacing, buttonTop),
                Size = new Size(buttonWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Font = buttonFont
            };
            btnAddRow.Click += BtnAddRow_Click;
            Controls.Add(btnAddRow);

            btnDeleteRow = new Button
            {
                Text = "行を削除",
                Location = new Point(sideMargin + captureWidth + buttonSpacing + buttonWidth + buttonSpacing, buttonTop),
                Size = new Size(buttonWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Font = buttonFont
            };
            btnDeleteRow.Click += BtnDeleteRow_Click;
            Controls.Add(btnDeleteRow);

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(ClientSize.Width - sideMargin - buttonWidth, buttonTop),
                Size = new Size(buttonWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Font = buttonFont
            };
            btnCancel.Click += BtnCancel_Click;
            Controls.Add(btnCancel);

            btnSave = new Button
            {
                Text = "保存",
                Location = new Point(ClientSize.Width - sideMargin - buttonWidth * 2 - buttonSpacing, buttonTop),
                Size = new Size(buttonWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Font = buttonFont
            };
            btnSave.Click += BtnSave_Click;
            Controls.Add(btnSave);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // 均等割りチェックボックスを右上に配置
            chkEvenSplit.Location = new Point(
                ClientSize.Width - chkEvenSplit.PreferredSize.Width - 20,
                15);
            UpdateRatioColumnState();
        }

        // 均等割りチェック変更時：比率列の活性/非活性を切り替え
        private void ChkEvenSplit_CheckedChanged(object? sender, EventArgs e)
        {
            UpdateRatioColumnState();
        }

        private void UpdateRatioColumnState()
        {
            bool evenSplit = chkEvenSplit.Checked;
            var ratioCol = dgvEntries.Columns["colRatio"];
            if (ratioCol == null) return;

            ratioCol.ReadOnly = evenSplit;
            ratioCol.DefaultCellStyle.BackColor = evenSplit
                ? Color.FromArgb(180, 180, 180)
                : SystemColors.Window;
            ratioCol.DefaultCellStyle.ForeColor = evenSplit
                ? Color.FromArgb(100, 100, 100)
                : SystemColors.WindowText;

            dgvEntries.Invalidate();
        }

        // ComboBoxの変更を即座にコミット
        private void DgvEntries_CurrentCellDirtyStateChanged(object? sender, EventArgs e)
        {
            if (dgvEntries.IsCurrentCellDirty)
                dgvEntries.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        // 配置モード変更時の自動処理
        private void DgvEntries_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = dgvEntries.Rows[e.RowIndex];

            if (dgvEntries.Columns[e.ColumnIndex].Name == "colPositionMode")
            {
                var mode = row.Cells["colPositionMode"].Value?.ToString() ?? "位置指定なし";

                if (mode == "比率指定")
                {
                    if (string.IsNullOrWhiteSpace(row.Cells["colMonitor"].Value?.ToString()))
                        row.Cells["colMonitor"].Value = "1";

                    if (string.IsNullOrWhiteSpace(row.Cells["colOrder"].Value?.ToString()))
                        row.Cells["colOrder"].Value = (e.RowIndex + 1).ToString();
                }
                else
                {
                    row.Cells["colOrder"].Value = "";
                    row.Cells["colRatio"].Value = "";
                }
            }
        }

        private void LoadProfile(Profile profile)
        {
            txtProfileName.Text = profile.ProfileName;
            chkEvenSplit.Checked = profile.EvenSplit;

            foreach (var entry in profile.Entries)
            {
                int rowIndex = dgvEntries.Rows.Add();
                var row = dgvEntries.Rows[rowIndex];

                row.Cells["colIsActive"].Value = entry.IsActive;
                row.Cells["colProcessPath"].Value = entry.ProcessPath;
                row.Cells["colFilePath"].Value = entry.FilePath ?? "";
                row.Cells["colMonitor"].Value = entry.Monitor?.ToString() ?? "";
                row.Cells["colPositionMode"].Value = ModeToDisplay(entry.PositionMode);
                row.Cells["colOrder"].Value = entry.Order?.ToString() ?? "";
                row.Cells["colRatio"].Value = entry.Ratio?.ToString() ?? "";
                row.Tag = entry;
                SetInternalDataColumns(row, entry);
            }
        }

        /// <summary>
        /// 内部データ列（Z順・見え方・状態・座標W）をグリッド行にセットする
        /// </summary>
        private static void SetInternalDataColumns(DataGridViewRow row, ProfileEntry entry)
        {
            row.Cells["colZOrder"].Value = entry.ZOrder?.ToString() ?? "";
            row.Cells["colIsVisible"].Value = entry.WindowState == "Minimized" ? "" : (entry.IsVisible ? "前面" : "背面");
            row.Cells["colWindowState"].Value = WindowStateToDisplay(entry.WindowState);
            row.Cells["colCapturedCoords"].Value =
                (entry.CapturedX.HasValue || entry.CapturedY.HasValue || entry.CapturedWidth.HasValue || entry.CapturedHeight.HasValue)
                    ? $"{entry.CapturedX ?? 0},{entry.CapturedY ?? 0},{entry.CapturedWidth ?? 0},{entry.CapturedHeight ?? 0}"
                    : "";
        }

        private void BtnCapture_Click(object? sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "現在のデスクトップ状況を取り込みます。\n既存の設定は置き換えられます。よろしいですか？",
                "StartForm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes) return;

            var captured = ScreenCapturer.CaptureCurrentWindows();

            dgvEntries.Rows.Clear();
            foreach (var entry in captured)
            {
                int rowIndex = dgvEntries.Rows.Add();
                var row = dgvEntries.Rows[rowIndex];

                row.Cells["colIsActive"].Value = entry.IsActive;
                row.Cells["colProcessPath"].Value = entry.ProcessPath;
                row.Cells["colFilePath"].Value = entry.FilePath ?? "";
                row.Cells["colMonitor"].Value = entry.Monitor?.ToString() ?? "";
                row.Cells["colPositionMode"].Value = ModeToDisplay(entry.PositionMode);
                row.Cells["colOrder"].Value = entry.Order?.ToString() ?? "";
                row.Cells["colRatio"].Value = "";
                row.Tag = entry;
                SetInternalDataColumns(row, entry);
            }

            // 状況抽出後は均等割りをOFFにする（行データ流し込み完了後にセット）
            chkEvenSplit.Checked = false;
        }

        // 行を追加：実行中プロセス一覧ダイアログ
        private void BtnAddRow_Click(object? sender, EventArgs e)
        {
            using var dialog = new ProcessSelectDialog();
            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            int rowIndex = dgvEntries.Rows.Add();
            var row = dgvEntries.Rows[rowIndex];
            row.Cells["colIsActive"].Value = true;
            row.Cells["colProcessPath"].Value = dialog.SelectedProcessPath;
            row.Cells["colPositionMode"].Value = "位置指定なし";
            row.Tag = null;
        }

        private void BtnDeleteRow_Click(object? sender, EventArgs e)
        {
            if (dgvEntries.CurrentRow != null && dgvEntries.CurrentRow.Index >= 0)
                dgvEntries.Rows.RemoveAt(dgvEntries.CurrentRow.Index);
        }

        // ドラッグ許可
        private void Form_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        // フォームへのドロップ：選択行のFilePath列にセット
        private void Form_DragDrop(object? sender, DragEventArgs e)
        {
            if (e.Data == null) return;
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;

            SetFilePathToCurrentRow(files[0]);
        }

        // グリッドへのドロップ：ドロップ先の行のFilePath列にセット
        private void DgvEntries_DragDrop(object? sender, DragEventArgs e)
        {
            if (e.Data == null) return;
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;

            // ドロップ位置の行を特定
            var clientPoint = dgvEntries.PointToClient(new Point(e.X, e.Y));
            var hitInfo = dgvEntries.HitTest(clientPoint.X, clientPoint.Y);

            if (hitInfo.RowIndex >= 0)
            {
                dgvEntries.Rows[hitInfo.RowIndex].Cells["colFilePath"].Value = files[0];
            }
            else
            {
                SetFilePathToCurrentRow(files[0]);
            }
        }

        private void SetFilePathToCurrentRow(string path)
        {
            if (dgvEntries.CurrentRow != null && dgvEntries.CurrentRow.Index >= 0)
                dgvEntries.CurrentRow.Cells["colFilePath"].Value = path;
            else if (dgvEntries.Rows.Count > 0)
                dgvEntries.Rows[dgvEntries.Rows.Count - 1].Cells["colFilePath"].Value = path;
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            string profileName = txtProfileName.Text.Trim();
            if (string.IsNullOrEmpty(profileName))
            {
                MessageBox.Show("プロファイル名を入力してください", "StartForm");
                return;
            }

            if (dgvEntries.Rows.Count == 0)
            {
                MessageBox.Show("アプリを1つ以上追加してください", "StartForm");
                return;
            }

            var entries = new List<ProfileEntry>();
            bool evenSplit = chkEvenSplit.Checked;

            // 比率バリデーション（均等割りOFFの場合）
            if (!evenSplit)
            {
                int totalRatio = 0;
                for (int i = 0; i < dgvEntries.Rows.Count; i++)
                {
                    var row = dgvEntries.Rows[i];
                    string mode = DisplayToMode(row.Cells["colPositionMode"].Value?.ToString() ?? "位置指定なし");
                    if (mode != "ratio") continue;

                    string ratioStr = row.Cells["colRatio"].Value?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(ratioStr))
                    {
                        if (!int.TryParse(ratioStr, out int rv) || rv <= 0)
                        {
                            MessageBox.Show($"行{i + 1}の「比率」は1以上の整数で入力してください", "StartForm");
                            return;
                        }
                        totalRatio += rv;
                    }
                }

                if (totalRatio > 10)
                {
                    MessageBox.Show($"前面アプリの比率合計が{totalRatio}になっています。合計10以下にしてください", "StartForm");
                    return;
                }
            }

            for (int i = 0; i < dgvEntries.Rows.Count; i++)
            {
                var row = dgvEntries.Rows[i];
                var originalEntry = row.Tag as ProfileEntry;

                var entry = new ProfileEntry
                {
                    IsActive = row.Cells["colIsActive"].Value is true,
                    ProcessPath = row.Cells["colProcessPath"].Value?.ToString() ?? string.Empty,
                    FilePath = string.IsNullOrWhiteSpace(row.Cells["colFilePath"].Value?.ToString())
                        ? null
                        : row.Cells["colFilePath"].Value!.ToString(),
                    PositionMode = DisplayToMode(row.Cells["colPositionMode"].Value?.ToString() ?? "位置指定なし"),
                    FilePaths = originalEntry?.FilePaths,
                    ChromeProfileDirectory = originalEntry?.ChromeProfileDirectory,
                    ChromeProfileName = originalEntry?.ChromeProfileName,
                    CapturedMonitorDeviceName = originalEntry?.CapturedMonitorDeviceName,
                    CapturedMonitorLeft = originalEntry?.CapturedMonitorLeft,
                    CapturedMonitorTop = originalEntry?.CapturedMonitorTop,
                    CapturedMonitorWidth = originalEntry?.CapturedMonitorWidth,
                    CapturedMonitorHeight = originalEntry?.CapturedMonitorHeight,
                    CapturedWindowTitle = originalEntry?.CapturedWindowTitle,
                    CapturedWindowClassName = originalEntry?.CapturedWindowClassName
                };

                if (!string.IsNullOrEmpty(entry.ProcessPath))
                    entry.ProcessName = Path.GetFileNameWithoutExtension(entry.ProcessPath);

                string monitorStr = row.Cells["colMonitor"].Value?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(monitorStr))
                {
                    if (!int.TryParse(monitorStr, out int monitorVal))
                    {
                        MessageBox.Show($"行{i + 1}の「画面」に数値以外が入力されています", "StartForm");
                        return;
                    }
                    entry.Monitor = monitorVal;
                }

                string orderStr = row.Cells["colOrder"].Value?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(orderStr))
                {
                    if (!int.TryParse(orderStr, out int orderVal))
                    {
                        MessageBox.Show($"行{i + 1}の「順番」に数値以外が入力されています", "StartForm");
                        return;
                    }
                    entry.Order = orderVal;
                }

                // 比率（均等割りOFFかつ前面のみ）
                if (!evenSplit && entry.PositionMode == "ratio")
                {
                    string ratioStr = row.Cells["colRatio"].Value?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(ratioStr) && int.TryParse(ratioStr, out int ratioVal))
                        entry.Ratio = ratioVal;
                }

                if (entry.PositionMode == "ratio")
                {
                    if (!entry.Monitor.HasValue)
                    {
                        MessageBox.Show($"行{i + 1}：比率指定では「画面」の入力が必要です", "StartForm");
                        return;
                    }
                    if (!entry.Order.HasValue)
                    {
                        MessageBox.Show($"行{i + 1}：比率指定では「順番」の入力が必要です", "StartForm");
                        return;
                    }
                }

                // 内部データ（Z順・見え方・状態・座標W）を保存
                string zOrderStr = row.Cells["colZOrder"].Value?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(zOrderStr) && int.TryParse(zOrderStr, out int zOrderVal))
                    entry.ZOrder = zOrderVal;

                string isVisibleStr = row.Cells["colIsVisible"].Value?.ToString() ?? "";
                entry.IsVisible = isVisibleStr != "背面";

                entry.WindowState = DisplayToWindowState(row.Cells["colWindowState"].Value?.ToString() ?? "通常");

                string coordsStr = row.Cells["colCapturedCoords"].Value?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(coordsStr))
                {
                    var parts = coordsStr.Split(',');
                    if (parts.Length == 4
                        && int.TryParse(parts[0], out int cx)
                        && int.TryParse(parts[1], out int cy)
                        && int.TryParse(parts[2], out int cw)
                        && int.TryParse(parts[3], out int ch))
                    {
                        entry.CapturedX = cx;
                        entry.CapturedY = cy;
                        entry.CapturedWidth = cw;
                        entry.CapturedHeight = ch;
                    }
                }

                entries.Add(entry);
            }

            var profile = new Profile
            {
                ProfileName = profileName,
                EvenSplit = evenSplit,
                Entries = entries
            };

            if (_originalProfileName != null && _originalProfileName != profileName)
                _profileService.Delete(_originalProfileName);

            _profileService.Save(profile);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click(object? sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private static string ModeToDisplay(string mode) => mode switch
        {
            "ratio" => "比率指定",
            "minimize" => "最小化",
            "none" => "位置指定なし",
            // 旧値の互換
            "front" => "比率指定",
            "back" => "位置指定なし",
            _ => "位置指定なし"
        };

        private static string DisplayToMode(string display) => display switch
        {
            "比率指定" => "ratio",
            "最小化" => "minimize",
            "位置指定なし" => "none",
            _ => "none"
        };

        private static string WindowStateToDisplay(string windowState) => windowState switch
        {
            "Maximized" => "最大化",
            "Minimized" => "最小化",
            _ => "通常"
        };

        private static string DisplayToWindowState(string display) => display switch
        {
            "最大化" => "Maximized",
            "最小化" => "Minimized",
            "通常" => "Normal",
            "Maximized" => "Maximized",
            "Minimized" => "Minimized",
            _ => "Normal"
        };
    }

    /// <summary>
    /// 実行中プロセス一覧から選択するダイアログ
    /// </summary>
    public class ProcessSelectDialog : Form
    {
        public string SelectedProcessPath { get; private set; } = string.Empty;

        private ListBox lstProcesses = null!;
        private Button btnOk = null!;
        private Button btnCancel = null!;
        private Button btnBrowse = null!;

        // プロセス名 → パスの対応
        private readonly Dictionary<string, string> _processMap = new();

        public ProcessSelectDialog()
        {
            Text = "アプリを選択 - StartForm";
            Width = 520;
            Height = 420;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            InitializeControls();
            LoadProcesses();
        }

        private void InitializeControls()
        {
            var font = new Font("メイリオ", 9f);

            var lbl = new Label
            {
                Text = "現在起動中のアプリ一覧（ダブルクリックで選択）",
                Location = new Point(15, 12),
                AutoSize = true,
                Font = font
            };
            Controls.Add(lbl);

            lstProcesses = new ListBox
            {
                Location = new Point(15, 35),
                Size = new Size(475, 290),
                Font = font,
                Sorted = true
            };
            lstProcesses.DoubleClick += (s, e) => AcceptSelection();
            Controls.Add(lstProcesses);

            btnOk = new Button
            {
                Text = "選択",
                Location = new Point(180, 340),
                Size = new Size(90, 32),
                Font = font,
                DialogResult = DialogResult.None
            };
            btnOk.Click += (s, e) => AcceptSelection();
            Controls.Add(btnOk);

            btnBrowse = new Button
            {
                Text = "参照...",
                Location = new Point(280, 340),
                Size = new Size(90, 32),
                Font = font
            };
            btnBrowse.Click += BtnBrowse_Click;
            Controls.Add(btnBrowse);

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(380, 340),
                Size = new Size(90, 32),
                Font = font,
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(btnCancel);
        }

        private void LoadProcesses()
        {
            _processMap.Clear();
            lstProcesses.Items.Clear();

            var excluded = ScreenCapturer.ExcludedProcesses;

            foreach (var p in Process.GetProcesses())
            {
                try
                {
                    if (excluded.Contains(p.ProcessName, StringComparer.OrdinalIgnoreCase))
                        continue;

                    string path = p.MainModule?.FileName ?? string.Empty;

                    // explorerは特別処理
                    if (p.ProcessName.Equals("explorer", StringComparison.OrdinalIgnoreCase))
                    {
                        path = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                            "explorer.exe");
                    }

                    if (string.IsNullOrEmpty(path)) continue;

                    string exeName = Path.GetFileName(path);
                    string displayName = $"{p.ProcessName}  ({exeName})";

                    if (!_processMap.ContainsKey(displayName))
                    {
                        _processMap[displayName] = path;
                        lstProcesses.Items.Add(displayName);
                    }
                }
                catch { }
            }
        }

        private void AcceptSelection()
        {
            if (lstProcesses.SelectedItem == null)
            {
                MessageBox.Show("アプリを選択してください", "StartForm");
                return;
            }

            string key = lstProcesses.SelectedItem.ToString() ?? "";
            if (_processMap.TryGetValue(key, out string? path))
            {
                SelectedProcessPath = path;
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Title = "アプリを選択",
                Filter = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*",
                CheckFileExists = true
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                SelectedProcessPath = ofd.FileName;
                DialogResult = DialogResult.OK;
                Close();
            }
        }
    }
}
