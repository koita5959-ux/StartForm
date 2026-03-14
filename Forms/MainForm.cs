using StartForm.Services;

namespace StartForm.Forms
{
    public class MainForm : Form
    {
        private readonly ProfileService _profileService = new();
        private bool _isExecuting;
        private CancellationTokenSource? _executionCts;

        private ListBox lstProfiles = null!;
        private Button btnExecute = null!;
        private Button btnNew = null!;
        private Button btnEdit = null!;
        private Button btnDelete = null!;

        public MainForm()
        {
            Text = "StartForm";
            Width = 600;
            Height = 450;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(500, 350);
            FormBorderStyle = FormBorderStyle.Sizable;

            InitializeControls();
            RefreshProfileList();
        }

        private void InitializeControls()
        {
            // ボタン行の基準位置
            int buttonHeight = 35;
            int buttonWidth = 100;
            int buttonMarginBottom = 20;
            int buttonMarginLeft = 20;
            int buttonSpacing = 10;
            int buttonTop = ClientSize.Height - buttonHeight - buttonMarginBottom;

            // プロファイル一覧（ListBox）
            lstProfiles = new ListBox
            {
                Location = new Point(20, 20),
                Size = new Size(ClientSize.Width - 40, buttonTop - 20 - 10),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                SelectionMode = SelectionMode.One,
                Font = new Font("メイリオ", 10f)
            };
            lstProfiles.SelectedIndexChanged += LstProfiles_SelectedIndexChanged;
            Controls.Add(lstProfiles);

            // ボタン共通フォント
            var buttonFont = new Font("メイリオ", 9f);

            // 実行ボタン
            btnExecute = new Button
            {
                Text = "実行",
                Location = new Point(buttonMarginLeft, buttonTop),
                Size = new Size(buttonWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Font = buttonFont,
                Enabled = false
            };
            btnExecute.Click += BtnExecute_Click;
            Controls.Add(btnExecute);

            // 新規作成ボタン
            btnNew = new Button
            {
                Text = "新規作成",
                Location = new Point(buttonMarginLeft + (buttonWidth + buttonSpacing) * 1, buttonTop),
                Size = new Size(buttonWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Font = buttonFont,
                Enabled = true
            };
            btnNew.Click += BtnNew_Click;
            Controls.Add(btnNew);

            // 編集ボタン
            btnEdit = new Button
            {
                Text = "編集",
                Location = new Point(buttonMarginLeft + (buttonWidth + buttonSpacing) * 2, buttonTop),
                Size = new Size(buttonWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Font = buttonFont,
                Enabled = false
            };
            btnEdit.Click += BtnEdit_Click;
            Controls.Add(btnEdit);

            // 削除ボタン
            btnDelete = new Button
            {
                Text = "削除",
                Location = new Point(buttonMarginLeft + (buttonWidth + buttonSpacing) * 3, buttonTop),
                Size = new Size(buttonWidth, buttonHeight),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Font = buttonFont,
                Enabled = false
            };
            btnDelete.Click += BtnDelete_Click;
            Controls.Add(btnDelete);
        }

        private void LstProfiles_SelectedIndexChanged(object? sender, EventArgs e)
        {
            bool hasSelection = lstProfiles.SelectedIndex >= 0;
            btnExecute.Enabled = hasSelection;
            btnEdit.Enabled = hasSelection;
            btnDelete.Enabled = hasSelection;
        }

        private async void BtnExecute_Click(object? sender, EventArgs e)
        {
            if (_isExecuting)
                return;

            if (lstProfiles.SelectedItem is not string profileName)
                return;

            var profile = _profileService.Load(profileName);
            if (profile == null)
                return;

            // 実行中はボタンを無効化
            _isExecuting = true;
            btnExecute.Enabled = false;
            btnNew.Enabled = false;
            btnEdit.Enabled = false;
            btnDelete.Enabled = false;
            _executionCts = new CancellationTokenSource();

            using var overlay = new ExecutionOverlayForm();
            using var overlayStatus = new ExecutionOverlayStatusForm();
            overlayStatus.CancelRequested += (_, _) => _executionCts.Cancel();
            overlay.Show(this);
            overlayStatus.Show(this);
            overlayStatus.BringToFront();
            await Task.Yield();

            try
            {
                var executor = new ProfileExecutor(new LastPositionService());
                await executor.ExecuteAsync(profile, _executionCts.Token);
            }
            catch (OperationCanceledException)
            {
                // 中断は想定内なのでメッセージ表示しない
            }
            catch (Exception ex)
            {
                MessageBox.Show($"実行中にエラーが発生しました：\n{ex.Message}", "StartForm", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                overlayStatus.Close();
                overlay.Close();
                _executionCts.Dispose();
                _executionCts = null;
                _isExecuting = false;
                // ボタンの有効/無効を復帰
                bool hasSelection = lstProfiles.SelectedIndex >= 0;
                btnExecute.Enabled = hasSelection;
                btnEdit.Enabled = hasSelection;
                btnDelete.Enabled = hasSelection;
                btnNew.Enabled = true;
            }
        }

        private void BtnNew_Click(object? sender, EventArgs e)
        {
            using var editForm = new EditForm();
            if (editForm.ShowDialog(this) == DialogResult.OK)
            {
                RefreshProfileList();
            }
        }

        private void BtnEdit_Click(object? sender, EventArgs e)
        {
            if (lstProfiles.SelectedItem is not string profileName)
                return;

            var profile = _profileService.Load(profileName);
            if (profile == null)
                return;

            using var editForm = new EditForm(profile);
            if (editForm.ShowDialog(this) == DialogResult.OK)
            {
                RefreshProfileList();
            }
        }

        private void BtnDelete_Click(object? sender, EventArgs e)
        {
            if (lstProfiles.SelectedItem is not string profileName)
                return;

            var result = MessageBox.Show(
                $"「{profileName}」を削除しますか？",
                "StartForm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                _profileService.Delete(profileName);
                RefreshProfileList();
            }
        }

        private void RefreshProfileList()
        {
            var profiles = _profileService.LoadAll();
            lstProfiles.Items.Clear();

            foreach (var profile in profiles)
            {
                lstProfiles.Items.Add(profile.ProfileName);
            }

            // ボタンの有効/無効を更新
            bool hasSelection = lstProfiles.SelectedIndex >= 0;
            btnExecute.Enabled = hasSelection;
            btnEdit.Enabled = hasSelection;
            btnDelete.Enabled = hasSelection;
        }
    }
}
