using StartForm.Helpers;

namespace StartForm.Forms
{
    internal sealed class ExecutionOverlayStatusForm : Form
    {
        private readonly Button _cancelButton;
        private readonly Label _titleLabel;

        public event EventHandler? CancelRequested;

        public ExecutionOverlayStatusForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.FromArgb(24, 24, 24);
            Size = new Size(280, 120);

            _titleLabel = new Label
            {
                AutoSize = false,
                Text = "復元中",
                ForeColor = Color.White,
                Font = new Font("メイリオ", 18f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.Transparent,
                Size = new Size(220, 42),
                Location = new Point(30, 18)
            };

            _cancelButton = new Button
            {
                Text = "中断",
                Size = new Size(132, 40),
                Font = new Font("メイリオ", 9f),
                BackColor = Color.FromArgb(244, 244, 244),
                ForeColor = Color.Black,
                UseVisualStyleBackColor = false,
                FlatStyle = FlatStyle.Flat,
                Location = new Point(74, 66)
            };
            _cancelButton.FlatAppearance.BorderSize = 0;
            _cancelButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(225, 225, 225);
            _cancelButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(235, 235, 235);
            _cancelButton.Click += (_, _) =>
            {
                ExecutionLogger.Warn("Execution cancel requested by user");
                SetCancelingState();
                CancelRequested?.Invoke(this, EventArgs.Empty);
            };

            Controls.Add(_cancelButton);
            Controls.Add(_titleLabel);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            CenterToVirtualScreen();
        }

        public void SetCancelingState()
        {
            if (InvokeRequired)
            {
                BeginInvoke(SetCancelingState);
                return;
            }

            _titleLabel.Text = "中断中";
            _cancelButton.Enabled = false;
        }

        private void CenterToVirtualScreen()
        {
            var bounds = SystemInformation.VirtualScreen;
            Location = new Point(
                bounds.Left + (bounds.Width - Width) / 2,
                bounds.Top + (bounds.Height - Height) / 2);
        }
    }
}
