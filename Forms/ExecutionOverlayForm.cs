namespace StartForm.Forms
{
    internal sealed class ExecutionOverlayForm : Form
    {
        public ExecutionOverlayForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.Black;
            Opacity = 0.68;
            Bounds = SystemInformation.VirtualScreen;
        }
    }
}
