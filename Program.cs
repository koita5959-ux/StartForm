namespace StartForm
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Helpers.ExecutionLogger.Initialize();
            Helpers.ExecutionLogger.Info("Application startup");

            Application.ThreadException += (_, e) =>
                Helpers.ExecutionLogger.Error("Unhandled UI exception", e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (_, e) =>
                Helpers.ExecutionLogger.Error("Unhandled non-UI exception", e.ExceptionObject as Exception);

            const string mutexName = "StartForm_SingleInstance_Mutex";
            using var mutex = new Mutex(true, mutexName, out bool createdNew);

            if (!createdNew)
            {
                Helpers.ExecutionLogger.Warn("Second instance launch blocked");
                MessageBox.Show("StartFormは既に起動しています。", "StartForm",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ApplicationConfiguration.Initialize();
            Helpers.ExecutionLogger.Info("Main form starting");
            Application.Run(new Forms.MainForm());
        }
    }
}
