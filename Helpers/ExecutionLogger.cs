using System.Text;

namespace StartForm.Helpers
{
    internal static class ExecutionLogger
    {
        private static readonly object Sync = new();
        private const int MaxLogFiles = 20;
        private static readonly TimeSpan MaxLogAge = TimeSpan.FromDays(14);
        private static string? _logFilePath;

        public static string LogFilePath => EnsureInitialized();

        public static void Initialize()
        {
            EnsureInitialized();
        }

        public static void Info(string message)
        {
            Write("INFO", message);
        }

        public static void Warn(string message)
        {
            Write("WARN", message);
        }

        public static void Error(string message, Exception? exception = null)
        {
            if (exception != null)
            {
                message = $"{message}{Environment.NewLine}{exception}";
            }

            Write("ERROR", message);
        }

        private static string EnsureInitialized()
        {
            if (!string.IsNullOrEmpty(_logFilePath))
                return _logFilePath;

            lock (Sync)
            {
                if (!string.IsNullOrEmpty(_logFilePath))
                    return _logFilePath;

                var logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "StartForm",
                    "logs");

                Directory.CreateDirectory(logDirectory);
                CleanupOldLogs(logDirectory);

                _logFilePath = Path.Combine(
                    logDirectory,
                    $"startform_{DateTime.Now:yyyyMMdd_HHmmss}_{Environment.ProcessId}.log");

                File.AppendAllText(
                    _logFilePath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [INFO] Log started{Environment.NewLine}",
                    Encoding.UTF8);

                return _logFilePath;
            }
        }

        private static void CleanupOldLogs(string logDirectory)
        {
            try
            {
                var directory = new DirectoryInfo(logDirectory);
                var files = directory.GetFiles("startform_*.log")
                    .OrderByDescending(f => f.LastWriteTimeUtc)
                    .ToList();

                var deleteBefore = DateTime.UtcNow - MaxLogAge;
                foreach (var file in files.Where(f => f.LastWriteTimeUtc < deleteBefore))
                {
                    try
                    {
                        file.Delete();
                    }
                    catch
                    {
                    }
                }

                files = directory.GetFiles("startform_*.log")
                    .OrderByDescending(f => f.LastWriteTimeUtc)
                    .ToList();

                foreach (var file in files.Skip(MaxLogFiles))
                {
                    try
                    {
                        file.Delete();
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        private static void Write(string level, string message)
        {
            var path = EnsureInitialized();
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";

            lock (Sync)
            {
                File.AppendAllText(path, line, Encoding.UTF8);
            }
        }
    }
}
