using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;

namespace RAGSearch
{
    /// <summary>
    /// Very early, fail-safe diagnostics for the VSTO startup path.  Each write
    /// opens and closes the file so the last BEGIN/END pair survives a blocked
    /// COM call, an Outlook crash, or a prompt displayed during startup.
    /// </summary>
    internal static class StartupTrace
    {
        private static readonly object Gate = new object();
        private static readonly string TracePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RAGSearch",
            "startup-trace.log");
        private static bool sessionStarted;

        public static void BeginSession()
        {
            lock (Gate)
            {
                if (sessionStarted)
                {
                    return;
                }

                sessionStarted = true;
                try
                {
                    var assembly = Assembly.GetExecutingAssembly();
                    var assemblyPath = assembly.Location;
                    var assemblyFile = new FileInfo(assemblyPath);
                    AppendUnsafe("=== Outlook add-in startup ===");
                    AppendUnsafe(string.Format(
                        "assembly path={0}; version={1}; mvid={2}; bytes={3}; write_utc={4:o}",
                        assemblyPath,
                        assembly.GetName().Version,
                        assembly.ManifestModule.ModuleVersionId,
                        assemblyFile.Exists ? assemblyFile.Length : -1,
                        assemblyFile.Exists ? assemblyFile.LastWriteTimeUtc : DateTime.MinValue));
                    using (var process = Process.GetCurrentProcess())
                    {
                        AppendUnsafe(string.Format(
                            "process pid={0}; start_utc={1:o}; clr={2}",
                            process.Id,
                            process.StartTime.ToUniversalTime(),
                            Environment.Version));
                    }
                }
                catch
                {
                    // A diagnostic failure must never prevent Outlook startup.
                }
            }
        }

        public static void Step(string message)
        {
            lock (Gate)
            {
                try
                {
                    AppendUnsafe(message);
                }
                catch
                {
                    // A diagnostic failure must never prevent Outlook startup.
                }
            }
        }

        public static void Failure(string step, Exception exception)
        {
            var error = exception == null
                ? "unknown exception"
                : exception.GetType().FullName + ": " + exception.Message;
            Step("FAIL " + step + "; " + SingleLine(error));
        }

        private static void AppendUnsafe(string message)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(TracePath));
            using (var process = Process.GetCurrentProcess())
            {
                File.AppendAllText(
                    TracePath,
                    string.Format(
                        "{0:o} pid={1} tid={2} {3}{4}",
                        DateTime.UtcNow,
                        process.Id,
                        Thread.CurrentThread.ManagedThreadId,
                        SingleLine(message),
                        Environment.NewLine));
            }
        }

        private static string SingleLine(string value)
        {
            return (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ");
        }
    }
}
