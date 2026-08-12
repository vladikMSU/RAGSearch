using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace RAGSearch
{
    /// <summary>
    /// Owns one native Extended MAPI import run.  This class deliberately has no
    /// Outlook Object Model dependency: the VSTO pane is only a process/UI host.
    /// </summary>
    internal sealed class OutlookMapiImportRunner : IDisposable
    {
        private const string ProgressPrefix = "RAGSEARCH_PROGRESS ";
#if DEBUG
        private const string ReaderBuildConfiguration = "Debug";
#else
        private const string ReaderBuildConfiguration = "Release";
#endif
        private static readonly TimeSpan HealthProbeTimeout = TimeSpan.FromSeconds(2);
        private static readonly TimeSpan ServiceStartupTimeout = TimeSpan.FromSeconds(90);
        private static readonly TimeSpan GracefulStopTimeout = TimeSpan.FromSeconds(5);

        private readonly object gate = new object();
        private readonly LocalServiceClient serviceClient;
        private CancellationTokenSource runCancellation;
        private Process adapterProcess;
        private Process ownedServiceProcess;
        private OwnedProcessJob adapterJob;
        private string cancelFilePath;
        private bool disposed;
        private bool isRunning;

        public OutlookMapiImportRunner(LocalServiceClient serviceClient)
        {
            this.serviceClient = serviceClient ?? throw new ArgumentNullException("serviceClient");
        }

        public event EventHandler<OutlookMapiImportProgress> ProgressChanged;

        public bool IsRunning
        {
            get
            {
                lock (gate)
                {
                    return isRunning;
                }
            }
        }

        public async Task RunAsync()
        {
            CancellationTokenSource cancellation;
            lock (gate)
            {
                ThrowIfDisposed();
                if (isRunning)
                {
                    throw new InvalidOperationException("Outlook MAPI indexing is already running.");
                }

                isRunning = true;
                cancellation = new CancellationTokenSource();
                runCancellation = cancellation;
            }

            try
            {
                var layout = WorkspaceLayout.Load();
                Report("checking_service", 0, 0, "Проверяю локальный сервис...", true);
                await EnsureServiceAsync(layout, cancellation.Token).ConfigureAwait(false);
                cancellation.Token.ThrowIfCancellationRequested();
                await RunAdapterAsync(layout, cancellation.Token).ConfigureAwait(false);
            }
            finally
            {
                CleanupAdapterState();
                lock (gate)
                {
                    if (ReferenceEquals(runCancellation, cancellation))
                    {
                        runCancellation = null;
                    }
                    isRunning = false;
                }

                cancellation.Dispose();
            }
        }

        public async Task EnsureServiceReadyAsync(CancellationToken cancellationToken)
        {
            lock (gate)
            {
                ThrowIfDisposed();
                if (isRunning)
                {
                    throw new InvalidOperationException(
                        "Нельзя отдельно запускать сервис во время Outlook MAPI-индексации.");
                }
            }

            if (await IsServiceHealthyAsync(cancellationToken).ConfigureAwait(false))
            {
                return;
            }

            var layout = WorkspaceLayout.LoadForService();
            await EnsureServiceAsync(layout, cancellationToken).ConfigureAwait(false);
        }

        public void RequestStop()
        {
            CancellationTokenSource cancellation;
            string sentinel;
            lock (gate)
            {
                cancellation = runCancellation;
                sentinel = cancelFilePath;
            }

            if (cancellation == null)
            {
                return;
            }

            Report("stopping", 0, 0, "Останавливаю Outlook MAPI-индексацию...", true);
            TryCreateCancellationSentinel(sentinel);
            cancellation.Cancel();
        }

        private async Task EnsureServiceAsync(WorkspaceLayout layout, CancellationToken cancellationToken)
        {
            if (await IsServiceHealthyAsync(cancellationToken).ConfigureAwait(false))
            {
                return;
            }

            Report("starting_service", 0, 0, "Python-сервис не отвечает; запускаю локально...", true);
            Process existingOwnedProcess;
            lock (gate)
            {
                existingOwnedProcess = ownedServiceProcess;
            }

            if (existingOwnedProcess == null || HasExited(existingOwnedProcess))
            {
                var arguments = new StringBuilder();
                arguments.Append("-m ragsearch_service");
                arguments.Append(" --port ").Append(serviceClient.ServiceUri.Port);
                arguments.Append(" --delete-spool-after-ingest");
                if (Directory.Exists(layout.EmbeddingModelDirectory))
                {
                    arguments.Append(" --embedding sentence-transformers --model ");
                    arguments.Append(QuoteArgument(layout.EmbeddingModelDirectory));
                }

                var startInfo = HiddenPython(
                    layout.PythonExecutable,
                    arguments.ToString(),
                    layout.ServiceDirectory);
                var serviceProcess = new Process { StartInfo = startInfo, EnableRaisingEvents = true };
                serviceProcess.OutputDataReceived += IgnoreProcessOutput;
                serviceProcess.ErrorDataReceived += IgnoreProcessOutput;
                try
                {
                    if (!serviceProcess.Start())
                    {
                        throw new InvalidOperationException("Windows did not start the Python service process.");
                    }
                    serviceProcess.BeginOutputReadLine();
                    serviceProcess.BeginErrorReadLine();
                }
                catch
                {
                    serviceProcess.Dispose();
                    throw;
                }

                lock (gate)
                {
                    ownedServiceProcess = serviceProcess;
                }
                existingOwnedProcess = serviceProcess;
            }

            var deadline = DateTime.UtcNow + ServiceStartupTimeout;
            while (DateTime.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (await IsServiceHealthyAsync(cancellationToken).ConfigureAwait(false))
                {
                    Report("service_ready", 0, 0, "Локальный сервис запущен.", true);
                    return;
                }
                if (HasExited(existingOwnedProcess))
                {
                    throw new InvalidOperationException(
                        "Python-сервис завершился при запуске. Проверьте модуль ragsearch_service и локальную модель embeddings.");
                }
                await Task.Delay(500, cancellationToken).ConfigureAwait(false);
            }

            throw new TimeoutException(
                "Python-сервис не стал доступен за 90 секунд на " + serviceClient.ServiceUri + ".");
        }

        private async Task<bool> IsServiceHealthyAsync(CancellationToken cancellationToken)
        {
            using (var probe = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                probe.CancelAfter(HealthProbeTimeout);
                try
                {
                    var health = await serviceClient.GetHealthAsync(probe.Token).ConfigureAwait(false);
                    return health != null && string.Equals(health.Status, "ok", StringComparison.Ordinal);
                }
                catch (OperationCanceledException)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return false;
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        private async Task RunAdapterAsync(WorkspaceLayout layout, CancellationToken cancellationToken)
        {
            var sentinel = Path.Combine(
                Path.GetTempPath(),
                "ragsearch-outlook-mapi-cancel-" + Guid.NewGuid().ToString("N") + ".flag");
            lock (gate)
            {
                cancelFilePath = sentinel;
            }

            var arguments = new StringBuilder();
            arguments.Append(QuoteArgument(layout.AdapterScript));
            arguments.Append(" --executable ").Append(QuoteArgument(layout.ReaderExecutable));
            arguments.Append(" --service-url ").Append(QuoteArgument(serviceClient.ServiceUri.GetLeftPart(UriPartial.Authority)));
            // Production indexing must not inherit the adapter's bounded scan defaults.
            arguments.Append(" --full-scan --body-preview-chars 4000000");
            arguments.Append(" --cancel-file ").Append(QuoteArgument(sentinel));

            var process = new Process
            {
                StartInfo = HiddenPython(
                    layout.PythonExecutable,
                    arguments.ToString(),
                    layout.WorkspaceRoot),
                EnableRaisingEvents = true
            };
            var exit = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
            process.OutputDataReceived += AdapterOutputReceived;
            // Drain stderr to avoid a full pipe blocking the adapter.  Its text may
            // contain mailbox metadata, so it is intentionally not logged or shown.
            process.ErrorDataReceived += IgnoreProcessOutput;
            process.Exited += delegate
            {
                try
                {
                    exit.TrySetResult(process.ExitCode);
                }
                catch (InvalidOperationException)
                {
                    exit.TrySetResult(-1);
                }
            };

            var job = new OwnedProcessJob();
            var started = false;
            var assigned = false;
            try
            {
                if (!process.Start())
                {
                    throw new InvalidOperationException("Windows did not start the Outlook MAPI adapter.");
                }
                started = true;
                job.Assign(process);
                assigned = true;
                lock (gate)
                {
                    adapterProcess = process;
                    adapterJob = job;
                }
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
            }
            catch
            {
                if (started && !HasExited(process))
                {
                    try
                    {
                        if (assigned)
                        {
                            job.Terminate();
                        }
                        else
                        {
                            // Assignment happens immediately after Start, before the
                            // adapter can normally create its reader child. If Windows
                            // rejects the Job, stop only this exact process.
                            process.Kill();
                        }
                        process.WaitForExit(3000);
                    }
                    catch (InvalidOperationException)
                    {
                    }
                }
                job.Dispose();
                process.Dispose();
                throw;
            }

            Report("starting", 0, 0, "Запускаю read-only Extended MAPI...", true);
            var cancelled = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            using (cancellationToken.Register(() => cancelled.TrySetResult(true)))
            {
                var completed = await Task.WhenAny(exit.Task, cancelled.Task).ConfigureAwait(false);
                if (completed == cancelled.Task)
                {
                    TryCreateCancellationSentinel(sentinel);
                    var graceful = await Task.WhenAny(
                            exit.Task,
                            Task.Delay(GracefulStopTimeout))
                        .ConfigureAwait(false);
                    if (graceful != exit.Task)
                    {
                        // The Job contains only the adapter process started above and
                        // its OutlookMapiReader child.  No pre-existing process is touched.
                        job.Terminate();
                    }
                    await WaitForExitAfterTerminationAsync(process, exit.Task).ConfigureAwait(false);
                    throw new OperationCanceledException(cancellationToken);
                }
            }

            var exitCode = await exit.Task.ConfigureAwait(false);
            process.WaitForExit(); // Flush asynchronous stdout callbacks.
            if (exitCode == 130 && cancellationToken.IsCancellationRequested)
            {
                throw new OperationCanceledException(cancellationToken);
            }
            if (exitCode != 0)
            {
                throw new InvalidOperationException(
                    "Outlook MAPI adapter завершился с кодом " + exitCode +
                    ". Текст stderr скрыт, чтобы не выводить данные почты; запустите adapter из консоли для диагностики.");
            }
        }

        private static async Task WaitForExitAfterTerminationAsync(Process process, Task<int> exit)
        {
            var completed = await Task.WhenAny(exit, Task.Delay(TimeSpan.FromSeconds(3))).ConfigureAwait(false);
            if (completed == exit)
            {
                return;
            }
            try
            {
                if (!process.HasExited)
                {
                    process.Kill();
                }
            }
            catch (InvalidOperationException)
            {
                // It exited between HasExited and Kill.
            }
        }

        private void AdapterOutputReceived(object sender, DataReceivedEventArgs eventArgs)
        {
            var line = eventArgs.Data;
            if (string.IsNullOrEmpty(line) || !line.StartsWith(ProgressPrefix, StringComparison.Ordinal))
            {
                return;
            }

            AdapterProgress value;
            try
            {
                var json = line.Substring(ProgressPrefix.Length);
                var serializer = new DataContractJsonSerializer(typeof(AdapterProgress));
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    value = (AdapterProgress)serializer.ReadObject(stream);
                }
            }
            catch (Exception)
            {
                return;
            }

            if (value == null || value.Current < 0 || value.Total < 0 || value.BodiesTruncated < 0)
            {
                return;
            }

            string status;
            bool running;
            switch (value.Phase ?? string.Empty)
            {
                case "starting":
                    status = "Запускаю OutlookMapiReader...";
                    running = true;
                    break;
                case "importing":
                    status = "Extended MAPI: проиндексировано " + value.Current + " писем" +
                             TruncatedSuffix(value.BodiesTruncated);
                    running = true;
                    break;
                case "complete":
                    status = "Outlook MAPI-индексация завершена: " + value.Current + " писем" +
                             TruncatedSuffix(value.BodiesTruncated);
                    running = false;
                    break;
                case "cancelled":
                    status = "Outlook MAPI-индексация остановлена: " + value.Current + " писем" +
                             TruncatedSuffix(value.BodiesTruncated);
                    running = false;
                    break;
                case "failed":
                    status = "Outlook MAPI-индексация завершилась с ошибкой.";
                    running = false;
                    break;
                default:
                    return;
            }
            Report(
                value.Phase,
                value.Current,
                value.Total,
                status,
                running);
        }

        private static string TruncatedSuffix(int bodiesTruncated)
        {
            return bodiesTruncated <= 0
                ? string.Empty
                : "; body усечено по лимиту: " + bodiesTruncated;
        }

        private void Report(string phase, int current, int total, string status, bool running)
        {
            var handler = ProgressChanged;
            if (handler != null)
            {
                handler(this, new OutlookMapiImportProgress
                {
                    Phase = phase,
                    Current = current,
                    Total = total,
                    Status = status,
                    IsRunning = running
                });
            }
        }

        private void CleanupAdapterState()
        {
            Process process;
            OwnedProcessJob job;
            string sentinel;
            lock (gate)
            {
                process = adapterProcess;
                job = adapterJob;
                sentinel = cancelFilePath;
                adapterProcess = null;
                adapterJob = null;
                cancelFilePath = null;
            }

            if (process != null)
            {
                process.OutputDataReceived -= AdapterOutputReceived;
                process.ErrorDataReceived -= IgnoreProcessOutput;
                process.Dispose();
            }
            if (job != null)
            {
                job.Dispose();
            }
            TryDelete(sentinel);
        }

        private static ProcessStartInfo HiddenPython(string python, string arguments, string workingDirectory)
        {
            return new ProcessStartInfo
            {
                FileName = python,
                Arguments = arguments,
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
        }

        private static string QuoteArgument(string argument)
        {
            if (string.IsNullOrEmpty(argument))
            {
                return "\"\"";
            }
            if (argument.IndexOfAny(new[] { ' ', '\t', '\n', '\v', '\"' }) < 0)
            {
                return argument;
            }

            var result = new StringBuilder("\"");
            var backslashes = 0;
            foreach (var character in argument)
            {
                if (character == '\\')
                {
                    backslashes++;
                    continue;
                }
                if (character == '\"')
                {
                    result.Append('\\', backslashes * 2 + 1);
                    result.Append('\"');
                    backslashes = 0;
                    continue;
                }
                result.Append('\\', backslashes);
                backslashes = 0;
                result.Append(character);
            }
            result.Append('\\', backslashes * 2);
            result.Append('\"');
            return result.ToString();
        }

        private static bool HasExited(Process process)
        {
            try
            {
                return process.HasExited;
            }
            catch (InvalidOperationException)
            {
                return true;
            }
        }

        private static void IgnoreProcessOutput(object sender, DataReceivedEventArgs eventArgs)
        {
            // Deliberately discard output: it can contain local paths or mail metadata.
        }

        private static void TryCreateCancellationSentinel(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }
            try
            {
                using (File.Create(path))
                {
                }
            }
            catch (IOException)
            {
                // Closing the owned Windows Job still stops the entire process tree.
            }
            catch (UnauthorizedAccessException)
            {
                // Closing the owned Windows Job still stops the entire process tree.
            }
        }

        private static void TryDelete(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }
            try
            {
                File.Delete(path);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private void ThrowIfDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(GetType().FullName);
            }
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }
            disposed = true;
            RequestStop();
            OwnedProcessJob job;
            Process adapter;
            lock (gate)
            {
                job = adapterJob;
                adapter = adapterProcess;
            }
            if (job != null)
            {
                job.Terminate();
            }
            if (adapter != null && !HasExited(adapter))
            {
                try
                {
                    adapter.Kill();
                }
                catch (InvalidOperationException)
                {
                }
            }

            // A service auto-started for the add-in is intentionally left running:
            // it owns the local index/search API and may serve the next Outlook run.
            // We never terminate a pre-existing service process.
        }

        [DataContract]
        private sealed class AdapterProgress
        {
            [DataMember(Name = "phase")]
            public string Phase { get; set; }
            [DataMember(Name = "current")]
            public int Current { get; set; }
            [DataMember(Name = "total")]
            public int Total { get; set; }
            [DataMember(Name = "bodies_truncated")]
            public int BodiesTruncated { get; set; }
        }

        private sealed class WorkspaceLayout
        {
            public string WorkspaceRoot { get; private set; }
            public string ServiceDirectory { get; private set; }
            public string PythonExecutable { get; private set; }
            public string AdapterScript { get; private set; }
            public string ReaderExecutable { get; private set; }
            public string EmbeddingModelDirectory { get; private set; }

            public static WorkspaceLayout Load()
            {
                return Load(false);
            }

            public static WorkspaceLayout LoadForService()
            {
                return Load(true);
            }

            private static WorkspaceLayout Load(bool serviceOnly)
            {
                var manifestDirectory = ReadManifestDirectory();
                var workspaceDirectory = new DirectoryInfo(manifestDirectory);
                for (var level = 0; level < 4; level++)
                {
                    workspaceDirectory = workspaceDirectory.Parent;
                    if (workspaceDirectory == null)
                    {
                        throw new DirectoryNotFoundException(
                            "Manifest RAGSearch не соответствует layout hosts\\outlook_vsto\\bin\\<Configuration>.");
                    }
                }

                var workspace = workspaceDirectory.FullName;
                var solution = Path.Combine(workspace, "RAGSearch.sln");
                var serviceDirectory = Path.Combine(workspace, "service");
                var python = Path.Combine(serviceDirectory, ".venv", "Scripts", "python.exe");
                var serviceEntrypoint = Path.Combine(
                    serviceDirectory,
                    "ragsearch_service",
                    "__main__.py");
                var connectorDirectory = Path.Combine(
                    workspace,
                    "connectors",
                    "outlook_mapi");
                var adapter = Path.Combine(connectorDirectory, "adapter.py");
                var reader = Path.Combine(
                    connectorDirectory,
                    "native",
                    "bin",
                    "x64",
                    ReaderBuildConfiguration,
                    "OutlookMapiReader.exe");
                var missing = serviceOnly
                    ? MissingTool(solution, python, serviceEntrypoint)
                    : MissingTool(solution, python, serviceEntrypoint, adapter, reader);
                if (missing != null)
                {
                    throw new FileNotFoundException(
                        "RAGSearch layout неполон: отсутствует " + missing +
                        ". Соберите solution и создайте service\\.venv.");
                }

                return new WorkspaceLayout
                {
                    WorkspaceRoot = workspace,
                    ServiceDirectory = serviceDirectory,
                    PythonExecutable = python,
                    AdapterScript = adapter,
                    ReaderExecutable = reader,
                    EmbeddingModelDirectory = Path.Combine(
                        serviceDirectory,
                        "models",
                        "paraphrase-multilingual-MiniLM-L12-v2")
                };
            }

            private static string ReadManifestDirectory()
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                           @"Software\Microsoft\Office\Outlook\Addins\RAGSearch",
                           false))
                {
                    if (key == null)
                    {
                        throw new InvalidOperationException(
                            "RAGSearch не зарегистрирован в текущем профиле Outlook.");
                    }
                    var manifest = key.GetValue("Manifest") as string;
                    if (string.IsNullOrWhiteSpace(manifest))
                    {
                        throw new InvalidOperationException(
                            "В регистрации RAGSearch отсутствует путь Manifest.");
                    }
                    var separator = manifest.IndexOf('|');
                    if (separator >= 0)
                    {
                        manifest = manifest.Substring(0, separator);
                    }
                    Uri uri;
                    if (!Uri.TryCreate(manifest, UriKind.Absolute, out uri) || !uri.IsFile)
                    {
                        throw new InvalidOperationException(
                            "Manifest RAGSearch должен быть абсолютным file URI.");
                    }
                    var manifestPath = Path.GetFullPath(uri.LocalPath);
                    var manifestDirectory = Path.GetDirectoryName(manifestPath);
                    if (string.IsNullOrEmpty(manifestDirectory) || !Directory.Exists(manifestDirectory))
                    {
                        throw new DirectoryNotFoundException(
                            "Каталог зарегистрированного Manifest RAGSearch не существует.");
                    }
                    return manifestDirectory;
                }
            }

            private static string MissingTool(params string[] paths)
            {
                foreach (var path in paths)
                {
                    if (!File.Exists(path))
                    {
                        return path;
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// A kill-on-close Windows Job containing only the adapter we start and its
        /// descendants.  Closing it is the enforced process boundary when cooperative
        /// cancellation cannot interrupt a native pipe read.
        /// </summary>
        private sealed class OwnedProcessJob : IDisposable
        {
            private const uint JobObjectLimitKillOnJobClose = 0x00002000;
            private IntPtr handle;

            public OwnedProcessJob()
            {
                handle = CreateJobObject(IntPtr.Zero, null);
                if (handle == IntPtr.Zero)
                {
                    throw new InvalidOperationException(
                        "Не удалось создать Windows Job для Outlook MAPI adapter: " + Marshal.GetLastWin32Error());
                }

                var information = new JobObjectExtendedLimitInformation();
                information.BasicLimitInformation.LimitFlags = JobObjectLimitKillOnJobClose;
                var size = Marshal.SizeOf(typeof(JobObjectExtendedLimitInformation));
                var pointer = Marshal.AllocHGlobal(size);
                try
                {
                    Marshal.StructureToPtr(information, pointer, false);
                    if (!SetInformationJobObject(handle, 9, pointer, (uint)size))
                    {
                        var error = Marshal.GetLastWin32Error();
                        CloseHandle(handle);
                        handle = IntPtr.Zero;
                        throw new InvalidOperationException(
                            "Не удалось настроить Windows Job: " + error);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(pointer);
                }
            }

            public void Assign(Process process)
            {
                if (!AssignProcessToJobObject(handle, process.Handle))
                {
                    throw new InvalidOperationException(
                        "Не удалось привязать Outlook MAPI adapter к Windows Job: " + Marshal.GetLastWin32Error());
                }
            }

            public void Terminate()
            {
                if (handle != IntPtr.Zero)
                {
                    TerminateJobObject(handle, 130);
                }
            }

            public void Dispose()
            {
                if (handle != IntPtr.Zero)
                {
                    CloseHandle(handle);
                    handle = IntPtr.Zero;
                }
            }

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern IntPtr CreateJobObject(IntPtr securityAttributes, string name);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool SetInformationJobObject(
                IntPtr job,
                int informationClass,
                IntPtr information,
                uint informationLength);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool TerminateJobObject(IntPtr job, uint exitCode);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool CloseHandle(IntPtr handle);

            [StructLayout(LayoutKind.Sequential)]
            private struct IoCounters
            {
                public ulong ReadOperationCount;
                public ulong WriteOperationCount;
                public ulong OtherOperationCount;
                public ulong ReadTransferCount;
                public ulong WriteTransferCount;
                public ulong OtherTransferCount;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct JobObjectBasicLimitInformation
            {
                public long PerProcessUserTimeLimit;
                public long PerJobUserTimeLimit;
                public uint LimitFlags;
                public UIntPtr MinimumWorkingSetSize;
                public UIntPtr MaximumWorkingSetSize;
                public uint ActiveProcessLimit;
                public UIntPtr Affinity;
                public uint PriorityClass;
                public uint SchedulingClass;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct JobObjectExtendedLimitInformation
            {
                public JobObjectBasicLimitInformation BasicLimitInformation;
                public IoCounters IoInfo;
                public UIntPtr ProcessMemoryLimit;
                public UIntPtr JobMemoryLimit;
                public UIntPtr PeakProcessMemoryUsed;
                public UIntPtr PeakJobMemoryUsed;
            }
        }
    }

    internal sealed class OutlookMapiImportProgress : EventArgs
    {
        public string Phase { get; set; }
        public int Current { get; set; }
        public int Total { get; set; }
        public string Status { get; set; }
        public bool IsRunning { get; set; }
    }
}
