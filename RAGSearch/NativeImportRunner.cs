using System;
using System.Collections.Generic;
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
    internal sealed class NativeImportRunner : IDisposable
    {
        private const string ProgressPrefix = "RAGSEARCH_PROGRESS ";
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

        public NativeImportRunner(LocalServiceClient serviceClient)
        {
            this.serviceClient = serviceClient ?? throw new ArgumentNullException("serviceClient");
        }

        public event EventHandler<NativeImportProgress> ProgressChanged;

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
                    throw new InvalidOperationException("Native MAPI indexing is already running.");
                }

                isRunning = true;
                cancellation = new CancellationTokenSource();
                runCancellation = cancellation;
            }

            try
            {
                var tools = WorkspaceTools.Find();
                Report("checking_service", 0, 0, "Проверяю локальный сервис...", true);
                await EnsureServiceAsync(tools, cancellation.Token).ConfigureAwait(false);
                cancellation.Token.ThrowIfCancellationRequested();
                await RunAdapterAsync(tools, cancellation.Token).ConfigureAwait(false);
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
                        "Нельзя отдельно запускать сервис во время native MAPI индексации.");
                }
            }

            if (await IsServiceHealthyAsync(cancellationToken).ConfigureAwait(false))
            {
                return;
            }

            var tools = WorkspaceTools.FindForService();
            await EnsureServiceAsync(tools, cancellationToken).ConfigureAwait(false);
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

            Report("stopping", 0, 0, "Останавливаю native-индексацию...", true);
            TryCreateCancellationSentinel(sentinel);
            cancellation.Cancel();
        }

        private async Task EnsureServiceAsync(WorkspaceTools tools, CancellationToken cancellationToken)
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
                arguments.Append(QuoteArgument(tools.ServiceScript));
                arguments.Append(" --port ").Append(serviceClient.ServiceUri.Port);
                arguments.Append(" --delete-spool-after-ingest");
                if (Directory.Exists(tools.EmbeddingModelDirectory))
                {
                    arguments.Append(" --embedding sentence-transformers --model ");
                    arguments.Append(QuoteArgument(tools.EmbeddingModelDirectory));
                }

                var startInfo = HiddenPython(tools.PythonExecutable, arguments.ToString(), tools.ServiceDirectory);
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
                        "Python-сервис завершился при запуске. Проверьте service\\run.py и локальную модель embeddings.");
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
                    return health != null && string.Equals(health.status, "ok", StringComparison.OrdinalIgnoreCase);
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

        private async Task RunAdapterAsync(WorkspaceTools tools, CancellationToken cancellationToken)
        {
            var sentinel = Path.Combine(
                Path.GetTempPath(),
                "ragsearch-native-cancel-" + Guid.NewGuid().ToString("N") + ".flag");
            lock (gate)
            {
                cancelFilePath = sentinel;
            }

            var arguments = new StringBuilder();
            arguments.Append(QuoteArgument(tools.AdapterScript));
            arguments.Append(" --executable ").Append(QuoteArgument(tools.NativeExecutable));
            arguments.Append(" --service-url ").Append(QuoteArgument(serviceClient.ServiceUri.GetLeftPart(UriPartial.Authority)));
            // Production indexing must not inherit the adapter's bounded probe defaults.
            arguments.Append(" --full-scan --body-preview-chars 4000000");
            arguments.Append(" --cancel-file ").Append(QuoteArgument(sentinel));

            var process = new Process
            {
                StartInfo = HiddenPython(tools.PythonExecutable, arguments.ToString(), tools.WorkspaceRoot),
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
                    throw new InvalidOperationException("Windows did not start the native adapter.");
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
                            // adapter can normally create its probe child.  If Windows
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
                        // its NativeMapiProbe child.  No pre-existing process is touched.
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
                    "Native MAPI adapter завершился с кодом " + exitCode +
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

            if (value == null || value.current < 0 || value.total < 0 || value.bodies_truncated < 0)
            {
                return;
            }

            string status;
            switch ((value.phase ?? string.Empty).ToLowerInvariant())
            {
                case "importing":
                    status = "Extended MAPI: проиндексировано " + value.current + " писем" +
                             TruncatedSuffix(value.bodies_truncated);
                    break;
                case "complete":
                    status = "Native-индексация завершена: " + value.current + " писем" +
                             TruncatedSuffix(value.bodies_truncated);
                    break;
                case "cancelled":
                    status = "Native-индексация остановлена: " + value.current + " писем" +
                             TruncatedSuffix(value.bodies_truncated);
                    break;
                default:
                    status = "Extended MAPI: подготовка...";
                    break;
            }
            Report(value.phase, value.current, value.total, status, value.phase != "complete" && value.phase != "cancelled");
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
                handler(this, new NativeImportProgress
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
                // Fallback job termination will still stop the owned process tree.
            }
            catch (UnauthorizedAccessException)
            {
                // Fallback job termination will still stop the owned process tree.
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
            public string phase { get; set; }
            [DataMember(Name = "current")]
            public int current { get; set; }
            [DataMember(Name = "total")]
            public int total { get; set; }
            [DataMember(Name = "bodies_truncated")]
            public int bodies_truncated { get; set; }
        }

        private sealed class WorkspaceTools
        {
            public string WorkspaceRoot { get; private set; }
            public string ServiceDirectory { get; private set; }
            public string PythonExecutable { get; private set; }
            public string ServiceScript { get; private set; }
            public string AdapterScript { get; private set; }
            public string NativeExecutable { get; private set; }
            public string EmbeddingModelDirectory { get; private set; }

            public static WorkspaceTools Find()
            {
                return Find(false);
            }

            public static WorkspaceTools FindForService()
            {
                return Find(true);
            }

            private static WorkspaceTools Find(bool serviceOnly)
            {
                var assemblyPath = typeof(NativeImportRunner).Assembly.Location;
                var startingDirectories = new List<string>();
                AddCandidate(startingDirectories, Environment.GetEnvironmentVariable("RAGSEARCH_WORKSPACE"));
                AddCandidate(startingDirectories, Path.GetDirectoryName(assemblyPath));
                AddCandidate(startingDirectories, ReadManifestDirectory());

                foreach (var startingDirectory in startingDirectories)
                {
                    var directory = new DirectoryInfo(startingDirectory);
                    for (var depth = 0; directory != null && depth < 10; depth++, directory = directory.Parent)
                    {
                        var solution = Path.Combine(directory.FullName, "RAGSearch.sln");
                        if (!File.Exists(solution))
                        {
                            continue;
                        }

                        var serviceDirectory = Path.Combine(directory.FullName, "service");
                        var python = Path.Combine(serviceDirectory, ".venv", "Scripts", "python.exe");
                        var service = Path.Combine(serviceDirectory, "run.py");
                        var adapter = Path.Combine(serviceDirectory, "import_native_mapi.py");
                        var native = Path.Combine(
                            directory.FullName,
                            "native-mapi-probe",
                            "build-direct",
                            "NativeMapiProbe.exe");
                        var missing = serviceOnly
                            ? MissingTool(python, service)
                            : MissingTool(python, service, adapter, native);
                        if (missing != null)
                        {
                            throw new FileNotFoundException(
                                "RAGSearch workspace найден, но отсутствует " + missing +
                                ". Соберите native worker и создайте service\\.venv.");
                        }

                        return new WorkspaceTools
                        {
                            WorkspaceRoot = directory.FullName,
                            ServiceDirectory = serviceDirectory,
                            PythonExecutable = python,
                            ServiceScript = service,
                            AdapterScript = adapter,
                            NativeExecutable = native,
                            EmbeddingModelDirectory = Path.Combine(
                                serviceDirectory,
                                "models",
                                "paraphrase-multilingual-MiniLM-L12-v2")
                        };
                    }
                }

                throw new DirectoryNotFoundException(
                    "Не найден workspace RAGSearch рядом с установленной надстройкой. " +
                    "Ожидались RAGSearch.sln, service и native-mapi-probe выше " + assemblyPath + ".");
            }

            private static void AddCandidate(ICollection<string> candidates, string path)
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    return;
                }
                try
                {
                    var fullPath = Path.GetFullPath(path.Trim());
                    if (Directory.Exists(fullPath) && !candidates.Contains(fullPath))
                    {
                        candidates.Add(fullPath);
                    }
                }
                catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException)
                {
                    // Ignore an invalid optional override/manifest and use the next source.
                }
            }

            private static string ReadManifestDirectory()
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(
                               @"Software\Microsoft\Office\Outlook\Addins\RAGSearch",
                               false))
                    {
                        var manifest = key == null ? null : key.GetValue("Manifest") as string;
                        if (string.IsNullOrWhiteSpace(manifest))
                        {
                            return null;
                        }
                        var separator = manifest.IndexOf('|');
                        if (separator >= 0)
                        {
                            manifest = manifest.Substring(0, separator);
                        }
                        Uri uri;
                        var manifestPath = Uri.TryCreate(manifest, UriKind.Absolute, out uri) && uri.IsFile
                            ? uri.LocalPath
                            : manifest;
                        return Path.GetDirectoryName(manifestPath);
                    }
                }
                catch (Exception ex) when (
                    ex is IOException ||
                    ex is UnauthorizedAccessException ||
                    ex is System.Security.SecurityException ||
                    ex is ArgumentException)
                {
                    return null;
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
        /// descendants.  It is the bounded fallback when cooperative cancellation
        /// cannot interrupt a native pipe read.
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
                        "Не удалось создать Windows Job для native adapter: " + Marshal.GetLastWin32Error());
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
                        "Не удалось привязать native adapter к Windows Job: " + Marshal.GetLastWin32Error());
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

    internal sealed class NativeImportProgress : EventArgs
    {
        public string Phase { get; set; }
        public int Current { get; set; }
        public int Total { get; set; }
        public string Status { get; set; }
        public bool IsRunning { get; set; }
    }
}
