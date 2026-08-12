using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
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
        private const int MaximumBodyChars = 4000000;
        private const int MaximumJsonLineChars = 32 * 1024 * 1024;
        // The service accepts at most 4096 parts; reserve one for the message body.
        private const int MaximumAttachments = 4095;
        private const int ServiceProtocol = 4;
        private const int MaximumSubjectChars = 65536;
        private const int MaximumLocatorIdChars = 16384;
        private const int MaximumMetadataUtf8Bytes = 1024 * 1024;
        private const int MaximumLocatorUtf8Bytes = 64 * 1024;
        private const int MaximumRecipientChars = 60000;
        private const long MaximumAttachmentBytes = 8L * 1024 * 1024;
        private const long MaximumDocumentAttachmentBytes = 8L * 1024 * 1024;
#if DEBUG
        private const string ReaderBuildConfiguration = "Debug";
#else
        private const string ReaderBuildConfiguration = "Release";
#endif
        private static readonly TimeSpan HealthProbeTimeout = TimeSpan.FromSeconds(2);
        private static readonly TimeSpan ServiceStartupTimeout = TimeSpan.FromSeconds(90);

        private readonly object gate = new object();
        private readonly LocalServiceClient serviceClient;
        private CancellationTokenSource runCancellation;
        private Process readerProcess;
        private Process ownedServiceProcess;
        private OwnedProcessJob readerJob;
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
                await RunReaderAsync(layout, cancellation.Token).ConfigureAwait(false);
            }
            finally
            {
                CleanupReaderState();
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
            var requested = false;
            lock (gate)
            {
                if (runCancellation != null)
                {
                    // Cancellation callbacks run synchronously. Keeping the gate
                    // prevents RunAsync from disposing this CTS between lookup and
                    // Cancel(), while no callback re-enters the runner gate.
                    runCancellation.Cancel();
                    requested = true;
                }
            }

            if (requested)
            {
                Report("stopping", 0, 0, "Останавливаю Outlook MAPI-индексацию...", true);
            }
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
                if (existingOwnedProcess != null)
                {
                    lock (gate)
                    {
                        if (ReferenceEquals(ownedServiceProcess, existingOwnedProcess))
                        {
                            ownedServiceProcess = null;
                        }
                    }
                    existingOwnedProcess.Dispose();
                    existingOwnedProcess = null;
                }

                var arguments = new StringBuilder();
                arguments.Append("-m ragsearch_service");
                arguments.Append(" --port ").Append(serviceClient.ServiceUri.Port);
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
                HealthResponse health;
                try
                {
                    health = await serviceClient.GetHealthAsync(probe.Token).ConfigureAwait(false);
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

                if (health == null || !string.Equals(health.Status, "ok", StringComparison.Ordinal))
                {
                    return false;
                }
                if (health.Protocol != ServiceProtocol)
                {
                    throw new InvalidOperationException(
                        "На 127.0.0.1:8765 запущен несовместимый RAGSearch service " +
                        "(protocol " + health.Protocol + ", ожидается " + ServiceProtocol + "). " +
                        "Завершите старый процесс сервиса и повторите операцию.");
                }
                return true;
            }
        }

        private async Task RunReaderAsync(WorkspaceLayout layout, CancellationToken cancellationToken)
        {
            var runDirectory = CreateRunDirectory();
            try
            {
                await RunReaderInDirectoryAsync(layout, runDirectory, cancellationToken)
                    .ConfigureAwait(false);
            }
            finally
            {
                CleanupRunDirectory(runDirectory);
            }
        }

        private async Task RunReaderInDirectoryAsync(
            WorkspaceLayout layout,
            string runDirectory,
            CancellationToken cancellationToken)
        {
            var arguments = new StringBuilder();
            arguments.Append("--jsonl --max-stores 0 --max-folders 0 --max-messages 0");
            arguments.Append(" --body-preview-chars ").Append(MaximumBodyChars);
            arguments.Append(" --attachment-dir ").Append(QuoteArgument(runDirectory));
            arguments.Append(" --max-attachment-bytes ").Append(MaximumAttachmentBytes);
            arguments.Append(" --max-message-attachment-bytes ")
                .Append(MaximumDocumentAttachmentBytes);
            // Keep the process-wide cap unlimited: the reader and relay enforce an
            // independent budget for every document, so later mail is not starved.
            arguments.Append(" --max-total-attachment-bytes 0");

            var process = new Process
            {
                StartInfo = HiddenProcess(
                    layout.ReaderExecutable,
                    arguments.ToString(),
                    layout.WorkspaceRoot),
                EnableRaisingEvents = true
            };
            var exit = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
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

            OwnedProcessJob job = null;
            var started = false;
            var assigned = false;
            try
            {
                job = new OwnedProcessJob();
                if (!process.Start())
                {
                    throw new InvalidOperationException("Windows did not start OutlookMapiReader.");
                }
                started = true;
                job.Assign(process);
                assigned = true;
                lock (gate)
                {
                    readerProcess = process;
                    readerJob = job;
                }

                // Keep both redirected pipes moving. Diagnostics can contain mailbox
                // data, so stderr is drained asynchronously and never surfaced.
                var stderrDrain = DrainReaderAsync(process.StandardError);
                Report("starting", 0, 0, "Запускаю read-only Extended MAPI...", true);
                var imported = 0;
                var bodiesTruncated = 0;
                var attachmentsTruncated = 0;
                using (cancellationToken.Register(() => TerminateReader(job, process)))
                {
                    while (true)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var line = await process.StandardOutput.ReadLineAsync().ConfigureAwait(false);
                        if (line == null)
                        {
                            break;
                        }
                        if (line.Length == 0)
                        {
                            continue;
                        }
                        if (line.Length > MaximumJsonLineChars)
                        {
                            throw new InvalidDataException("OutlookMapiReader emitted an oversized JSONL record.");
                        }

                        var record = DeserializeReaderRecord(line);
                        var document = MapDocument(record, runDirectory);
                        var outcome = await serviceClient.UpsertDocumentAsync(document, cancellationToken)
                            .ConfigureAwait(false);
                        if (outcome == null ||
                            !string.Equals(outcome.SourceKey, document.SourceKey, StringComparison.Ordinal) ||
                            !string.Equals(outcome.Status, "upserted", StringComparison.Ordinal))
                        {
                            throw new InvalidDataException("RAGSearch service returned an invalid document response.");
                        }

                        TryDeleteRecordAttachments(record, runDirectory);
                        imported++;
                        if (record.BodyTruncated)
                        {
                            bodiesTruncated++;
                        }
                        if (record.AttachmentsTruncated)
                        {
                            attachmentsTruncated++;
                        }
                        Report(
                            "importing",
                            imported,
                            0,
                            "Extended MAPI: проиндексировано " + imported + " писем" +
                            TruncatedSuffix(bodiesTruncated, attachmentsTruncated),
                            true);
                    }
                }

                var exitCode = await exit.Task.ConfigureAwait(false);
                await stderrDrain.ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                if (exitCode != 0)
                {
                    throw new InvalidOperationException(
                        "OutlookMapiReader завершился с кодом " + exitCode +
                        ". Диагностика stderr скрыта, чтобы не выводить данные почты.");
                }
                Report(
                    "complete",
                    imported,
                    imported,
                    "Outlook MAPI-индексация завершена: " + imported + " писем" +
                    TruncatedSuffix(bodiesTruncated, attachmentsTruncated),
                    false);
            }
            catch
            {
                if (started && !HasExited(process))
                {
                    if (assigned)
                    {
                        TerminateReader(job, process);
                    }
                    else
                    {
                        process.Kill();
                    }
                    await WaitForExitAfterTerminationAsync(process, exit.Task).ConfigureAwait(false);
                }
                throw;
            }
            finally
            {
                if (!assigned)
                {
                    process.Dispose();
                    if (job != null)
                    {
                        job.Dispose();
                    }
                }
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
            await Task.WhenAny(exit, Task.Delay(TimeSpan.FromSeconds(3))).ConfigureAwait(false);
        }

        private static void TerminateReader(OwnedProcessJob job, Process process)
        {
            if (job != null && job.TryTerminate())
            {
                return;
            }
            try
            {
                if (process != null && !HasExited(process))
                {
                    process.Kill();
                }
            }
            catch (InvalidOperationException)
            {
                // It exited between HasExited and Kill.
            }
        }

        private static ReaderRecord DeserializeReaderRecord(string json)
        {
            try
            {
                var serializer = new DataContractJsonSerializer(typeof(ReaderRecord));
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    var record = serializer.ReadObject(stream) as ReaderRecord;
                    if (record == null)
                    {
                        throw new InvalidDataException("OutlookMapiReader emitted an empty JSON record.");
                    }
                    return record;
                }
            }
            catch (SerializationException exception)
            {
                throw new InvalidDataException("OutlookMapiReader emitted invalid JSONL.", exception);
            }
        }

        private static async Task DrainReaderAsync(StreamReader reader)
        {
            while (await reader.ReadLineAsync().ConfigureAwait(false) != null)
            {
                // Deliberately discard diagnostics: they may contain mailbox names,
                // folder paths, and MAPI identifiers.
            }
        }

        private static DocumentDto MapDocument(ReaderRecord record, string runDirectory)
        {
            var storeId = ReaderText(record.StoreId, "store_id", true, MaximumLocatorIdChars);
            var entryId = ReaderText(record.EntryId, "entry_id", true, MaximumLocatorIdChars);
            var folderEntryId = ReaderText(
                record.FolderEntryId,
                "folder_entry_id",
                true,
                MaximumLocatorIdChars);
            var body = ReaderText(record.Body, "body", false, MaximumBodyChars);
            if (!record.BodyAvailable && body.Length != 0)
            {
                throw new InvalidDataException("Reader body must be empty when body_available is false.");
            }
            if (record.BodyTruncated && !record.BodyAvailable)
            {
                throw new InvalidDataException("Reader body cannot be truncated when unavailable.");
            }
            if (record.Attachments == null || record.Attachments.Count > MaximumAttachments)
            {
                throw new InvalidDataException("Reader attachments must be a bounded array.");
            }

            var sourceKey = BuildSourceKey(storeId, entryId);
            var parts = new List<DocumentPartDto>();
            if (record.BodyAvailable)
            {
                parts.Add(new DocumentPartDto
                {
                    Key = "body",
                    Kind = "body",
                    Name = "body",
                    MediaType = "text/plain",
                    Size = Encoding.UTF8.GetByteCount(body),
                    Text = body,
                    Truncated = record.BodyTruncated
                });
            }

            long inlineAttachmentBytes = 0;
            for (var index = 0; index < record.Attachments.Count; index++)
            {
                var attachment = record.Attachments[index];
                if (attachment == null || attachment.Size < 0)
                {
                    throw new InvalidDataException("Reader attachment has an invalid size.");
                }
                var part = new DocumentPartDto
                {
                    Key = "attachment:" + index,
                    Kind = "attachment",
                    Name = ReaderText(attachment.Name, "attachment.name", false, 32768),
                    MediaType = ReaderText(attachment.ContentType, "attachment.content_type", false, 4096),
                    Size = attachment.Size
                };
                var tempPath = ReaderText(attachment.TempPath, "attachment.temp_path", false, 32768);
                if (tempPath.Length != 0)
                {
                    var safePath = ValidateAttachmentPath(tempPath, runDirectory);
                    var length = new FileInfo(safePath).Length;
                    if (length != attachment.Size || length > MaximumAttachmentBytes)
                    {
                        throw new InvalidDataException("Reader attachment size does not match its temporary file.");
                    }
                    if (inlineAttachmentBytes + length <= MaximumDocumentAttachmentBytes)
                    {
                        part.ContentBase64 = Convert.ToBase64String(File.ReadAllBytes(safePath));
                        inlineAttachmentBytes += length;
                    }
                }
                parts.Add(part);
            }

            var document = new DocumentDto
            {
                SourceKey = sourceKey,
                Kind = "email",
                Title = ReaderText(record.Subject, "subject", false, MaximumSubjectChars),
                Metadata = new OutlookDocumentMetadata
                {
                    SenderName = ReaderText(record.SenderName, "sender_name", false, 32768),
                    SenderEmail = ReaderText(record.SenderEmail, "sender_email", false, 32768),
                    To = ReaderText(record.To, "to", false, MaximumRecipientChars),
                    Cc = ReaderText(record.Cc, "cc", false, MaximumRecipientChars),
                    SentAt = ReaderTimestamp(record.SentAt, "sent_at"),
                    ReceivedAt = ReaderTimestamp(record.ReceivedAt, "received_at"),
                    ModifiedAt = ReaderTimestamp(record.ModifiedAt, "modified_at"),
                    FolderPath = ReaderText(record.FolderPath, "folder_path", true, 16384),
                    StoreName = ReaderText(record.StoreName, "store_name", false, 4096),
                    InternetMessageId = ReaderText(record.InternetMessageId, "internet_message_id", false, 32768),
                    ConversationId = ReaderText(record.ConversationId, "conversation_id", false, 32768),
                    AttachmentsTruncated = record.AttachmentsTruncated
                },
                Locator = new OutlookDocumentLocator
                {
                    Connector = "outlook_mapi",
                    StoreId = storeId,
                    EntryId = entryId,
                    FolderEntryId = folderEntryId
                },
                Parts = parts
            };
            ValidateNeutralEnvelope(document);
            return document;
        }

        private static void ValidateNeutralEnvelope(DocumentDto document)
        {
            if (SerializedUtf8Length(document.Metadata) > MaximumMetadataUtf8Bytes)
            {
                throw new InvalidDataException(
                    "Reader metadata exceeds the neutral service contract.");
            }
            if (SerializedUtf8Length(document.Locator) > MaximumLocatorUtf8Bytes)
            {
                throw new InvalidDataException(
                    "Reader locator exceeds the neutral service contract.");
            }
        }

        private static long SerializedUtf8Length(object value)
        {
            var serializer = new DataContractJsonSerializer(value.GetType());
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, value);
                return stream.Length;
            }
        }

        private static string ReaderText(string value, string name, bool required, int maximumLength)
        {
            if (value == null || (required && string.IsNullOrWhiteSpace(value)))
            {
                throw new InvalidDataException("Reader field " + name + " is missing.");
            }
            value = value ?? string.Empty;
            if (value.Length > maximumLength || value.IndexOf('\0') >= 0)
            {
                throw new InvalidDataException("Reader field " + name + " is invalid.");
            }
            return value;
        }

        private static string ReaderTimestamp(string value, string name)
        {
            if (value == null)
            {
                return null;
            }
            DateTimeOffset parsed;
            if (value.Length > 64 || value.IndexOf('T') < 0 ||
                !DateTimeOffset.TryParse(
                    value,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind,
                    out parsed))
            {
                throw new InvalidDataException("Reader field " + name + " is not an ISO-8601 timestamp.");
            }
            return value;
        }

        private static string BuildSourceKey(string storeId, string entryId)
        {
            using (var sha256 = SHA256.Create())
            {
                var identity = storeId.Length.ToString(CultureInfo.InvariantCulture) + ":" +
                               storeId + entryId.Length.ToString(CultureInfo.InvariantCulture) + ":" +
                               entryId;
                var digest = sha256.ComputeHash(Encoding.UTF8.GetBytes(identity));
                var value = new StringBuilder("outlook_mapi:");
                foreach (var item in digest)
                {
                    value.Append(item.ToString("x2", CultureInfo.InvariantCulture));
                }
                return value.ToString();
            }
        }

        private static string ValidateAttachmentPath(string rawPath, string runDirectory)
        {
            var root = Path.GetFullPath(runDirectory).TrimEnd(Path.DirectorySeparatorChar);
            var candidate = Path.GetFullPath(rawPath);
            var parent = Path.GetDirectoryName(candidate);
            if (!Directory.Exists(root) ||
                (File.GetAttributes(root) & FileAttributes.ReparsePoint) != 0 ||
                !string.Equals(parent, root, StringComparison.OrdinalIgnoreCase) ||
                !File.Exists(candidate) ||
                (File.GetAttributes(candidate) & FileAttributes.ReparsePoint) != 0)
            {
                throw new InvalidDataException("Reader attachment path escaped its private run directory.");
            }
            return candidate;
        }

        private static string TruncatedSuffix(int bodiesTruncated, int attachmentsTruncated)
        {
            var value = new StringBuilder();
            if (bodiesTruncated > 0)
            {
                value.Append("; body усечено по лимиту: ").Append(bodiesTruncated);
            }
            if (attachmentsTruncated > 0)
            {
                value.Append("; список вложений усечён: ").Append(attachmentsTruncated);
            }
            return value.ToString();
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

        private void CleanupReaderState()
        {
            Process process;
            OwnedProcessJob job;
            lock (gate)
            {
                process = readerProcess;
                job = readerJob;
                readerProcess = null;
                readerJob = null;
            }

            if (process != null)
            {
                process.Dispose();
            }
            if (job != null)
            {
                job.Dispose();
            }
        }

        private static ProcessStartInfo HiddenPython(string python, string arguments, string workingDirectory)
        {
            return HiddenProcess(python, arguments, workingDirectory);
        }

        private static ProcessStartInfo HiddenProcess(string executable, string arguments, string workingDirectory)
        {
            return new ProcessStartInfo
            {
                FileName = executable,
                Arguments = arguments,
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = new UTF8Encoding(false, true),
                StandardErrorEncoding = new UTF8Encoding(false, true)
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

        private static string CreateRunDirectory()
        {
            var root = GetReaderRunsRoot();
            Directory.CreateDirectory(root);
            if ((File.GetAttributes(root) & FileAttributes.ReparsePoint) != 0)
            {
                throw new InvalidDataException("RAGSearch reader-runs must not be a reparse point.");
            }
            var run = Path.Combine(root, "outlook-mapi-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(run);
            run = Path.GetFullPath(run);
            if ((File.GetAttributes(run) & FileAttributes.ReparsePoint) != 0)
            {
                throw new InvalidDataException("RAGSearch reader run directory must not be a reparse point.");
            }
            return run;
        }

        private static string GetReaderRunsRoot()
        {
            var localAppData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(localAppData))
            {
                throw new InvalidOperationException("LOCALAPPDATA is unavailable.");
            }
            return Path.GetFullPath(Path.Combine(localAppData, "RAGSearch", "reader-runs"));
        }

        private static void TryDeleteRecordAttachments(ReaderRecord record, string runDirectory)
        {
            foreach (var attachment in record.Attachments)
            {
                if (attachment == null || string.IsNullOrEmpty(attachment.TempPath))
                {
                    continue;
                }
                try
                {
                    var safePath = ValidateAttachmentPath(attachment.TempPath, runDirectory);
                    File.Delete(safePath);
                }
                catch (InvalidDataException)
                {
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }
        }

        private static void CleanupRunDirectory(string runDirectory)
        {
            try
            {
                var expectedRoot = GetReaderRunsRoot();
                var run = new DirectoryInfo(Path.GetFullPath(runDirectory));
                if (run.Parent == null ||
                    !string.Equals(
                        Path.GetFullPath(run.Parent.FullName),
                        expectedRoot,
                        StringComparison.OrdinalIgnoreCase) ||
                    !run.Name.StartsWith("outlook-mapi-", StringComparison.Ordinal) ||
                    !run.Exists)
                {
                    return;
                }
                var rootAttributes = File.GetAttributes(expectedRoot);
                var runAttributes = File.GetAttributes(run.FullName);
                if ((rootAttributes & FileAttributes.ReparsePoint) != 0 ||
                    (runAttributes & FileAttributes.ReparsePoint) != 0 ||
                    run.GetDirectories().Length != 0)
                {
                    return;
                }
                foreach (var file in run.GetFiles())
                {
                    if ((file.Attributes & FileAttributes.ReparsePoint) != 0)
                    {
                        return;
                    }
                }
                foreach (var file in run.GetFiles())
                {
                    file.Delete();
                }
                run.Delete(false);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            catch (InvalidOperationException)
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
            Process reader;
            Process service;
            lock (gate)
            {
                job = readerJob;
                reader = readerProcess;
                service = ownedServiceProcess;
                ownedServiceProcess = null;
            }
            if (job != null)
            {
                TerminateReader(job, reader);
            }
            if (reader != null && !HasExited(reader))
            {
                try
                {
                    reader.Kill();
                }
                catch (InvalidOperationException)
                {
                }
            }

            if (service != null)
            {
                service.Dispose();
            }

            // A service auto-started for the add-in is intentionally left running:
            // it owns the local index/search API and may serve the next Outlook run.
            // We never terminate a pre-existing service process.
        }

        [DataContract]
        private sealed class ReaderRecord
        {
            [DataMember(Name = "store_id", IsRequired = true)]
            public string StoreId { get; set; }
            [DataMember(Name = "store_name", IsRequired = true)]
            public string StoreName { get; set; }
            [DataMember(Name = "entry_id", IsRequired = true)]
            public string EntryId { get; set; }
            [DataMember(Name = "folder_entry_id", IsRequired = true)]
            public string FolderEntryId { get; set; }
            [DataMember(Name = "folder_path", IsRequired = true)]
            public string FolderPath { get; set; }
            [DataMember(Name = "subject", IsRequired = true)]
            public string Subject { get; set; }
            [DataMember(Name = "body", IsRequired = true)]
            public string Body { get; set; }
            [DataMember(Name = "body_available", IsRequired = true)]
            public bool BodyAvailable { get; set; }
            [DataMember(Name = "body_truncated", IsRequired = true)]
            public bool BodyTruncated { get; set; }
            [DataMember(Name = "sender_name", IsRequired = true)]
            public string SenderName { get; set; }
            [DataMember(Name = "sender_email", IsRequired = true)]
            public string SenderEmail { get; set; }
            [DataMember(Name = "to", IsRequired = true)]
            public string To { get; set; }
            [DataMember(Name = "cc", IsRequired = true)]
            public string Cc { get; set; }
            [DataMember(Name = "sent_at", IsRequired = true)]
            public string SentAt { get; set; }
            [DataMember(Name = "received_at", IsRequired = true)]
            public string ReceivedAt { get; set; }
            [DataMember(Name = "modified_at", IsRequired = true)]
            public string ModifiedAt { get; set; }
            [DataMember(Name = "internet_message_id", IsRequired = true)]
            public string InternetMessageId { get; set; }
            [DataMember(Name = "conversation_id", IsRequired = true)]
            public string ConversationId { get; set; }
            [DataMember(Name = "attachments", IsRequired = true)]
            public List<ReaderAttachment> Attachments { get; set; }
            [DataMember(Name = "attachments_truncated", IsRequired = true)]
            public bool AttachmentsTruncated { get; set; }
        }

        [DataContract]
        private sealed class ReaderAttachment
        {
            [DataMember(Name = "name", IsRequired = true)]
            public string Name { get; set; }
            [DataMember(Name = "size", IsRequired = true)]
            public long Size { get; set; }
            [DataMember(Name = "content_type", IsRequired = true)]
            public string ContentType { get; set; }
            [DataMember(Name = "temp_path", IsRequired = true)]
            public string TempPath { get; set; }
        }

        private sealed class WorkspaceLayout
        {
            public string WorkspaceRoot { get; private set; }
            public string ServiceDirectory { get; private set; }
            public string PythonExecutable { get; private set; }
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
                var reader = Path.Combine(
                    connectorDirectory,
                    "native",
                    "bin",
                    "x64",
                    ReaderBuildConfiguration,
                    "OutlookMapiReader.exe");
                var missing = serviceOnly
                    ? MissingTool(solution, python, serviceEntrypoint)
                    : MissingTool(solution, python, serviceEntrypoint, reader);
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
        /// A kill-on-close Windows Job containing only the native reader we start.
        /// Closing it is the enforced process boundary when cancellation interrupts a pipe read.
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
                        "Не удалось создать Windows Job для OutlookMapiReader: " + Marshal.GetLastWin32Error());
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
                        "Не удалось привязать OutlookMapiReader к Windows Job: " + Marshal.GetLastWin32Error());
                }
            }

            public bool TryTerminate()
            {
                return handle != IntPtr.Zero && TerminateJobObject(handle, 130);
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
