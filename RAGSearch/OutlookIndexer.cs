using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RAGSearch
{
    internal sealed class OutlookIndexer : IDisposable
    {
        // One DTO can contain up to four million Unicode characters. Keeping one
        // message per request makes the 32 MiB service limit deterministic.
        private const int BatchSize = 1;
        private readonly Outlook.Application application;
        private Outlook.NameSpace session;
        private readonly LocalServiceClient serviceClient;
        private readonly OutlookItemExtractor extractor;
        private readonly System.Windows.Forms.Timer timer;
        private readonly Queue<FolderCursor> folders = new Queue<FolderCursor>();
        private readonly string logPath;

        private CancellationTokenSource cancellation;
        private Control uiControl;
        private bool tickInProgress;
        private int processed;
        private int failed;
        private int estimatedTotal;

        public OutlookIndexer(
            Outlook.Application application,
            LocalServiceClient serviceClient,
            OutlookItemExtractor extractor)
        {
            this.application = application ?? throw new ArgumentNullException("application");
            this.serviceClient = serviceClient ?? throw new ArgumentNullException("serviceClient");
            this.extractor = extractor ?? throw new ArgumentNullException("extractor");
            logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RAGSearch",
                "addin.log");
            timer = new System.Windows.Forms.Timer { Interval = 40 };
            timer.Tick += TimerOnTick;
        }

        public event EventHandler<IndexProgress> ProgressChanged;

        public bool IsRunning { get; private set; }

        public void SetUiControl(Control control)
        {
            uiControl = control;
        }

        public void Start()
        {
            Stop();
            folders.Clear();
            processed = 0;
            failed = 0;
            estimatedTotal = 0;
            cancellation = new CancellationTokenSource();
            ResetLog();
            EnsureSession();
            var trustState = ReadTrustState();
            Log("Outlook.Application.IsTrusted=" + trustState);

            try
            {
                AttachPstArchives();
                DiscoverFolders();
                IsRunning = folders.Count > 0;
                PublishProgress(
                    IsRunning
                        ? "Индексация запущена; OOM IsTrusted=" + trustState
                        : "Почтовые папки не найдены; OOM IsTrusted=" + trustState,
                    null);
                if (IsRunning)
                {
                    timer.Start();
                }
            }
            catch (COMException ex)
            {
                IsRunning = false;
                PublishProgress("Ошибка Outlook при обходе папок: " + ex.Message, null);
            }
            catch (Exception ex)
            {
                IsRunning = false;
                PublishProgress("Не удалось подготовить индексацию: " + ex.Message, null);
            }
        }

        public void Stop()
        {
            timer.Stop();
            if (cancellation != null)
            {
                cancellation.Cancel();
                cancellation.Dispose();
                cancellation = null;
            }

            if (IsRunning)
            {
                IsRunning = false;
                PublishProgress("Индексация остановлена", null);
            }
        }

        public async void IndexNewMail(string entryIdCollection)
        {
            if (string.IsNullOrWhiteSpace(entryIdCollection))
            {
                return;
            }

            EnsureSession();

            foreach (var entryId in entryIdCollection.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                MessagePayload payload = null;
                string folderPath = null;
                object item = null;
                Outlook.MAPIFolder folder = null;
                Outlook.Store store = null;
                try
                {
                    item = session.GetItemFromID(entryId.Trim(), Type.Missing);
                    var mail = item as Outlook.MailItem;
                    if (mail == null)
                    {
                        continue;
                    }

                    folder = mail.Parent as Outlook.MAPIFolder;
                    if (folder == null)
                    {
                        continue;
                    }

                    store = folder.Store;
                    folderPath = folder.FolderPath;
                    payload = extractor.Extract(mail, folder, store == null ? null : store.DisplayName);
                }
                catch (COMException ex)
                {
                    RunOnUi(() => PublishProgress(
                        "Не удалось прочитать новое письмо: " + ex.Message,
                        null));
                }
                catch (Exception ex)
                {
                    RunOnUi(() => PublishProgress(
                        "Не удалось подготовить новое письмо: " + ex.Message,
                        null));
                }
                finally
                {
                    ComRelease.Final(store);
                    ComRelease.Final(folder);
                    ComRelease.Final(item);
                }

                if (payload == null)
                {
                    continue;
                }

                try
                {
                    await serviceClient.IngestAsync(
                            new[] { payload },
                            CancellationToken.None)
                        .ConfigureAwait(false);
                    var completedFolder = folderPath;
                    RunOnUi(() => PublishProgress(
                        "Новое письмо добавлено в индекс",
                        completedFolder));
                }
                catch (Exception ex)
                {
                    RunOnUi(() => PublishProgress(
                        "Не удалось индексировать новое письмо: " + ex.Message,
                        null));
                }
            }
        }

        private void DiscoverFolders()
        {
            Outlook.Stores stores = null;
            try
            {
                stores = session.Stores;
                for (var index = 1; index <= stores.Count; index++)
                {
                    Outlook.Store store = null;
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        store = stores[index];
                        root = store.GetRootFolder();
                        DiscoverFolderRecursive(
                            root,
                            store.StoreID,
                            store.DisplayName);
                    }
                    catch (COMException ex)
                    {
                        failed++;
                        Log("Store discovery failed: " + ex.Message);
                    }
                    finally
                    {
                        ComRelease.Final(root);
                        ComRelease.Final(store);
                    }
                }
            }
            finally
            {
                ComRelease.Final(stores);
            }
        }

        private void AttachPstArchives()
        {
            var outlookFiles = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Outlook Files");
            if (!Directory.Exists(outlookFiles))
            {
                return;
            }

            var attached = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Outlook.Stores stores = null;
            try
            {
                stores = session.Stores;
                for (var index = 1; index <= stores.Count; index++)
                {
                    Outlook.Store store = null;
                    try
                    {
                        store = stores[index];
                        var filePath = store.FilePath;
                        if (!string.IsNullOrWhiteSpace(filePath) &&
                            filePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
                        {
                            attached.Add(Path.GetFullPath(filePath));
                        }
                    }
                    catch (COMException ex)
                    {
                        Log("Could not inspect a connected store: " + ex.Message);
                    }
                    finally
                    {
                        ComRelease.Final(store);
                    }
                }
            }
            finally
            {
                ComRelease.Final(stores);
            }

            foreach (var filePath in Directory.EnumerateFiles(
                outlookFiles,
                "*.pst",
                SearchOption.TopDirectoryOnly))
            {
                var fullPath = Path.GetFullPath(filePath);
                if (attached.Contains(fullPath))
                {
                    continue;
                }

                try
                {
                    session.AddStore(fullPath);
                    attached.Add(fullPath);
                    Log("Attached PST store: " + fullPath);
                }
                catch (COMException ex)
                {
                    failed++;
                    Log("Could not attach PST store " + fullPath + ": " + ex.Message);
                }
            }
        }

        private void DiscoverFolderRecursive(
            Outlook.MAPIFolder folder,
            string storeId,
            string storeName)
        {
            if (folder == null)
            {
                return;
            }

            Outlook.Items items = null;
            Outlook.Folders children = null;
            try
            {
                if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
                {
                    items = folder.Items;
                    var itemCount = items.Count;
                    if (itemCount > 0)
                    {
                        folders.Enqueue(new FolderCursor
                        {
                            StoreId = storeId,
                            StoreName = storeName,
                            FolderEntryId = folder.EntryID,
                            FolderPath = folder.FolderPath,
                            NextIndex = 1
                        });
                        estimatedTotal += itemCount;
                    }
                }

                children = folder.Folders;
                for (var index = 1; index <= children.Count; index++)
                {
                    Outlook.MAPIFolder child = null;
                    try
                    {
                        child = children[index];
                        DiscoverFolderRecursive(child, storeId, storeName);
                    }
                    catch (COMException ex)
                    {
                        failed++;
                        Log("Folder discovery failed: " + ex.Message);
                    }
                    finally
                    {
                        ComRelease.Final(child);
                    }
                }
            }
            finally
            {
                ComRelease.Final(children);
                ComRelease.Final(items);
            }
        }

        private async void TimerOnTick(object sender, EventArgs eventArgs)
        {
            if (tickInProgress || !IsRunning)
            {
                return;
            }

            tickInProgress = true;
            timer.Stop();
            List<MessagePayload> batch;
            CancellationToken token;
            try
            {
                token = cancellation == null ? CancellationToken.None : cancellation.Token;
                token.ThrowIfCancellationRequested();
                batch = ReadNextBatch(token);
            }
            catch (OperationCanceledException)
            {
                IsRunning = false;
                PublishProgress("Индексация остановлена", null);
                tickInProgress = false;
                return;
            }
            catch (COMException ex)
            {
                IsRunning = false;
                failed++;
                PublishProgress("Индексация приостановлена: " + ex.Message, null);
                tickInProgress = false;
                return;
            }
            catch (Exception ex)
            {
                IsRunning = false;
                failed++;
                PublishProgress("Индексация приостановлена: " + ex.Message, null);
                tickInProgress = false;
                return;
            }

            try
            {
                var response = batch.Count == 0
                    ? null
                    : await serviceClient.IngestAsync(batch, token).ConfigureAwait(false);
                RunOnUi(() => CompleteTick(response));
            }
            catch (OperationCanceledException)
            {
                RunOnUi(() =>
                {
                    IsRunning = false;
                    tickInProgress = false;
                    PublishProgress("Индексация остановлена", null);
                });
            }
            catch (Exception ex)
            {
                RunOnUi(() =>
                {
                    IsRunning = false;
                    tickInProgress = false;
                    failed++;
                    PublishProgress("Индексация приостановлена: " + ex.Message, null);
                });
            }
        }

        private void CompleteTick(IngestResponse response)
        {
            if (response != null)
            {
                failed += response.failed;
                if (response.errors != null)
                {
                    foreach (var error in response.errors)
                    {
                        Log(string.Format(
                            "Service rejected batch item {0}; entry={1}; store={2}; error={3}",
                            error.index,
                            error.entry_id,
                            error.store_id,
                            error.error));
                    }
                }
            }

            tickInProgress = false;
            if (!IsRunning)
            {
                return;
            }

            if (folders.Count == 0)
            {
                IsRunning = false;
                PublishProgress("Индексация завершена", null);
                return;
            }

            PublishProgress("Индексируется", folders.Peek().FolderPath);
            timer.Start();
        }

        private List<MessagePayload> ReadNextBatch(CancellationToken token)
        {
            var batch = new List<MessagePayload>(BatchSize);
            while (batch.Count < BatchSize && folders.Count > 0)
            {
                token.ThrowIfCancellationRequested();
                var cursor = folders.Peek();
                Outlook.MAPIFolder folder = null;
                Outlook.Items items = null;
                try
                {
                    folder = session.GetFolderFromID(cursor.FolderEntryId, cursor.StoreId);
                    items = folder.Items;
                    var count = items.Count;
                    while (batch.Count < BatchSize && cursor.NextIndex <= count)
                    {
                        object item = null;
                        try
                        {
                            item = items[cursor.NextIndex];
                            cursor.NextIndex++;
                            processed++;
                            var mail = item as Outlook.MailItem;
                            if (mail == null)
                            {
                                continue;
                            }

                            var payload = extractor.Extract(mail, folder, cursor.StoreName);
                            if (payload != null)
                            {
                                batch.Add(payload);
                            }
                        }
                        catch (COMException ex)
                        {
                            failed++;
                            Log(string.Format(
                                "Item extraction failed in {0} at index {1}: 0x{2:X8} {3}",
                                cursor.FolderPath,
                                cursor.NextIndex - 1,
                                ex.ErrorCode,
                                ex.Message));
                        }
                        finally
                        {
                            ComRelease.Final(item);
                        }
                    }

                    if (cursor.NextIndex > count)
                    {
                        folders.Dequeue();
                    }
                }
                catch (COMException ex)
                {
                    failed++;
                    Log("Folder read failed for " + cursor.FolderPath + ": " + ex.Message);
                    folders.Dequeue();
                }
                finally
                {
                    ComRelease.Final(items);
                    ComRelease.Final(folder);
                }
            }

            return batch;
        }

        private void PublishProgress(string status, string currentFolder)
        {
            var handler = ProgressChanged;
            if (handler == null)
            {
                return;
            }

            handler(this, new IndexProgress
            {
                Processed = processed,
                EstimatedTotal = estimatedTotal,
                Failed = failed,
                CurrentFolder = currentFolder,
                Status = status,
                IsRunning = IsRunning
            });
        }

        private void RunOnUi(Action action)
        {
            var control = uiControl;
            if (control == null || control.IsDisposed)
            {
                return;
            }

            if (control.InvokeRequired)
            {
                try
                {
                    control.BeginInvoke(action);
                }
                catch (InvalidOperationException)
                {
                    // Outlook is shutting down or the pane handle has not been created.
                }
                return;
            }

            action();
        }

        private void ResetLog()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                File.WriteAllText(
                    logPath,
                    DateTime.UtcNow.ToString("o") + " Index scan started" + Environment.NewLine);
            }
            catch (IOException)
            {
                // Diagnostics must never prevent indexing.
            }
            catch (UnauthorizedAccessException)
            {
                // Diagnostics must never prevent indexing.
            }
        }

        private void EnsureSession()
        {
            if (session != null)
            {
                return;
            }

            // This legacy OOM path is intentionally reachable only through the
            // labelled Debug OOM button.  Production startup and native import
            // never acquire an Outlook NameSpace.
            StartupTrace.Step("DEBUG OOM only: BEGIN Application.Session acquisition");
            session = application.Session;
            StartupTrace.Step("DEBUG OOM only: END Application.Session acquisition");
        }

        private string ReadTrustState()
        {
            try
            {
                return application.IsTrusted ? "True" : "False";
            }
            catch (COMException ex)
            {
                return string.Format("error 0x{0:X8}", ex.ErrorCode);
            }
        }

        private void Log(string message)
        {
            try
            {
                File.AppendAllText(
                    logPath,
                    DateTime.UtcNow.ToString("o") + " " + message + Environment.NewLine);
            }
            catch (IOException)
            {
                // Diagnostics must never prevent indexing.
            }
            catch (UnauthorizedAccessException)
            {
                // Diagnostics must never prevent indexing.
            }
        }

        public void Dispose()
        {
            Stop();
            timer.Tick -= TimerOnTick;
            timer.Dispose();
        }

        private sealed class FolderCursor
        {
            public string StoreId { get; set; }
            public string StoreName { get; set; }
            public string FolderEntryId { get; set; }
            public string FolderPath { get; set; }
            public int NextIndex { get; set; }
        }
    }
}
