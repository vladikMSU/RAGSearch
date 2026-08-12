using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RAGSearch
{
    internal sealed class SearchPaneControl : UserControl
    {
        private readonly LocalServiceClient serviceClient;
        private readonly OomGuardProbe oomGuardProbe;
        private readonly NativeImportRunner nativeImportRunner;
        private readonly Func<NativeFolderScope> getNativeSearchScope;
        private readonly Func<NativeFolderScope, IList<SearchResultDto>, NativeFilterSummary> showNativeResults;
        private readonly Action clearNativeSearch;
        private readonly TextBox queryBox;
        private readonly Button searchButton;
        private readonly Button clearButton;
        private readonly Button indexButton;
        private readonly Button stopButton;
        private readonly Button resetIndexButton;
        private readonly Button debugOomButton;
        private readonly Label statusLabel;
        private CancellationTokenSource searchCancellation;
        private CancellationTokenSource resetCancellation;
        private int searchGeneration;
        private bool probeRunning;
        private bool resetRunning;

        public SearchPaneControl(
            LocalServiceClient serviceClient,
            OomGuardProbe oomGuardProbe,
            NativeImportRunner nativeImportRunner,
            Func<NativeFolderScope> getNativeSearchScope,
            Func<NativeFolderScope, IList<SearchResultDto>, NativeFilterSummary> showNativeResults,
            Action clearNativeSearch)
        {
            this.serviceClient = serviceClient ?? throw new ArgumentNullException("serviceClient");
            this.oomGuardProbe = oomGuardProbe ?? throw new ArgumentNullException("oomGuardProbe");
            this.nativeImportRunner = nativeImportRunner ?? throw new ArgumentNullException("nativeImportRunner");
            this.getNativeSearchScope = getNativeSearchScope ?? throw new ArgumentNullException("getNativeSearchScope");
            this.showNativeResults = showNativeResults ?? throw new ArgumentNullException("showNativeResults");
            this.clearNativeSearch = clearNativeSearch ?? throw new ArgumentNullException("clearNativeSearch");

            Dock = DockStyle.Fill;
            BackColor = SystemColors.Window;
            MinimumSize = new Size(600, 96);

            queryBox = new TextBox
            {
                Width = 520,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(8, 7, 5, 3)
            };
            queryBox.KeyDown += QueryBoxOnKeyDown;

            searchButton = CreateButton("Семантический поиск");
            searchButton.Click += SearchButtonOnClick;

            clearButton = CreateButton("Сбросить фильтр");
            clearButton.Click += ClearButtonOnClick;

            indexButton = CreateButton("Индексировать PST + OST (MAPI)");
            indexButton.Click += IndexButtonOnClick;

            stopButton = CreateButton("Стоп");
            stopButton.Enabled = false;
            stopButton.Click += StopButtonOnClick;

            resetIndexButton = CreateButton("Очистить базу");
            resetIndexButton.Click += ResetIndexButtonOnClick;

            debugOomButton = CreateButton("Debug OOM: 1 письмо → Guard");
            debugOomButton.Click += DebugOomButtonOnClick;

            statusLabel = new Label
            {
                AutoEllipsis = true,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = SystemColors.GrayText,
                Padding = new Padding(8, 2, 4, 0),
                Text = "Локальный сервис: проверка..."
            };

            var searchRow = CreateRow();
            searchRow.Controls.Add(queryBox);
            searchRow.Controls.Add(searchButton);
            searchRow.Controls.Add(clearButton);

            var toolsRow = CreateRow();
            toolsRow.Controls.Add(indexButton);
            toolsRow.Controls.Add(stopButton);
            toolsRow.Controls.Add(resetIndexButton);
            toolsRow.Controls.Add(debugOomButton);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = new Padding(0),
                // Current Microsoft 365 builds draw the vertical app rail over the
                // left edge of a top task pane. Keep controls out from under it.
                Padding = new Padding(52, 0, 0, 0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.Controls.Add(searchRow, 0, 0);
            layout.Controls.Add(toolsRow, 0, 1);
            layout.Controls.Add(statusLabel, 0, 2);
            Controls.Add(layout);

            nativeImportRunner.ProgressChanged += NativeImportRunnerOnProgressChanged;
            Load += async (sender, args) =>
            {
                StartupTrace.Step("BEGIN SearchPaneControl.Load (loopback health only)");
                try
                {
                    await RefreshHealthAsync();
                    StartupTrace.Step("END SearchPaneControl.Load");
                }
                catch (Exception ex)
                {
                    StartupTrace.Failure("SearchPaneControl.Load", ex);
                    throw;
                }
            };
        }

        private static FlowLayoutPanel CreateRow()
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = SystemColors.ControlLightLight,
                Padding = new Padding(0)
            };
        }

        private static Button CreateButton(string text)
        {
            return new Button
            {
                AutoSize = true,
                FlatStyle = FlatStyle.System,
                Font = new Font("Segoe UI", 9F),
                Margin = new Padding(3, 5, 3, 3),
                Padding = new Padding(5, 1, 5, 1),
                Text = text,
                UseVisualStyleBackColor = true
            };
        }

        private async void SearchButtonOnClick(object sender, EventArgs eventArgs)
        {
            await SearchAsync();
        }

        private async Task SearchAsync()
        {
            if (probeRunning || resetRunning)
            {
                return;
            }

            var query = queryBox.Text.Trim();
            if (query.Length == 0)
            {
                statusLabel.Text = "Введите запрос. Можно описать смысл, точная подстрока не обязательна.";
                return;
            }

            if (searchCancellation != null)
            {
                searchCancellation.Cancel();
            }

            var currentCancellation = new CancellationTokenSource();
            var currentGeneration = ++searchGeneration;
            searchCancellation = currentCancellation;
            queryBox.Enabled = false;
            searchButton.Enabled = false;
            debugOomButton.Enabled = false;
            UpdateResetIndexButtonEnabled();
            statusLabel.Text = "Векторный + полнотекстовый поиск...";
            try
            {
                var scope = getNativeSearchScope();
                var response = await serviceClient.SearchAsync(
                    query,
                    12,
                    currentCancellation.Token);
                currentCancellation.Token.ThrowIfCancellationRequested();
                if (probeRunning || currentGeneration != searchGeneration)
                {
                    throw new OperationCanceledException(currentCancellation.Token);
                }
                var results = response == null || response.results == null
                    ? new List<SearchResultDto>()
                    : response.results;

                var summary = showNativeResults(scope, results);
                statusLabel.Text = FormatNativeFilterStatus(summary, response);
            }
            catch (OperationCanceledException)
            {
                if (!probeRunning && currentGeneration == searchGeneration)
                {
                    statusLabel.Text = "Поиск отменён";
                }
            }
            catch (Exception ex)
            {
                if (!probeRunning && currentGeneration == searchGeneration)
                {
                    statusLabel.Text = "Поиск не выполнен: " + ex.Message;
                }
            }
            finally
            {
                if (ReferenceEquals(searchCancellation, currentCancellation))
                {
                    searchCancellation = null;
                    RunOnUi(() =>
                    {
                        queryBox.Enabled = !probeRunning && !resetRunning;
                        searchButton.Enabled = !probeRunning;
                        debugOomButton.Enabled = !probeRunning &&
                                                       !nativeImportRunner.IsRunning;
                        UpdateResetIndexButtonEnabled();
                    });
                }
                currentCancellation.Dispose();
            }
        }

        private static string FormatNativeFilterStatus(
            NativeFilterSummary summary,
            SearchResponse response)
        {
            if (summary == null)
            {
                return "Поиск Outlook All Mailboxes запущен.";
            }

            var hasRankingDetails = response != null &&
                                    string.Equals(
                                        response.ranking,
                                        "lexical_gate_then_vector_distance_asc",
                                        StringComparison.Ordinal);
            var rankingText = string.Empty;
            if (hasRankingDetails && response.lexical_gate)
            {
                rankingText = string.Format(
                    "Literal-совпадений: {0}; сервис вернул: {1}; top-N: {2}. ",
                    response.lexical_match_count,
                    response.total,
                    response.max_results);
            }
            else if (hasRankingDetails &&
                     string.Equals(
                         response.mode,
                         "single-token-no-literal",
                         StringComparison.Ordinal))
            {
                rankingText =
                    "Literal-совпадений нет; для чисто семантического поиска введите фразу из двух или более слов. ";
            }
            else if (hasRankingDetails)
            {
                rankingText = string.Format(
                    "Кандидатов: {0}; cutoff d≤{1}; прошло: {2}; literal: {3}; top-N: {4}. ",
                    response.candidate_count,
                    response.cutoff_distance.ToString("0.000"),
                    response.eligible_count,
                    response.lexical_match_count,
                    response.max_results);
            }
            if (summary.ResultCount == 0)
            {
                return rankingText +
                    "Релевантных совпадений нет; агрегированный список All Mailboxes оставлен пустым до сброса.";
            }

            if (summary.AppliedClauseCount == 0)
            {
                return string.Format(
                    "{0}Сервис вернул {1}, но ни один результат нельзя представить через Outlook Instant Search; список оставлен пустым.",
                    rankingText,
                    summary.ResultCount);
            }

            var text = string.Format(
                "{0}Сервис вернул {1}; в All Mailboxes передано условий: {2}. Поиск охватывает выбранные в Outlook почтовые хранилища и архивы.",
                rankingText,
                summary.ResultCount,
                summary.AppliedClauseCount);
            if (hasRankingDetails)
            {
                text += response.lexical_gate
                    ? " Сервис отдаёт literal-совпадения первыми; native-список сохраняет порядок Outlook."
                    : " Векторная часть сервиса ранжирована по distance (d↑); native-список сохраняет порядок Outlook.";
            }
            if (summary.ApproximateClauseCount > 0)
            {
                text += " Проекция приблизительная: Outlook AQS ищет кавыченные фразы из тем результатов (в том числе в теле письма), а не EntryID.";
            }
            if (summary.SkippedCount > 0)
            {
                text += string.Format(
                    " Не удалось представить в view: {0}.",
                    summary.SkippedCount);
            }
            if (summary.ProjectionWasTruncated)
            {
                text += " Технический запрос укорочен до безопасного лимита Outlook.";
            }
            return text;
        }

        private void ClearButtonOnClick(object sender, EventArgs eventArgs)
        {
            try
            {
                searchGeneration++;
                if (searchCancellation != null)
                {
                    searchCancellation.Cancel();
                }
                clearNativeSearch();
                statusLabel.Text = "Поиск Outlook All Mailboxes сброшен.";
            }
            catch (Exception ex)
            {
                statusLabel.Text = "Не удалось сбросить фильтр Outlook: " + ex.Message;
            }
        }

        private async void IndexButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (resetRunning)
            {
                statusLabel.Text = "Дождитесь завершения очистки локального индекса.";
                return;
            }
            if (nativeImportRunner.IsRunning)
            {
                statusLabel.Text = "Индексация уже выполняется.";
                return;
            }

            indexButton.Enabled = false;
            debugOomButton.Enabled = false;
            UpdateResetIndexButtonEnabled();
            stopButton.Enabled = true;
            statusLabel.Text = "Готовлю read-only Extended MAPI индексацию...";
            try
            {
                await nativeImportRunner.RunAsync();
            }
            catch (OperationCanceledException)
            {
                RunOnUi(() => statusLabel.Text = "Native-индексация остановлена пользователем.");
            }
            catch (Exception ex)
            {
                var message = "Native-индексация не выполнена: " + ex.Message;
                RunOnUi(() => statusLabel.Text = message);
            }
            finally
            {
                RunOnUi(() =>
                {
                    indexButton.Enabled = !resetRunning;
                    debugOomButton.Enabled = !resetRunning && searchCancellation == null;
                    stopButton.Enabled = false;
                    UpdateResetIndexButtonEnabled();
                });
            }
        }

        private void StopButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (nativeImportRunner.IsRunning)
            {
                nativeImportRunner.RequestStop();
                stopButton.Enabled = false;
                return;
            }
        }

        private void DebugOomButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (resetRunning)
            {
                statusLabel.Text = "Дождитесь завершения очистки локального индекса.";
                return;
            }
            if (probeRunning || nativeImportRunner.IsRunning)
            {
                statusLabel.Text = "Сначала остановите текущую индексацию.";
                return;
            }
            if (searchCancellation != null)
            {
                statusLabel.Text = "Дождитесь завершения семантического поиска перед Debug OOM.";
                return;
            }

            probeRunning = true;
            queryBox.Enabled = false;
            searchButton.Enabled = false;
            clearButton.Enabled = false;
            indexButton.Enabled = false;
            debugOomButton.Enabled = false;
            UpdateResetIndexButtonEnabled();
            stopButton.Enabled = false;
            try
            {
                // Keep this synchronous on the Outlook UI thread.  The probe
                // locates one MailItem and immediately calls one documented
                // protected getter; it does not depend on the service or Timer.
                var result = oomGuardProbe.Run(message =>
                {
                    statusLabel.Text = message;
                    statusLabel.Update();
                });
                statusLabel.Text = result;
            }
            catch (Exception ex)
            {
                statusLabel.Text = "Debug OOM завершился внутренней ошибкой: " + ex.GetType().Name;
            }
            finally
            {
                probeRunning = false;
                queryBox.Enabled = true;
                searchButton.Enabled = true;
                clearButton.Enabled = true;
                indexButton.Enabled = true;
                debugOomButton.Enabled = true;
                stopButton.Enabled = false;
                UpdateResetIndexButtonEnabled();
            }
        }

        private async void ResetIndexButtonOnClick(object sender, EventArgs eventArgs)
        {
            if (resetRunning)
            {
                return;
            }
            if (nativeImportRunner.IsRunning)
            {
                statusLabel.Text = "Сначала остановите текущую индексацию.";
                return;
            }
            if (searchCancellation != null)
            {
                statusLabel.Text = "Дождитесь завершения семантического поиска.";
                return;
            }
            if (probeRunning)
            {
                statusLabel.Text = "Дождитесь завершения Debug OOM.";
                return;
            }

            var confirmation = MessageBox.Show(
                this,
                "Удалить весь локальный поисковый индекс RAGSearch?\r\n\r\n" +
                "Будут удалены только данные локальной базы: проиндексированные сообщения, " +
                "вложения и векторные чанки. Письма и вложения в Outlook, PST и OST " +
                "не изменятся.\r\n\r\nДля поиска потребуется заново выполнить индексацию.",
                "RAGSearch: очистка локального индекса",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (confirmation != DialogResult.Yes)
            {
                return;
            }

            // A modal dialog pumps messages. Recheck all activity immediately
            // before issuing the destructive local-service request.
            if (nativeImportRunner.IsRunning || searchCancellation != null || probeRunning)
            {
                statusLabel.Text = "Очистка отменена: другая операция уже выполняется.";
                UpdateResetIndexButtonEnabled();
                return;
            }

            var currentCancellation = new CancellationTokenSource();
            resetCancellation = currentCancellation;
            resetRunning = true;
            queryBox.Enabled = false;
            searchButton.Enabled = false;
            clearButton.Enabled = false;
            indexButton.Enabled = false;
            stopButton.Enabled = false;
            debugOomButton.Enabled = false;
            resetIndexButton.Enabled = false;
            statusLabel.Text = "Проверяю локальный сервис перед очисткой индекса...";
            var completionStatus = "Очистка локального индекса завершилась без результата.";
            try
            {
                await nativeImportRunner.EnsureServiceReadyAsync(currentCancellation.Token);
                currentCancellation.Token.ThrowIfCancellationRequested();
                var response = await serviceClient.ResetIndexAsync(currentCancellation.Token);
                currentCancellation.Token.ThrowIfCancellationRequested();
                completionStatus = response == null
                    ? "Локальный индекс очищен; сервис не вернул счётчики."
                    : string.Format(
                        "Локальный индекс очищен: сообщений {0}, вложений {1}, чанков {2}.",
                        response.deleted_messages,
                        response.deleted_attachments,
                        response.deleted_chunks);
            }
            catch (OperationCanceledException)
            {
                completionStatus = "Очистка локального индекса отменена.";
            }
            catch (Exception ex)
            {
                completionStatus = "Не удалось очистить локальный индекс: " + ex.Message;
            }
            finally
            {
                if (ReferenceEquals(resetCancellation, currentCancellation))
                {
                    resetCancellation = null;
                }
                currentCancellation.Dispose();
                RunOnUi(() =>
                {
                    resetRunning = false;
                    statusLabel.Text = completionStatus;
                    queryBox.Enabled = true;
                    clearButton.Enabled = true;
                    searchButton.Enabled = searchCancellation == null && !probeRunning;
                    var indexing = nativeImportRunner.IsRunning;
                    indexButton.Enabled = !indexing && !probeRunning;
                    stopButton.Enabled = indexing;
                    debugOomButton.Enabled = !indexing &&
                                                   !probeRunning &&
                                                   searchCancellation == null;
                    UpdateResetIndexButtonEnabled();
                });
            }
        }

        private void QueryBoxOnKeyDown(object sender, KeyEventArgs eventArgs)
        {
            if (eventArgs.KeyCode != Keys.Enter)
            {
                return;
            }

            eventArgs.SuppressKeyPress = true;
            if (!searchButton.Enabled)
            {
                return;
            }
            SearchButtonOnClick(sender, EventArgs.Empty);
        }

        private void NativeImportRunnerOnProgressChanged(object sender, NativeImportProgress progress)
        {
            RunOnUi(() =>
            {
                var active = nativeImportRunner.IsRunning;
                stopButton.Enabled = progress.IsRunning && active;
                indexButton.Enabled = !active && !resetRunning;
                debugOomButton.Enabled = !active && !resetRunning;
                UpdateResetIndexButtonEnabled();
                statusLabel.Text = progress.Status;
            });
        }

        private void UpdateResetIndexButtonEnabled()
        {
            resetIndexButton.Enabled = !resetRunning &&
                                       !probeRunning &&
                                       searchCancellation == null &&
                                       !nativeImportRunner.IsRunning;
        }

        private async Task RefreshHealthAsync()
        {
            try
            {
                var health = await serviceClient.GetHealthAsync(CancellationToken.None);
                statusLabel.Text = string.Format(
                    "Локальный сервис готов. Embeddings: {0}",
                    health == null ? "unknown" : health.embedding_backend ?? health.status ?? "unknown");
            }
            catch (Exception)
            {
                statusLabel.Text = "Python-сервис не запущен. Запустите python service\\run.py.";
            }
        }

        private void RunOnUi(Action action)
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(action);
                }
                catch (InvalidOperationException)
                {
                    // Outlook is closing or the task pane handle is gone.
                }
                return;
            }

            action();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                nativeImportRunner.ProgressChanged -= NativeImportRunnerOnProgressChanged;
                if (nativeImportRunner.IsRunning)
                {
                    nativeImportRunner.RequestStop();
                }
                if (searchCancellation != null)
                {
                    var cancellation = searchCancellation;
                    searchCancellation = null;
                    cancellation.Cancel();
                }
                if (resetCancellation != null)
                {
                    var cancellation = resetCancellation;
                    resetCancellation = null;
                    cancellation.Cancel();
                }
            }

            base.Dispose(disposing);
        }
    }
}
