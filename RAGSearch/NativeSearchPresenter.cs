using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RAGSearch
{
    internal sealed class NativeFolderScope
    {
        public NativeFolderScope(string storeId, string folderEntryId, string folderName)
            : this(storeId, folderEntryId, folderName, string.Empty, -1)
        {
        }

        public NativeFolderScope(
            string storeId,
            string folderEntryId,
            string folderName,
            string viewName,
            long explorerStateVersion)
        {
            StoreId = storeId ?? string.Empty;
            FolderEntryId = folderEntryId ?? string.Empty;
            FolderName = folderName ?? string.Empty;
            ViewName = viewName ?? string.Empty;
            ExplorerStateVersion = explorerStateVersion;
        }

        public string StoreId { get; private set; }
        public string FolderEntryId { get; private set; }
        public string FolderName { get; private set; }
        public string ViewName { get; private set; }
        public long ExplorerStateVersion { get; private set; }
    }

    internal sealed class NativeFilterSummary
    {
        public string FolderName { get; set; }
        public int ResultCount { get; set; }
        public int AppliedClauseCount { get; set; }
        public int MessageIdClauseCount { get; set; }
        public int ApproximateClauseCount { get; set; }
        public int SkippedCount { get; set; }
        public int OutOfScopeCount { get; set; }
        public bool ProjectionWasTruncated { get; set; }
        public string ProjectionMode { get; set; }
    }

    internal sealed class NativeSearchPresenter : IDisposable
    {
        private const int MaximumSearchClauses = 12;
        private const int MaximumAqsCharacters = 2400;
        private const string EmptySearchSentinel =
            "__RAGSearch_No_Matches_8F43168D_6ECF_48D8_9A79_89D4D09938B2__";

        private readonly Outlook.Explorer explorer;
        private bool suppressExplorerEvents;
        private bool folderSwitchSubscribed;
        private bool viewSwitchSubscribed;
        private bool selectionChangeSubscribed;
        private bool closeSubscribed;
        private long explorerStateVersion;
        private bool disposed;

        public NativeSearchPresenter(Outlook.Explorer explorer)
        {
            this.explorer = explorer ?? throw new ArgumentNullException("explorer");
            SubscribeExplorerEvents();
        }

        public NativeFolderScope GetCurrentScope()
        {
            ThrowIfDisposed();

            Outlook.MAPIFolder folder = null;
            Outlook.View currentView = null;
            try
            {
                folder = explorer.CurrentFolder;
                if (folder == null)
                {
                    throw new InvalidOperationException("В Outlook не выбрана текущая папка.");
                }
                if (folder.DefaultItemType != Outlook.OlItemType.olMailItem)
                {
                    throw new InvalidOperationException(
                        "Откройте папку с письмами, чтобы запустить поиск All Mailboxes.");
                }

                var storeId = folder.StoreID;
                var folderEntryId = folder.EntryID;
                if (string.IsNullOrWhiteSpace(storeId) ||
                    string.IsNullOrWhiteSpace(folderEntryId))
                {
                    throw new InvalidOperationException(
                        "Outlook не вернул идентификаторы текущей папки.");
                }

                currentView = explorer.CurrentView as Outlook.View;
                if (currentView == null)
                {
                    throw new InvalidOperationException(
                        "Outlook не вернул текущее представление списка.");
                }

                return new NativeFolderScope(
                    storeId,
                    folderEntryId,
                    "All Mailboxes",
                    currentView.Name,
                    explorerStateVersion);
            }
            finally
            {
                ComRelease.Final(currentView);
                ComRelease.Final(folder);
            }
        }

        public NativeFilterSummary Show(
            NativeFolderScope requestedScope,
            IList<SearchResultDto> results)
        {
            ThrowIfDisposed();
            if (requestedScope == null)
            {
                throw new ArgumentNullException("requestedScope");
            }

            Outlook.MAPIFolder folder = null;
            try
            {
                folder = explorer.CurrentFolder;
                EnsureSameExplorerState(requestedScope, folder);
                var projection = BuildSearchProjection(results);

                var previousSuppression = suppressExplorerEvents;
                suppressExplorerEvents = true;
                try
                {
                    explorer.Search(
                        projection.Query,
                        Outlook.OlSearchScope.olSearchScopeAllFolders);
                }
                finally
                {
                    suppressExplorerEvents = previousSuppression;
                }

                return new NativeFilterSummary
                {
                    FolderName = "All Mailboxes",
                    ResultCount = results == null ? 0 : results.Count,
                    AppliedClauseCount = projection.ClauseCount,
                    MessageIdClauseCount = 0,
                    ApproximateClauseCount = projection.ClauseCount,
                    SkippedCount = projection.SkippedCount,
                    OutOfScopeCount = 0,
                    ProjectionWasTruncated = projection.Truncated,
                    ProjectionMode = results == null || results.Count == 0
                        ? "empty"
                        : "quoted-phrases"
                };
            }
            finally
            {
                ComRelease.Final(folder);
            }
        }

        public void Clear()
        {
            ThrowIfDisposed();
            var previousSuppression = suppressExplorerEvents;
            suppressExplorerEvents = true;
            try
            {
                explorer.ClearSearch();
            }
            finally
            {
                suppressExplorerEvents = previousSuppression;
            }
        }

        public void Dispose()
        {
            if (disposed)
            {
                UnsubscribeExplorerEvents();
                return;
            }

            UnsubscribeExplorerEvents();
            disposed = true;
        }

        private static SearchProjection BuildSearchProjection(
            IList<SearchResultDto> results)
        {
            var clauses = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var queryLength = 0;
            var skipped = 0;
            var truncated = false;
            const string orOperator = " OR ";

            if (results != null)
            {
                foreach (var result in results)
                {
                    var subject = CleanAqsValue(
                        result == null ? null : result.subject,
                        220);
                    if (subject.Length == 0)
                    {
                        skipped++;
                        continue;
                    }

                    // Plain quoted phrases are intentional. Property keywords
                    // (`subject:`/`тема:` and canonical System.Subject) were
                    // locale/provider-dependent on the target Outlook build.
                    var clause = "\"" + subject + "\"";
                    if (!seen.Add(clause))
                    {
                        continue;
                    }

                    var additionalLength = clause.Length +
                        (clauses.Count == 0 ? 0 : orOperator.Length);
                    if (clauses.Count >= MaximumSearchClauses ||
                        queryLength + additionalLength > MaximumAqsCharacters)
                    {
                        skipped++;
                        truncated = true;
                        continue;
                    }

                    clauses.Add(clause);
                    queryLength += additionalLength;
                }
            }

            if (clauses.Count == 0)
            {
                return new SearchProjection(
                    "\"" + EmptySearchSentinel + "\"",
                    0,
                    skipped,
                    false);
            }

            return new SearchProjection(
                string.Join(orOperator, clauses),
                clauses.Count,
                skipped,
                truncated);
        }

        private static string CleanAqsValue(string value, int maximumLength)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(Math.Min(value.Length, maximumLength));
            var previousWasSpace = false;
            foreach (var character in value)
            {
                if (builder.Length >= maximumLength)
                {
                    break;
                }

                var current = character == '"' || char.IsControl(character)
                    ? ' '
                    : character;
                if (char.IsWhiteSpace(current))
                {
                    if (!previousWasSpace)
                    {
                        builder.Append(' ');
                        previousWasSpace = true;
                    }
                }
                else
                {
                    builder.Append(current);
                    previousWasSpace = false;
                }
            }

            return builder.ToString().Trim();
        }

        private void EnsureSameExplorerState(
            NativeFolderScope requestedScope,
            Outlook.MAPIFolder folder)
        {
            if (folder == null ||
                !string.Equals(
                    folder.StoreID,
                    requestedScope.StoreId,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    folder.EntryID,
                    requestedScope.FolderEntryId,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Текущая папка Outlook изменилась во время поиска. Повторите запрос.");
            }
            if (requestedScope.ExplorerStateVersion >= 0 &&
                requestedScope.ExplorerStateVersion != explorerStateVersion)
            {
                throw new InvalidOperationException(
                    "Список Outlook изменился во время поиска. Повторите запрос.");
            }

            Outlook.View currentView = null;
            try
            {
                currentView = explorer.CurrentView as Outlook.View;
                if (currentView == null || !string.Equals(
                    currentView.Name,
                    requestedScope.ViewName,
                    StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Представление Outlook изменилось во время поиска. Повторите запрос.");
                }
            }
            finally
            {
                ComRelease.Final(currentView);
            }
        }

        private void SubscribeExplorerEvents()
        {
            try
            {
                explorer.FolderSwitch += ExplorerOnFolderSwitch;
                folderSwitchSubscribed = true;
                explorer.ViewSwitch += ExplorerOnViewSwitch;
                viewSwitchSubscribed = true;
                explorer.SelectionChange += ExplorerOnSelectionChange;
                selectionChangeSubscribed = true;
                ((Outlook.ExplorerEvents_10_Event)explorer).Close += ExplorerOnClose;
                closeSubscribed = true;
            }
            catch
            {
                UnsubscribeExplorerEvents();
                throw;
            }
        }

        private void UnsubscribeExplorerEvents()
        {
            if (folderSwitchSubscribed)
            {
                try
                {
                    explorer.FolderSwitch -= ExplorerOnFolderSwitch;
                    folderSwitchSubscribed = false;
                }
                catch (Exception ex)
                {
                    StartupTrace.Failure("unsubscribe Explorer.FolderSwitch", ex);
                }
            }
            if (viewSwitchSubscribed)
            {
                try
                {
                    explorer.ViewSwitch -= ExplorerOnViewSwitch;
                    viewSwitchSubscribed = false;
                }
                catch (Exception ex)
                {
                    StartupTrace.Failure("unsubscribe Explorer.ViewSwitch", ex);
                }
            }
            if (selectionChangeSubscribed)
            {
                try
                {
                    explorer.SelectionChange -= ExplorerOnSelectionChange;
                    selectionChangeSubscribed = false;
                }
                catch (Exception ex)
                {
                    StartupTrace.Failure("unsubscribe Explorer.SelectionChange", ex);
                }
            }
            if (closeSubscribed)
            {
                try
                {
                    ((Outlook.ExplorerEvents_10_Event)explorer).Close -= ExplorerOnClose;
                    closeSubscribed = false;
                }
                catch (Exception ex)
                {
                    StartupTrace.Failure("unsubscribe Explorer.Close", ex);
                }
            }
        }

        private void ExplorerOnFolderSwitch()
        {
            MarkExplorerStateChanged();
        }

        private void ExplorerOnViewSwitch()
        {
            MarkExplorerStateChanged();
        }

        private void ExplorerOnSelectionChange()
        {
            MarkExplorerStateChanged();
        }

        private void MarkExplorerStateChanged()
        {
            if (!disposed && !suppressExplorerEvents)
            {
                explorerStateVersion++;
            }
        }

        private void ExplorerOnClose()
        {
            Dispose();
        }

        private void ThrowIfDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException("NativeSearchPresenter");
            }
        }

        private sealed class SearchProjection
        {
            public SearchProjection(
                string query,
                int clauseCount,
                int skippedCount,
                bool truncated)
            {
                Query = query;
                ClauseCount = clauseCount;
                SkippedCount = skippedCount;
                Truncated = truncated;
            }

            public string Query { get; private set; }
            public int ClauseCount { get; private set; }
            public int SkippedCount { get; private set; }
            public bool Truncated { get; private set; }
        }
    }
}
