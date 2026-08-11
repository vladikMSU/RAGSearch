using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RAGSearch
{
    internal sealed class NativeSearchPresenter
    {
        // Instant Search silently produces an empty view for very long generated
        // queries on some Microsoft 365 builds. Keep the native filter compact.
        private const int MaximumResultClauses = 8;
        private const int MaximumAqsCharacters = 900;
        private readonly Outlook.Application application;

        public NativeSearchPresenter(Outlook.Application application)
        {
            this.application = application ?? throw new ArgumentNullException("application");
        }

        public int Show(IList<SearchResultDto> results)
        {
            if (results == null || results.Count == 0)
            {
                Clear();
                return 0;
            }

            // Outlook can force the add-in thread culture to the Office UI
            // language while Instant Search still parses keywords using the
            // Windows user's regional locale.  That is the case on the test
            // machine (English Outlook UI, Russian `тема:` parser).
            var russianSearch = IsRussianSearchLocale();
            var refined = BuildProjection(results, russianSearch, includeRefiners: true);
            var fallback = BuildProjection(results, russianSearch, includeRefiners: false);

            if (fallback.ClauseCount == 0)
            {
                throw new InvalidOperationException(
                    "У найденных элементов нет темы, поэтому штатный поиск Outlook не может представить их без записи служебного свойства в письма.");
            }

            Outlook.Explorer explorer = null;
            try
            {
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    throw new InvalidOperationException("Нет активного окна Outlook Explorer.");
                }

                try
                {
                    explorer.Search(
                        refined.Query,
                        Outlook.OlSearchScope.olSearchScopeAllOutlookItems);
                }
                catch (COMException)
                {
                    // Some Outlook/Windows Search combinations accept `subject:`
                    // but reject an additional localized refiner. Preserve a usable
                    // projection instead of turning the semantic result into an
                    // empty/error view. Instant Search has no results callback, so
                    // a syntactically accepted query cannot be checked here.
                    if (string.Equals(refined.Query, fallback.Query, StringComparison.Ordinal))
                    {
                        throw;
                    }

                    explorer.Search(
                        fallback.Query,
                        Outlook.OlSearchScope.olSearchScopeAllOutlookItems);
                }
                explorer.Activate();
                return refined.ClauseCount;
            }
            finally
            {
                ComRelease.Final(explorer);
            }
        }

        public void Clear()
        {
            Outlook.Explorer explorer = null;
            try
            {
                explorer = application.ActiveExplorer();
                if (explorer != null)
                {
                    explorer.ClearSearch();
                }
            }
            finally
            {
                ComRelease.Final(explorer);
            }
        }

        private static SearchProjection BuildProjection(
            IList<SearchResultDto> results,
            bool russianSearch,
            bool includeRefiners)
        {
            var clauses = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var queryLength = 0;
            var subjectKeyword = russianSearch ? "тема:" : "subject:";
            var fromKeyword = russianSearch ? "откого:" : "from:";
            var receivedKeyword = russianSearch ? "получено:" : "received:";
            var orOperator = russianSearch ? " ИЛИ " : " OR ";

            foreach (var result in results)
            {
                var subject = CleanAqsValue(result == null ? null : result.subject, 220);
                if (subject.Length == 0)
                {
                    continue;
                }

                var parts = new List<string>
                {
                    subjectKeyword + "\"" + subject + "\""
                };
                if (includeRefiners && result != null)
                {
                    var sender = SelectSearchableSender(result);
                    if (sender.Length > 0)
                    {
                        parts.Add(fromKeyword + "\"" + sender + "\"");
                    }

                    var receivedDate = FormatReceivedDate(result.received_at, russianSearch);
                    if (receivedDate.Length > 0)
                    {
                        parts.Add(receivedKeyword + receivedDate);
                    }
                }

                var clause = parts.Count == 1
                    ? parts[0]
                    : "(" + string.Join(" ", parts) + ")";
                if (!seen.Add(clause))
                {
                    continue;
                }

                var additionalLength = clause.Length +
                    (clauses.Count == 0 ? 0 : orOperator.Length);
                if (clauses.Count >= MaximumResultClauses ||
                    queryLength + additionalLength > MaximumAqsCharacters)
                {
                    break;
                }

                clauses.Add(clause);
                queryLength += additionalLength;
            }

            return new SearchProjection(
                string.Join(orOperator, clauses),
                clauses.Count);
        }

        private static string SelectSearchableSender(SearchResultDto result)
        {
            var email = CleanAqsValue(result.sender_email, 160);
            if (email.IndexOf('@') > 0)
            {
                return email;
            }

            return CleanAqsValue(result.sender_name, 160);
        }

        private static string FormatReceivedDate(string value, bool russianSearch)
        {
            DateTimeOffset parsed;
            if (string.IsNullOrWhiteSpace(value) ||
                !DateTimeOffset.TryParse(
                    value,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal,
                    out parsed))
            {
                return string.Empty;
            }

            var local = parsed.ToLocalTime();
            return local.ToString(
                russianSearch ? "dd.MM.yyyy" : "M/d/yyyy",
                CultureInfo.InvariantCulture);
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

        private sealed class SearchProjection
        {
            public SearchProjection(string query, int clauseCount)
            {
                Query = query;
                ClauseCount = clauseCount;
            }

            public string Query { get; private set; }
            public int ClauseCount { get; private set; }
        }

        private static bool IsRussianSearchLocale()
        {
            try
            {
                using (var international = Registry.CurrentUser.OpenSubKey(
                    @"Control Panel\International",
                    writable: false))
                {
                    var localeName = international == null
                        ? null
                        : international.GetValue("LocaleName") as string;
                    if (!string.IsNullOrWhiteSpace(localeName))
                    {
                        return localeName.StartsWith(
                            "ru",
                            StringComparison.OrdinalIgnoreCase);
                    }
                }
            }
            catch
            {
                // Fall back to CLR cultures when the registry is unavailable.
            }

            return string.Equals(
                       CultureInfo.InstalledUICulture.TwoLetterISOLanguageName,
                       "ru",
                       StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(
                       CultureInfo.CurrentCulture.TwoLetterISOLanguageName,
                       "ru",
                       StringComparison.OrdinalIgnoreCase);
        }
    }
}
