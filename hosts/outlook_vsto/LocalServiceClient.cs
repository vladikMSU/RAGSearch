using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace RAGSearch
{
    internal sealed class LocalServiceClient : IDisposable
    {
        private const string TokenHeader = "X-RAGSearch-Token";
        private const int MaximumRequestBytes = 48 * 1024 * 1024;
        private const int MaximumResponseBytes = 128 * 1024 * 1024;
        private const int MaximumSearchResponseChars = 128 * 1024 * 1024;
        private static readonly Encoding StrictUtf8 = new UTF8Encoding(false, true);
        private readonly HttpClient httpClient;
        private readonly string tokenPath;

        public Uri ServiceUri
        {
            get { return httpClient.BaseAddress; }
        }

        public LocalServiceClient()
        {
            httpClient = new HttpClient
            {
                BaseAddress = new Uri("http://127.0.0.1:8765/", UriKind.Absolute),
                Timeout = TimeSpan.FromMinutes(5)
            };
            tokenPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RAGSearch",
                "service-token");
        }

        public async Task<HealthResponse> GetHealthAsync(CancellationToken cancellationToken)
        {
            return await SendAsync<HealthResponse>(
                    HttpMethod.Get,
                    "health",
                    null,
                    cancellationToken,
                    false)
                .ConfigureAwait(true);
        }

        public async Task<SearchResponse> SearchAsync(
            string query,
            int limit,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                throw new ArgumentException("Search query must not be empty.", "query");
            }
            if (limit < 1 || limit > 100)
            {
                throw new ArgumentOutOfRangeException("limit", "Search limit must be between 1 and 100.");
            }
            var request = new SearchRequest
            {
                Query = query,
                Limit = limit
            };
            return await SendAsync<SearchResponse>(
                    HttpMethod.Post,
                    "v1/search",
                    request,
                    cancellationToken)
                .ConfigureAwait(true);
        }

        public async Task<DocumentUpsertResponse> UpsertDocumentAsync(
            DocumentDto document,
            CancellationToken cancellationToken)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }
            return await SendAsync<DocumentUpsertResponse>(
                    HttpMethod.Post,
                    "v1/documents",
                    document,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        public async Task<ResetIndexResponse> ResetIndexAsync(CancellationToken cancellationToken)
        {
            return await SendAsync<ResetIndexResponse>(
                    HttpMethod.Delete,
                    "v1/index",
                    null,
                    cancellationToken)
                .ConfigureAwait(true);
        }

        private async Task<T> SendAsync<T>(
            HttpMethod method,
            string relativeUrl,
            object payload,
            CancellationToken cancellationToken,
            bool requireToken = true)
        {
            using (var request = new HttpRequestMessage(method, relativeUrl))
            {
                if (requireToken)
                {
                    var token = ReadToken();
                    request.Headers.TryAddWithoutValidation(TokenHeader, token);
                }

                if (payload != null)
                {
                    var json = SerializeUtf8(payload);
                    if (json.Length > MaximumRequestBytes)
                    {
                        throw new InvalidDataException(
                            "RAGSearch request exceeds the 48 MiB service contract.");
                    }
                    request.Content = new ByteArrayContent(json);
                    request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json")
                    {
                        CharSet = "utf-8"
                    };
                }

                using (var response = await httpClient.SendAsync(
                           request,
                           HttpCompletionOption.ResponseHeadersRead,
                           cancellationToken).ConfigureAwait(true))
                {
                    var responseBody = await ReadResponseBodyAsync(
                            response.Content,
                            cancellationToken)
                        .ConfigureAwait(true);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            string.Format(
                                "RAGSearch service returned HTTP {0}: {1}",
                                (int)response.StatusCode,
                                Compact(responseBody)));
                    }

                    if (string.IsNullOrWhiteSpace(responseBody))
                    {
                        throw new InvalidDataException(
                            "RAGSearch service returned an empty success response.");
                    }

                    return typeof(T) == typeof(SearchResponse)
                        ? (T)(object)DeserializeSearchResponse(responseBody)
                        : Deserialize<T>(responseBody);
                }
            }
        }

        private string ReadToken()
        {
            if (!File.Exists(tokenPath))
            {
                throw new FileNotFoundException(
                    "RAGSearch service token does not exist.",
                    tokenPath);
            }
            var token = File.ReadAllText(tokenPath).Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidDataException("RAGSearch service token is empty.");
            }
            return token;
        }

        private static byte[] SerializeUtf8(object payload)
        {
            var serializer = CreateSerializer(payload.GetType());
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, payload);
                return stream.ToArray();
            }
        }

        private static async Task<string> ReadResponseBodyAsync(
            HttpContent content,
            CancellationToken cancellationToken)
        {
            if (content == null)
            {
                return string.Empty;
            }

            var declaredLength = content.Headers.ContentLength;
            if (declaredLength.HasValue && declaredLength.Value > MaximumResponseBytes)
            {
                throw new InvalidDataException(
                    "RAGSearch response exceeds the 128 MiB client limit.");
            }

            var initialCapacity = declaredLength.HasValue
                ? checked((int)declaredLength.Value)
                : 0;
            using (var input = await content.ReadAsStreamAsync().ConfigureAwait(true))
            using (var output = new MemoryStream(initialCapacity))
            {
                var buffer = new byte[81920];
                var total = 0;
                while (true)
                {
                    var count = await input.ReadAsync(
                            buffer,
                            0,
                            buffer.Length,
                            cancellationToken)
                        .ConfigureAwait(true);
                    if (count == 0)
                    {
                        break;
                    }
                    if (count > MaximumResponseBytes - total)
                    {
                        throw new InvalidDataException(
                            "RAGSearch response exceeds the 128 MiB client limit.");
                    }
                    output.Write(buffer, 0, count);
                    total += count;
                }

                ArraySegment<byte> bytes;
                if (!output.TryGetBuffer(out bytes))
                {
                    bytes = new ArraySegment<byte>(output.ToArray());
                }
                try
                {
                    return StrictUtf8.GetString(bytes.Array, bytes.Offset, total);
                }
                catch (DecoderFallbackException exception)
                {
                    throw new InvalidDataException(
                        "RAGSearch service returned a response that is not valid UTF-8.",
                        exception);
                }
            }
        }

        private static T Deserialize<T>(string json)
        {
            var serializer = CreateSerializer(typeof(T));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(stream);
            }
        }

        private static SearchResponse DeserializeSearchResponse(string json)
        {
            object parsed;
            try
            {
                parsed = new JavaScriptSerializer
                {
                    MaxJsonLength = MaximumSearchResponseChars,
                    RecursionLimit = 32
                }.DeserializeObject(json);
            }
            catch (ArgumentException exception)
            {
                throw new InvalidDataException(
                    "RAGSearch service returned an invalid search response.",
                    exception);
            }
            catch (InvalidOperationException exception)
            {
                throw new InvalidDataException(
                    "RAGSearch service returned an invalid search response.",
                    exception);
            }

            var root = RequireObject(parsed, "search response");
            object rawResults;
            if (!root.TryGetValue("results", out rawResults))
            {
                throw new InvalidDataException("RAGSearch search response has no results array.");
            }
            var resultItems = rawResults as object[];
            if (resultItems == null)
            {
                throw new InvalidDataException("RAGSearch search response results must be an array.");
            }

            var results = new List<SearchResultDto>(resultItems.Length);
            for (var index = 0; index < resultItems.Length; index++)
            {
                var item = RequireObject(resultItems[index], "search result");
                results.Add(new SearchResultDto
                {
                    SourceKey = ReadString(item, "source_key"),
                    Kind = ReadString(item, "kind"),
                    Title = ReadString(item, "title"),
                    Metadata = ReadObject(item, "metadata"),
                    Locator = ReadObject(item, "locator"),
                    Rank = ReadInt32(item, "rank"),
                    Snippet = ReadString(item, "snippet"),
                    SnippetPart = ReadString(item, "snippet_part"),
                    MatchedParts = ReadStringList(item, "matched_parts")
                });
            }

            return new SearchResponse
            {
                Results = results,
                Mode = ReadString(root, "mode")
            };
        }

        private static Dictionary<string, object> RequireObject(object value, string name)
        {
            var result = value as Dictionary<string, object>;
            if (result == null)
            {
                throw new InvalidDataException(name + " must be a JSON object.");
            }
            return result;
        }

        private static Dictionary<string, object> ReadObject(
            Dictionary<string, object> values,
            string name)
        {
            object value;
            if (!values.TryGetValue(name, out value))
            {
                throw new InvalidDataException("Search result has no " + name + " object.");
            }
            return RequireObject(value, "Search result " + name);
        }

        private static string ReadString(Dictionary<string, object> values, string name)
        {
            object value;
            if (!values.TryGetValue(name, out value) || !(value is string))
            {
                throw new InvalidDataException("Search response field " + name + " must be a string.");
            }
            return (string)value;
        }

        private static int ReadInt32(Dictionary<string, object> values, string name)
        {
            object value;
            if (!values.TryGetValue(name, out value) || !(value is int))
            {
                throw new InvalidDataException("Search response field " + name + " must be an integer.");
            }
            return (int)value;
        }

        private static List<string> ReadStringList(
            Dictionary<string, object> values,
            string name)
        {
            object value;
            if (!values.TryGetValue(name, out value))
            {
                throw new InvalidDataException("Search result has no " + name + " array.");
            }
            var items = value as object[];
            if (items == null)
            {
                throw new InvalidDataException("Search response field " + name + " must be an array.");
            }
            var result = new List<string>(items.Length);
            foreach (var item in items)
            {
                var text = item as string;
                if (text == null)
                {
                    throw new InvalidDataException(
                        "Search response field " + name + " must contain only strings.");
                }
                result.Add(text);
            }
            return result;
        }

        private static DataContractJsonSerializer CreateSerializer(Type type)
        {
            return new DataContractJsonSerializer(
                type,
                new DataContractJsonSerializerSettings
                {
                    UseSimpleDictionaryFormat = true
                });
        }

        private static string Compact(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "empty response";
            }

            var compact = value.Replace('\r', ' ').Replace('\n', ' ').Trim();
            return compact.Length <= 500 ? compact : compact.Substring(0, 500) + "...";
        }

        public void Dispose()
        {
            httpClient.Dispose();
        }
    }
}
