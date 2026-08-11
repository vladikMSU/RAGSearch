using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RAGSearch
{
    internal sealed class LocalServiceClient : IDisposable
    {
        private const string TokenHeader = "X-RAGSearch-Token";
        private readonly HttpClient httpClient;
        private readonly string tokenPath;

        public Uri ServiceUri
        {
            get { return httpClient.BaseAddress; }
        }

        public LocalServiceClient()
        {
            var configuredUrl = Environment.GetEnvironmentVariable("RAGSEARCH_SERVICE_URL");
            var baseUrl = string.IsNullOrWhiteSpace(configuredUrl)
                ? "http://127.0.0.1:8765/"
                : configuredUrl.TrimEnd('/') + "/";

            Uri serviceUri;
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out serviceUri) ||
                !serviceUri.IsLoopback ||
                !string.Equals(serviceUri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                !string.IsNullOrEmpty(serviceUri.UserInfo) ||
                !string.Equals(serviceUri.AbsolutePath, "/", StringComparison.Ordinal) ||
                !string.IsNullOrEmpty(serviceUri.Query) ||
                !string.IsNullOrEmpty(serviceUri.Fragment))
            {
                throw new InvalidOperationException(
                    "RAGSEARCH_SERVICE_URL must be a plain HTTP loopback URL; mail data never leaves this computer.");
            }

            httpClient = new HttpClient
            {
                BaseAddress = serviceUri,
                Timeout = TimeSpan.FromMinutes(5)
            };
            tokenPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RAGSearch",
                "service-token");
        }

        public async Task<HealthResponse> GetHealthAsync(CancellationToken cancellationToken)
        {
            return await SendAsync<HealthResponse>(HttpMethod.Get, "health", null, cancellationToken)
                .ConfigureAwait(true);
        }

        public async Task<IngestResponse> IngestAsync(
            IList<MessagePayload> messages,
            CancellationToken cancellationToken)
        {
            var payload = new IngestRequest
            {
                messages = new List<MessagePayload>(messages)
            };
            return await SendAsync<IngestResponse>(
                    HttpMethod.Post,
                    "v1/messages",
                    payload,
                    cancellationToken)
                .ConfigureAwait(true);
        }

        public async Task<SearchResponse> SearchAsync(
            string query,
            int limit,
            CancellationToken cancellationToken)
        {
            return await SearchAsync(
                    query,
                    limit,
                    null,
                    cancellationToken)
                .ConfigureAwait(true);
        }

        public async Task<SearchResponse> SearchAsync(
            string query,
            int limit,
            IDictionary<string, object> filters,
            CancellationToken cancellationToken)
        {
            var request = new SearchRequest
            {
                query = query ?? string.Empty,
                limit = limit,
                filters = filters == null
                    ? new Dictionary<string, object>()
                    : new Dictionary<string, object>(filters)
            };
            return await SendAsync<SearchResponse>(
                    HttpMethod.Post,
                    "v1/search",
                    request,
                    cancellationToken)
                .ConfigureAwait(true);
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
            CancellationToken cancellationToken)
        {
            using (var request = new HttpRequestMessage(method, relativeUrl))
            {
                var token = ReadToken();
                if (!string.IsNullOrWhiteSpace(token))
                {
                    request.Headers.TryAddWithoutValidation(TokenHeader, token);
                }

                if (payload != null)
                {
                    var json = Serialize(payload);
                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                }

                using (var response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(true))
                {
                    var responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(true);
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
                        return default(T);
                    }

                    return Deserialize<T>(responseBody);
                }
            }
        }

        private string ReadToken()
        {
            try
            {
                return File.Exists(tokenPath) ? File.ReadAllText(tokenPath).Trim() : null;
            }
            catch (IOException)
            {
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return null;
            }
        }

        private static string Serialize(object payload)
        {
            var serializer = new DataContractJsonSerializer(
                payload.GetType(),
                CreateJsonSettings());
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, payload);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        private static T Deserialize<T>(string json)
        {
            var serializer = new DataContractJsonSerializer(
                typeof(T),
                CreateJsonSettings());
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(stream);
            }
        }

        private static DataContractJsonSerializerSettings CreateJsonSettings()
        {
            return new DataContractJsonSerializerSettings
            {
                MaxItemsInObjectGraph = int.MaxValue,
                UseSimpleDictionaryFormat = true
            };
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
