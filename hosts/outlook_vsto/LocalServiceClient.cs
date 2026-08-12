using System;
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
            return await SendAsync<HealthResponse>(HttpMethod.Get, "health", null, cancellationToken)
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
                request.Headers.TryAddWithoutValidation(TokenHeader, token);

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
                        throw new InvalidDataException(
                            "RAGSearch service returned an empty success response.");
                    }

                    return Deserialize<T>(responseBody);
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

        private static string Serialize(object payload)
        {
            var serializer = new DataContractJsonSerializer(payload.GetType());
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, payload);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        private static T Deserialize<T>(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(stream);
            }
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
