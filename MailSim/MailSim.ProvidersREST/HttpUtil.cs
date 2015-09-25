using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Net.Http.Formatting;
using System.Dynamic;
using System.Linq;
using MailSim.Common;
using System.Text;

namespace MailSim.ProvidersREST
{
    static class HttpUtil
    {
        internal delegate string TokenFunc(bool isRefresh);

        private const string baseOutlookUri = @"https://outlook.office.com/api/v1.0/Me/";

        internal static async Task<T> GetItemAsync<T>(string uri, TokenFunc getToken)
        {
            return await DoHttp<EmptyBody,T>(HttpMethod.Get, uri, null, getToken);
        }

        internal static async Task<ODataCollection<T>> GetCollectionAsync<T>(string uri, TokenFunc getToken)
        {
            return await GetItemAsync<ODataCollection<T>>(uri, getToken);
        }

        internal static async Task<T> PostItemAsync<T>(string uri, T item, TokenFunc getToken)
        {
            return await DoHttp<T, T>(HttpMethod.Post, uri, item, getToken);
        }

        internal static async Task<T> PostItemDynamicAsync<T>(string uri, dynamic body, TokenFunc getToken)
        {
            return await DoHttp<ExpandoObject, T>(HttpMethod.Post, uri, body, getToken);
        }

        internal static async Task DeleteItemAsync(string uri, TokenFunc getToken)
        {
            await DoHttp<EmptyBody, EmptyBody>(HttpMethod.Delete, uri, null, getToken);
        }

        internal static async Task<T> PatchItemAsync<T>(string uri, T item, TokenFunc getToken)
        {
            return await DoHttp<T,T>("PATCH", uri, item, getToken);
        }

        private static async Task<TResult> DoHttp<TBody, TResult>(HttpMethod method, string uri, TBody body, TokenFunc getToken)
        {
            Log.Out(Log.Severity.Info, "DoHttp", string.Format("Uri=[{0}]", uri));

            for (bool isRefresh = false; ; isRefresh = true)
            {
                HttpResponseMessage response = await SendRequestAsync(method, uri, body, getToken, isRefresh);

                string jsonResponse = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    return JsonConvert.DeserializeObject<TResult>(jsonResponse);
                }

                var code = response.StatusCode;

                if (isRefresh == false && code == System.Net.HttpStatusCode.Unauthorized)
                {
                    var hasInvalidToken = response.Headers.WwwAuthenticate.FirstOrDefault(x => x.Parameter.Contains("error=\"invalid_token\"")) != null;

                    if (hasInvalidToken)
                    {
                        Log.Out(Log.Severity.Info, "DoHttp", "Found invalid_token!!!");
                        continue;
                    }
                }

                throw new System.Exception(GetErrorMessage(jsonResponse, code));
            }
        }

        private static string GetErrorMessage(string jsonResponse, System.Net.HttpStatusCode code)
        {
            string errorMessage;

            try
            {
                var errorDetail = string.IsNullOrEmpty(jsonResponse) ? null : JsonConvert.DeserializeObject<ODataError>(jsonResponse);
                errorMessage = errorDetail.error.message;
            }
            catch
            {
                errorMessage = string.Format("Error code: {0}; response = \"{1}\"", code, jsonResponse);
            }

            return errorMessage;
        }

        private static async Task<HttpResponseMessage> SendRequestAsync<TBody>(HttpMethod method, string uri, TBody body, TokenFunc getToken, bool isRefresh)
        {
            var request = new HttpRequestMessage(method, BuildUri(uri));

            if (body != null)
            {
                if (body is string)
                {
                    request.Content = new StringContent(body as string, Encoding.UTF8, "application/x-www-form-urlencoded");
                }
                else
                {
                    request.Content = new ObjectContent<TBody>(body, new JsonMediaTypeFormatter());
                }
            }

            string token = getToken(isRefresh);

            if (string.IsNullOrEmpty(token) == false)
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            using (HttpClient client = GetHttpClient())
            {
                return await client.SendAsync(request);
            }
        }

        internal static async Task<TResult> DoHttp<TBody, TResult>(string methodName, string uri, TBody body, TokenFunc getToken)
        {
            return await DoHttp<TBody,TResult>(new HttpMethod(methodName), uri, body, getToken);
        }

        private static string BuildUri(string subUri)
        {
            if (subUri.StartsWith("http"))
            {
                return subUri;
            }

            return baseOutlookUri + subUri;
        }

        private static HttpClient GetHttpClient()
        {
            return new HttpClient();
        }

        private class ODataError
        {
            public class Error
            {
                public string code { get; set; }
                public string message { get; set; }
            }
            public Error error { get; set; }
        }

        internal class ODataCollection<TCollection>
        {
            [Newtonsoft.Json.JsonProperty("@odata.nextLink")]
            public string NextLink { get; set; }

            public TCollection value { get; set; }
        }

        private class EmptyBody { }
    }
}
